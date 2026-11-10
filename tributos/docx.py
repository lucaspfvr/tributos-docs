"""Fill WordprocessingML text while keeping the document's layout and assets."""

import base64
import io
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SPACE = "{http://www.w3.org/XML/1998/namespace}space"
KEY = re.compile(r"<([A-Za-z][A-Za-z0-9_]{0,63})>")
PART = re.compile(r"word/(?:document|header[0-9]+|footer[0-9]+|footnotes|endnotes)\.xml")
EXTERNAL_FIELD = re.compile(r"\b(?:DDE(?:AUTO)?|INCLUDE(?:TEXT|PICTURE)|LINK|DATABASE|RD)\b", re.I)

WORD_EXPORT = """
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
} catch {
    exit 10
}
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    # ConfirmConversions, ReadOnly, AddToRecentFiles
    $document = $word.Documents.Open($env:TRIBUTOS_WORD_SOURCE, $false, $true, $false)
    # wdExportFormatPDF
    $document.ExportAsFixedFormat($env:TRIBUTOS_WORD_PDF, 17)
} finally {
    try {
        if ($null -ne $document) { $document.Close(0) }
    } finally {
        $word.Quit(0)
    }
}
"""


def xml(data):
    from lxml import etree

    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def paragraphs(root):
    for paragraph in root.iter(W + "p"):
        yield paragraph


def segments(paragraph):
    """Group editable text between visible breaks within the current paragraph."""
    nodes = []
    for element in paragraph.iter(W + "t", W + "tab", W + "br", W + "cr", W + "drawing"):
        if next(element.iterancestors(W + "p"), None) is not paragraph:
            continue
        if element.tag == W + "t":
            if element.text:
                nodes.append(element)
        elif nodes:
            yield nodes
            nodes = []
    if nodes:
        yield nodes


def set_text(node, value):
    from lxml import etree

    lines = value.split("\n")
    node.set(SPACE, "preserve")
    node.text = lines[0]
    tail = node
    for line in lines[1:]:
        br = etree.Element(W + "br")
        tail.addnext(br)
        tail = etree.Element(W + "t", {SPACE: "preserve"})
        br.addnext(tail)
        tail.text = line


def replace_span(nodes, start, end, value, *, clear_red=False):
    # Text after the match stays in its original run and keeps its formatting.
    offset = 0
    for node in nodes:
        original = node.text or ""
        stop = offset + len(original)
        if start < stop and end > offset:
            prefix = original[:max(0, start - offset)]
            suffix = original[max(0, end - offset):]
            insertion = value if offset <= start < stop else ""
            set_text(node, prefix + insertion + suffix)
            if clear_red:
                color = node.getparent().find(f"{W}rPr/{W}color")
                if color is not None and color.get(W + "val", "").upper() == "FF0000":
                    color.getparent().remove(color)
        offset = stop
        if offset >= end:
            break


def read_package(data):
    try:
        with ZipFile(io.BytesIO(data)) as package:
            members = package.infolist()
            if len(members) > 5000 or sum(item.file_size for item in members) > 100_000_000:
                raise ValueError("O modelo excede o tamanho permitido.")
            entries = {item.filename: package.read(item) for item in members}
    except BadZipFile as exc:
        raise ValueError("Selecione um documento Word válido no formato .docx.") from exc
    if "word/document.xml" not in entries or "[Content_Types].xml" not in entries:
        raise ValueError("O arquivo não é um modelo Word .docx.")
    for name, data in entries.items():
        lower = name.lower()
        if lower.startswith(("word/activex/", "word/embeddings/")) or "vbaproject" in lower:
            raise ValueError("Use um modelo sem macros ou objetos incorporados.")
        if lower.endswith(".rels"):
            for rel in xml(data):
                if rel.get("TargetMode", "").lower() == "external" and not rel.get("Type", "").endswith("/hyperlink"):
                    raise ValueError("Remova as referências externas do modelo antes de importar.")
    return entries


def write_package(entries):
    output = io.BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as package:
        for name, content in entries.items():
            package.writestr(name, content)
    return output.getvalue()


def resolve_assessment_font(root, *, footer=False):
    from lxml import etree

    # The original assessment uses the East Asian font hint on ordinary
    # Latin text. Resolve that inherited face on the generated copy only.
    for paragraph in paragraphs(root):
        props = paragraph.find(W + "pPr")
        if props is None or props.find(W + "pStyle") is not None:
            continue
        fonts = props.find(f"{W}rPr/{W}rFonts")
        if not footer and (fonts is None or fonts.attrib != {W + "hint": "eastAsia"}):
            continue
        if footer and fonts is not None and set(fonts.attrib) - {W + "hint"}:
            continue
        for run in paragraph.iter(W + "r"):
            if next(run.iterancestors(W + "p"), None) is not paragraph or run.find(W + "t") is None:
                continue
            run_props = run.find(W + "rPr")
            if run_props is None:
                run_props = etree.Element(W + "rPr")
                run.insert(0, run_props)
            fonts = run_props.find(W + "rFonts")
            if run_props.find(W + "rStyle") is not None or (fonts is not None and
                    set(fonts.attrib) - {W + "hint"}):
                continue
            if fonts is None:
                fonts = etree.Element(W + "rFonts")
                run_props.insert(0, fonts)
            fonts.attrib.pop(W + "hint", None)
            for script in ("ascii", "hAnsi", "eastAsia", "cs"):
                fonts.set(W + script, "Times New Roman")


def replace_tags(data, values=None, *, assessment=False):
    from lxml import etree

    entries = read_package(data)
    if values is not None and assessment:
        with ZipFile(Path(__file__).with_name("assets") / "avaliacao.docx") as original:
            assessment = entries.get("word/styles.xml") == original.read("word/styles.xml")
    keys = {}
    for name, content in list(entries.items()):
        if not PART.fullmatch(name):
            continue
        root = xml(content)
        instructions = "".join(t.text or "" for t in root.iter(W + "instrText"))
        simple_fields = (t.get(W + "instr", "") for t in root.iter(W + "fldSimple"))
        if EXTERNAL_FIELD.search(instructions) or any(EXTERNAL_FIELD.search(x) for x in simple_fields):
            raise ValueError("Remova os campos do Word que acessam arquivos ou dados externos.")
        for paragraph in paragraphs(root):
            for nodes in segments(paragraph):
                matches = list(KEY.finditer("".join(t.text for t in nodes)))
                for match in matches:
                    keys.setdefault(match[1], None)
                if values is not None:
                    for match in reversed(matches):
                        if match[1] not in values:
                            raise ValueError(f"Campo sem valor no formulário: <{match[1]}>.")
                        replace_span(nodes, *match.span(), values[match[1]])
        if values is not None:
            if assessment and (name == "word/document.xml" or name.startswith("word/footer")):
                resolve_assessment_font(root, footer=name.startswith("word/footer"))
            entries[name] = etree.tostring(root)
    return list(keys) if values is None else write_package(entries)


def libreoffice():
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for base in ("C:/Program Files", "C:/Program Files (x86)"):
        for executable in ("soffice.com", "soffice.exe"):
            candidate = Path(base, "LibreOffice", "program", executable)
            if candidate.is_file():
                return str(candidate)
    return None


def export_with_word(source, pdf):
    powershell = Path(os.environ.get("SystemRoot", r"C:\Windows"),
                      "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    command = base64.b64encode(WORD_EXPORT.encode("utf-16-le")).decode("ascii")
    env = {**os.environ, "TRIBUTOS_WORD_SOURCE": str(source), "TRIBUTOS_WORD_PDF": str(pdf)}
    try:
        subprocess.run(
            [str(powershell), "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", command],
            env=env, capture_output=True, timeout=90, check=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except OSError as exc:
        raise ValueError("Não foi possível iniciar o PowerShell para usar o Word.") from exc
    except subprocess.TimeoutExpired as exc:
        raise ValueError("O Word demorou demais para gerar o PDF. Abra o Word e verifique se há uma janela pendente.") from exc
    except subprocess.CalledProcessError as exc:
        if exc.returncode == 10:
            raise ValueError("Não foi possível iniciar o Word instalado no Windows.") from exc
        raise ValueError("O Word não conseguiu exportar este documento. Abra o Word e verifique se ele está pronto para uso.") from exc
    if not pdf.is_file():
        raise ValueError("O Word terminou sem gerar o PDF.")


def export_with_libreoffice(executable, source, pdf):
    folder = source.parent
    args = [executable, "-env:UserInstallation=" + (folder / "perfil").as_uri(),
            "--headless", "--nologo", "--nodefault", "--norestore", "--convert-to",
            "pdf:writer_pdf_Export", "--outdir", str(folder), str(source)]
    try:
        subprocess.run(args, capture_output=True, timeout=90, check=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired as exc:
        raise ValueError("A conversão para PDF excedeu o tempo de espera.") from exc
    except (OSError, subprocess.CalledProcessError) as exc:
        raise ValueError("O LibreOffice não conseguiu converter o modelo em PDF.") from exc
    if not pdf.is_file():
        raise ValueError("O LibreOffice terminou sem gerar o PDF. Revise o modelo no Writer.")


def export_pdf(document):
    with TemporaryDirectory(prefix="conversao-", ignore_cleanup_errors=True) as folder:
        path = Path(folder)
        source = path / "Documento.docx"
        pdf = path / "Documento.pdf"
        source.write_bytes(document)
        word_error = None
        if sys.platform == "win32":
            try:
                export_with_word(source, pdf)
                return pdf.read_bytes()
            except ValueError as exc:
                word_error = exc
        executable = libreoffice()
        if not executable:
            if word_error:
                raise word_error
            raise ValueError("Instale o LibreOffice Writer para gerar o PDF neste sistema.")
        pdf.unlink(missing_ok=True)
        export_with_libreoffice(executable, source, pdf)
        return pdf.read_bytes()

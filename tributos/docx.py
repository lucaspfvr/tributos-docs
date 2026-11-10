"""Fill WordprocessingML text while keeping the document's layout and assets."""

import base64
import ctypes
import io
import os
import re
import shutil
import subprocess
import sys
from contextlib import closing
from copy import deepcopy
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


def replace_rich_span(nodes, start, end, pieces):
    from lxml import etree

    offset = 0
    for node in nodes:
        stop = offset + len(node.text)
        if start < stop:
            break
        offset = stop
    original_run = node.getparent()
    prefix = start - offset
    replace_span(nodes, start, end, "")
    remainder = (node.text or "")[prefix:]
    set_text(node, (node.text or "")[:prefix])
    properties = original_run.find(W + "rPr")
    tail = original_run
    for value, bold in [*pieces, (remainder, False)]:
        if not value:
            continue
        run = etree.Element(W + "r")
        if properties is not None:
            run.append(deepcopy(properties))
        if bold:
            props = run.find(W + "rPr")
            if props is None:
                props = etree.SubElement(run, W + "rPr")
            flag = props.find(W + "b")
            if flag is None:
                flag = etree.SubElement(props, W + "b")
            flag.set(W + "val", "1")
        set_text(etree.SubElement(run, W + "t"), value)
        tail.addnext(run)
        tail = run


def replace_tags(data, values=None):
    from lxml import etree

    entries = read_package(data)
    keys = {}
    for name, content in list(entries.items()):
        if not PART.fullmatch(name):
            continue
        root = xml(content)
        instructions = "".join(t.text or "" for t in root.iter(W + "instrText"))
        simple_fields = (t.get(W + "instr", "") for t in root.iter(W + "fldSimple"))
        if EXTERNAL_FIELD.search(instructions) or any(EXTERNAL_FIELD.search(x) for x in simple_fields):
            raise ValueError("Remova os campos do Word que acessam arquivos ou dados externos.")
        changed = False
        for paragraph in paragraphs(root):
            for nodes in segments(paragraph):
                matches = list(KEY.finditer("".join(t.text for t in nodes)))
                for match in matches:
                    keys.setdefault(match[1], None)
                if values is not None:
                    for match in reversed(matches):
                        if match[1] not in values:
                            raise ValueError(f"Campo sem valor no formulário: <{match[1]}>.")
                        value = values[match[1]]
                        if match[1] == "FracoesEmLinha" and isinstance(value, list):
                            replace_rich_span(nodes, *match.span(), value)
                        else:
                            replace_span(nodes, *match.span(), value)
                        changed = True
        if changed:
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


HEADING = "MODELO DEMONSTRATIVO"
BERNARD = "Bernard MT Condensed"


def layout_font(document, kind):
    if kind not in {"desdobro", "retificacao", "avaliacao"}:
        return None
    with ZipFile(io.BytesIO(document)) as package:
        if kind == "avaliacao":
            if "word/styles.xml" in package.namelist():
                styles = xml(package.read("word/styles.xml"))
                normal = styles.find(f'{W}style[@{W}default="1"][@{W}type="paragraph"]')
                fonts = normal.find(f"{W}rPr/{W}rFonts") if normal is not None else None
                if fonts is not None and fonts.get(W + "ascii") == "Aptos":
                    for paragraph in paragraphs(xml(package.read("word/document.xml"))):
                        title = "".join(t.text or "" for t in paragraph.iter(W + "t"))
                        if title == "AVALIAÇÃO MUNICIPAL DE IMÓVEL":
                            return "Aptos"
        else:
            for name in package.namelist():
                if re.fullmatch(r"word/header[0-9]+\.xml", name):
                    for run in xml(package.read(name)).iter(W + "r"):
                        fonts = run.find(f"{W}rPr/{W}rFonts")
                        if fonts is not None and fonts.get(W + "ascii") == BERNARD:
                            if HEADING in "".join(t.text or "" for t in run.iter(W + "t")):
                                return BERNARD
    return None


def verify_pdf_font(data, required):
    if required is None:
        return
    import pypdfium2 as pdfium
    from pypdfium2 import raw

    def font_at(text, index):
        buffer = ctypes.create_string_buffer(256)
        raw.FPDFText_GetFontInfo(text, index, buffer, len(buffer), None)
        return re.sub(r"[^a-z]", "", buffer.value.decode("ascii", errors="ignore").lower())

    name = re.sub(r"[^a-z]", "", required.lower())
    with pdfium.PdfDocument(data) as pdf, closing(pdf[0]) as page, closing(page.get_textpage()) as text:
        characters = [chr(raw.FPDFText_GetUnicode(text, i)) for i in range(text.count_chars())]
        if required == BERNARD:
            match = re.search(HEADING.replace(" ", r"\s+"), "".join(characters))
            if match:
                letters = [i for i in range(*match.span()) if characters[i].isalpha()]
                baselines = [text.get_charbox(i)[1] for i in letters]
                if max(baselines) - min(baselines) < 5 and all(name in font_at(text, i) for i in letters):
                    return
        elif any(name in font_at(text, i) for i, c in enumerate(characters) if c.isalpha()):
            return
    if required == BERNARD:
        raise ValueError("O PDF não preservou o cabeçalho em Bernard MT Condensed. Abra o modelo no Word, "
                         "confirme que a fonte Bernard MT Condensed está disponível e tente novamente.")
    raise ValueError("O PDF não utilizou a fonte Aptos. Instale a fonte pelo instalador do Tributos Docs "
                     "e reabra o Word ou o Writer.")


def export_pdf(document, *, required_font=None):
    with TemporaryDirectory(prefix="conversao-", ignore_cleanup_errors=True) as folder:
        path = Path(folder)
        source = path / "Documento.docx"
        pdf = path / "Documento.pdf"
        source.write_bytes(document)
        word_error = None
        if sys.platform == "win32":
            try:
                export_with_word(source, pdf)
                data = pdf.read_bytes()
                verify_pdf_font(data, required_font)
                return data
            except ValueError as exc:
                word_error = exc
        executable = libreoffice()
        if not executable:
            if word_error:
                raise word_error
            raise ValueError("Instale o LibreOffice Writer para gerar o PDF neste sistema.")
        pdf.unlink(missing_ok=True)
        export_with_libreoffice(executable, source, pdf)
        data = pdf.read_bytes()
        verify_pdf_font(data, required_font)
        return data

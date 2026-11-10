import io
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import pypdfium2 as pdfium
from lxml import etree

from tributos.docx import W, export_pdf, replace_tags, write_package, xml
from tributos.models import Model, defaults, format_tax_id, prepare
from tributos.storage import Storage


def pdf_text(data):
    with pdfium.PdfDocument(data) as pdf:
        lines = []
        for page in pdf:
            try:
                text = page.get_textpage()
                try:
                    lines.append(text.get_text_bounded())
                finally:
                    text.close()
            finally:
                page.close()
        return len(pdf), " ".join(" ".join(lines).split())


class ProjectTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.store = Storage(self.root)
        self.addCleanup(self.store.connection.close)
        self.models = {model.kind: model for model in self.store.models()}
        self.fields = dict(defaults(), NomeProprietario="PESSOA DE TESTE",
                           CpfCnpj="000.000.000-00", Matricula="12345", Livro="4",
                           CimImovel="90001", Quadra="21", Lote="07",
                           AreaMensurada="381,68", LogradouroImovel="Rua de Teste",
                           NumeroImovel="S/N", NumeroDocumento="TESTE-02",
                           ValorAvaliacao="60.000,00", ValorExtenso="Sessenta mil reais.",
                           AreaTerreno="999,14")
        self.fractions = [
            {"lote": f"07-{letter}", "area": area, "endereco": "Rua de Teste", "numero": str(i)}
            for i, (letter, area) in enumerate(zip("ABCD", ("150", "130,31", "141,25", "778,44")), 1)
        ]

    @unittest.skipUnless(shutil.which("soffice") or shutil.which("libreoffice"), "Requer LibreOffice no PATH")
    def test_three_real_models_and_pdf_layout(self):
        for kind, model in self.models.items():
            with self.subTest(kind=kind):
                document = prepare(model, self.fields, self.fractions)
                with ZipFile(io.BytesIO(document)) as filled, ZipFile(io.BytesIO(model.document)) as original:
                    image = "word/media/image1.png"
                    self.assertEqual(filled.read(image), original.read(image))
                count, text = pdf_text(export_pdf(document))
                self.assertEqual(count, 1, text)
                self.assertIn("PESSOA DE TESTE", text)
                if kind == "desdobro":
                    self.assertIn("04 frações", text)
                    self.assertIn("Lote 07-D", text)
                    self.assertIn("778,44 m²", text)
                    self.assertIn("nº 4", text)
                elif kind == "retificacao":
                    self.assertIn("381,68", text)
                    self.assertIn("90001", text)
                else:
                    self.assertIn("60.000,00", text)
                    self.assertIn("999,14", text)
                    self.assertIn("TESTE-02", text)

    def test_tax_id_mask_and_uppercase_owners_in_generated_document(self):
        self.assertEqual(format_tax_id("1234567", partial=True), "123.456.7")
        self.assertEqual(format_tax_id("123.456.789-012", partial=True), "12.345.678/9012")
        self.assertEqual(format_tax_id("12345678901"), "123.456.789-01")
        self.assertEqual(format_tax_id("12345678000199"), "12.345.678/0001-99")
        self.assertEqual(format_tax_id("12.ABC.345/0001-99"), "12.ABC.345/0001-99")
        fields = dict(self.fields, NomeProprietario="Lúcia da Silva", CpfCnpj="12345678901",
                      NomeProprietario2="José D’Ávila", CpfCnpj2="12345678000199")
        for kind, model in self.models.items():
            with self.subTest(kind=kind):
                filled = prepare(model, fields, self.fractions)
                with ZipFile(io.BytesIO(filled)) as package:
                    text = "".join(xml(package.read("word/document.xml")).itertext())
                self.assertIn("LÚCIA DA SILVA", text)
                self.assertIn("123.456.789-01", text)
                if kind == "desdobro":
                    self.assertIn("JOSÉ D’ÁVILA", text)
                    self.assertIn("12.345.678/0001-99", text)

    def test_assessment_derives_currency_and_preserves_template_styles(self):
        model = self.models["avaliacao"]
        self.assertNotIn("ValorExtenso", model.groups()["Avaliação"])
        fields = dict(self.fields, ValorAvaliacao="450000", ValorExtenso="Texto antigo")
        filled = prepare(model, fields, [])
        with ZipFile(io.BytesIO(filled)) as output, ZipFile(io.BytesIO(model.document)) as original:
            self.assertEqual(output.read("word/styles.xml"), original.read("word/styles.xml"))
            self.assertEqual(output.read("word/theme/theme1.xml"), original.read("word/theme/theme1.xml"))
            self.assertEqual(output.read("word/media/image1.png"), original.read("word/media/image1.png"))
            root = xml(output.read("word/document.xml"))
            self.assertIn("Quatrocentos e cinquenta mil reais.", "".join(root.itertext()))
            title = next(t for t in root.iter(W + "t") if t.text == "AVALIAÇÃO MUNICIPAL DE IMÓVEL")
            original_title = next(t for t in xml(original.read("word/document.xml")).iter(W + "t")
                                  if t.text == title.text)
            self.assertEqual(etree.tostring(title.getparent()), etree.tostring(original_title.getparent()))
            styles = xml(output.read("word/styles.xml"))
            normal = styles.find(f'{W}style[@{W}styleId="Normal"]')
            self.assertEqual(normal.find(f"{W}rPr/{W}rFonts").get(W + "ascii"), "Aptos")
        filled = prepare(model, dict(fields, ValorAvaliacao="1.000.000,01"), [])
        with ZipFile(io.BytesIO(filled)) as output:
            self.assertIn("Um milhão de reais e um centavo.", "".join(xml(output.read("word/document.xml")).itertext()))

    def test_desdobro_uses_original_header_and_one_line_per_fraction(self):
        model = self.models["desdobro"]
        fractions = [dict(self.fractions[0], numero=""), *self.fractions[1:]]
        document = prepare(model, self.fields, fractions)
        with ZipFile(io.BytesIO(document)) as filled, ZipFile(io.BytesIO(model.document)) as original:
            header = filled.read("word/header1.xml")
            self.assertEqual(header, original.read("word/header1.xml"))
            self.assertTrue(any(font.get(W + "ascii") == "Bernard MT Condensed"
                                for font in xml(header).iter(W + "rFonts")))
            body = xml(filled.read("word/document.xml"))
            lots = next(p for p in body.iter(W + "p") if "Lote 07-A" in "".join(p.itertext()))
            self.assertNotIn("04 frações", "".join(lots.itertext()))
            lines = [t.text for t in lots.iter(W + "t")]
            self.assertEqual(len(list(lots.iter(W + "br"))), 3)
            self.assertEqual([line.split(",")[0] for line in lines], [f"Lote 07-{x}" for x in "ABCD"])
            self.assertNotIn("nº", lines[0])

    def test_new_key_in_split_word_runs_and_history(self):
        model = self.models["desdobro"]
        with ZipFile(io.BytesIO(model.document)) as package:
            entries = {part: package.read(part) for part in package.namelist()}
        root = xml(entries["word/document.xml"])
        body = root.find(W + "body")
        paragraph = etree.Element(W + "p")
        body.insert(1, paragraph)
        for part in ("Texto <Obser", "vacaoLocal> e <CimImovel>."):
            etree.SubElement(etree.SubElement(paragraph, W + "r"), W + "t").text = part
        entries["word/document.xml"] = etree.tostring(root)
        edited = Model(model.id, model.title, model.kind, write_package(entries))
        self.store.save_model(edited)
        self.assertIn("ObservacaoLocal", edited.groups()["Campos adicionais"])
        fields = dict(self.fields, NomeProprietario="PESSOA FICTÍCIA", ObservacaoLocal="Exemplo & teste")
        filled = prepare(edited, fields, self.fractions)
        self.assertEqual(replace_tags(filled), [])
        with ZipFile(io.BytesIO(filled)) as package:
            text = "".join(xml(package.read("word/document.xml")).itertext())
            self.assertIn("Texto Exemplo & teste e 90001.", text)
        rid, v1, pdf1 = self.store.save_record(edited, fields, self.fractions, b"test pdf 1", "25/2026")
        _, v2, pdf2 = self.store.save_record(edited, fields, self.fractions, b"test pdf 2", "26/2026", rid)
        self.assertEqual((v1, v2), (1, 2))
        self.assertEqual(pdf1.read_bytes(), b"test pdf 1")
        self.assertEqual(pdf2.read_bytes(), b"test pdf 2")
        self.assertEqual(len(self.store.history("Ficticia")), 2)
        self.assertEqual(len(self.store.history("25/2026")), 1)
        self.assertEqual(len(self.store.history(f"TD-{rid:06d}")), 2)
        self.store.delete_model(edited.id)
        self.assertEqual(json.loads(self.store.record(rid, v1)["form"])["ObservacaoLocal"], "Exemplo & teste")
        backup = self.root / "backup.sqlite3"
        self.store.backup(backup)
        other = Storage(self.root / "outra-estacao")
        self.addCleanup(other.connection.close)
        other.restore(backup)
        row = other.record(rid, v2)
        self.assertEqual(other.ensure_pdf(row["file"], row["pdf"]).read_bytes(), b"test pdf 2")

    def test_history_follows_files_and_retains_legacy_metadata(self):
        model = self.models["desdobro"]
        record_id, version, pdf = self.store.save_record(
            model, self.fields, self.fractions, b"%PDF-old", "REF-100")
        self.assertEqual(pdf.parent.name, "DESDOBRO")

        # Simulate an installation with the prior database and disk layout.
        pdf.unlink()
        record = self.store.record(record_id, version)
        old = self.root / "Documentos" / record["file"]
        old.parent.mkdir(parents=True)
        old.write_bytes(b"%PDF-old")
        with self.store.connection:
            self.store.connection.execute("DELETE FROM settings")
        migrated, = self.store.history("REF-100")
        self.assertEqual(migrated["path"], pdf)
        self.assertEqual(pdf.read_bytes(), b"%PDF-old")
        self.assertTrue(old.exists())

        manual = self.store.archive / "RETIFICAÇÃO" / "Retificação da Praça.PDF"
        manual.write_bytes(b"%PDF-manual")
        match, = self.store.history("retificacao da praca")
        self.assertEqual(match["path"], manual)
        self.assertIsNone(match["record_id"])
        pdf.unlink()
        self.assertEqual(self.store.history("REF-100"), [])
        self.assertFalse(pdf.exists())

    def test_pdf_names_use_reference_owner_and_preserve_existing_files(self):
        model = self.models["desdobro"]
        fields = dict(self.fields, NomeProprietario="Pessoa de Teste")
        rid, _, first = self.store.save_record(model, fields, self.fractions, b"PDF original", "001")
        _, _, second = self.store.save_record(model, fields, self.fractions, b"PDF revisado", "001", rid)
        self.assertEqual(first.name, "DESDOBRO 001 PESSOA DE TESTE.pdf")
        self.assertEqual(second.name, "DESDOBRO 001 PESSOA DE TESTE (2).pdf")
        self.assertEqual(first.read_bytes(), b"PDF original")
        self.assertEqual(second.read_bytes(), b"PDF revisado")
        self.assertEqual(len(self.store.history("001 Pessoa de Teste")), 2)
        _, _, sanitized = self.store.save_record(self.models["retificacao"],
            fields, [], b"PDF", "001/2026")
        self.assertEqual(sanitized.name, "RETIFICAÇÃO 001-2026 PESSOA DE TESTE.pdf")

    def test_independent_sequences_follow_archived_numbers(self):
        split = self.models["desdobro"]
        rectification = self.models["retificacao"]
        assessment = self.models["avaliacao"]
        self.assertEqual([self.store.next_reference(m) for m in self.models.values()], ["001"] * 3)
        rid, _, first = self.store.save_record(split, self.fields, self.fractions, b"%PDF", "001", automatic=True)
        self.assertTrue(first.name.startswith("DESDOBRO 001 "))
        self.store.save_record(split, self.fields, self.fractions, b"%PDF", "001", rid)
        self.assertEqual(self.store.next_reference(split), "002")
        self.assertEqual(self.store.next_reference(rectification), "001")
        self.assertEqual(self.store.next_reference(assessment), "001")
        self.store.save_record(rectification, self.fields, [], b"%PDF", "006/2026")
        self.assertEqual(self.store.next_reference(rectification), "007")
        self.store.save_record(assessment, self.fields, [], b"%PDF", "")
        self.assertEqual(self.store.next_reference(assessment), "004")
        self.assertEqual(self.store.next_reference(split), "002")

    def test_stale_automatic_number_and_failed_write_do_not_consume_numbers(self):
        model = self.models["desdobro"]
        other = Storage(self.root)
        self.addCleanup(other.connection.close)
        stale = other.next_reference(model)
        self.store.save_record(model, self.fields, self.fractions, b"%PDF", stale, automatic=True)
        with self.assertRaisesRegex(ValueError, "Outro documento"):
            other.save_record(model, self.fields, self.fractions, b"%PDF", stale, automatic=True)
        with patch("tributos.storage.atomic_write", side_effect=OSError("falha de teste")):
            with self.assertRaises(OSError):
                other.save_record(model, self.fields, self.fractions, b"%PDF", "002", automatic=True)
        self.assertEqual(other.next_reference(model), "002")
        self.assertEqual(len(other.history()), 1)
        self.assertEqual(other.save_record(model, self.fields, self.fractions, b"%PDF",
                                           "002", automatic=True)[:2], (2, 1))

    def test_desdobro_validation_and_one_fraction(self):
        model = self.models["desdobro"]
        with self.assertRaisesRegex(ValueError, "Fração 1"):
            prepare(model, self.fields, [{}])
        with self.assertRaisesRegex(ValueError, "segundo proprietário"):
            prepare(model, dict(self.fields, NomeProprietario2="EXEMPLO"), self.fractions)
        one = prepare(model, self.fields, self.fractions[:1])
        self.assertEqual(replace_tags(one), [])

import unittest
from unittest.mock import patch

from tributos import docx


class ExportTest(unittest.TestCase):
    def test_uses_word_when_available_on_windows(self):
        def word(source, pdf):
            self.assertEqual(source.read_bytes(), b"docx")
            pdf.write_bytes(b"%PDF-word")

        with patch.object(docx.sys, "platform", "win32"), \
             patch.object(docx, "export_with_word", side_effect=word), \
             patch.object(docx, "libreoffice") as libreoffice:
            self.assertEqual(docx.export_pdf(b"docx"), b"%PDF-word")
            libreoffice.assert_not_called()

    def test_falls_back_to_libreoffice_when_word_fails(self):
        def libreoffice(executable, source, pdf):
            self.assertEqual(executable, "soffice")
            self.assertEqual(source.read_bytes(), b"docx")
            self.assertFalse(pdf.exists())
            pdf.write_bytes(b"%PDF-libreoffice")

        with patch.object(docx.sys, "platform", "win32"), \
             patch.object(docx, "export_with_word", side_effect=ValueError("Word indisponível")), \
             patch.object(docx, "libreoffice", return_value="soffice"), \
             patch.object(docx, "export_with_libreoffice", side_effect=libreoffice):
            self.assertEqual(docx.export_pdf(b"docx"), b"%PDF-libreoffice")

    def test_reports_word_failure_when_no_fallback_is_available(self):
        with patch.object(docx.sys, "platform", "win32"), \
             patch.object(docx, "export_with_word", side_effect=ValueError("Word indisponível")), \
             patch.object(docx, "libreoffice", return_value=None):
            with self.assertRaisesRegex(ValueError, "Word indisponível"):
                docx.export_pdf(b"docx")

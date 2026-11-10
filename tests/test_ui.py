import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from tkinter import ttk
from unittest.mock import patch
from types import SimpleNamespace

from tributos.storage import Storage
from tributos.ui import App, review_details


def descendants(widget):
    yield widget
    for child in widget.winfo_children():
        yield from descendants(child)


class DetailsTest(unittest.TestCase):
    def test_each_type_has_its_own_summary(self):
        form = dict(NomeProprietario="EXEMPLO", Quadra="21", Lote="07",
                    Matricula="124", LogradouroImovel="Avenida Central", NumeroImovel="6",
                    AreaTerreno="500", AreaMensurada="450", AreaConstruida="100", ValorAvaliacao="450000")
        fractions = [dict(lote="07-A", area="150", endereco="Rua do Teste", numero="8"),
                     dict(lote="07-B", area="150,50", endereco="Rua do Teste", numero="10")]
        summary = dict(review_details("desdobro", form, fractions, "001"))
        self.assertEqual(summary["Lote de origem"], "Quadra 21 · Lote 07")
        self.assertEqual(summary["Lote 07-B"], "Rua do Teste, 10\n150,50 m²")
        self.assertEqual(summary["Soma das áreas dos lotes"], "300,50 m²")
        self.assertNotIn("Valor da avaliação", summary)
        summary = dict(review_details("retificacao", form, [], "001"))
        self.assertEqual(summary["Área mensurada"], "450,00 m²")
        self.assertNotIn("Frações resultantes", summary)
        summary = dict(review_details("avaliacao", form, [], "001"))
        self.assertEqual(summary["Valor da avaliação"], "R$ 450.000,00")
        self.assertEqual(summary["Área construída"], "100,00 m²")
        self.assertNotIn("Área mensurada", summary)


class NavigationTest(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Requer sessão gráfica: {exc}")
        self.addCleanup(self.root.destroy)
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.store = Storage(Path(folder.name))
        self.app = App(self.root, self.store)
        self.addCleanup(self.app.executor.shutdown)
        self.addCleanup(self.store.connection.close)
        self.root.geometry("820x550")
        self.root.update()

    def wait_animation(self):
        end = time.monotonic() + .35
        while self.app.sidebar_job is not None and time.monotonic() < end:
            self.root.update()
            time.sleep(.01)
        self.root.update()

    def test_menu_only_toggles_from_its_button_and_persists(self):
        app = self.app
        self.assertFalse(app.sidebar_expanded)
        self.root.geometry("640x420")
        self.root.update()
        self.assertTrue(app.menu_toggle.winfo_viewable())
        self.assertLess(app.menu_toggle.winfo_y(), app.nav_buttons["Início"].winfo_y())
        self.assertGreaterEqual(app.menu_toggle.winfo_width(), 40)
        self.assertIn(str(app.controls.icons["expand", False]), app.menu_toggle.cget("image"))
        app.menu_toggle.invoke()
        self.assertEqual(int(app.sidebar.cget("width")), 56)
        self.wait_animation()
        self.assertEqual(app.sidebar.winfo_width(), 180)
        self.assertEqual(app.menu_toggle.cget("text"), "Recolher menu")
        for page in ("Desdobro", "Configurações", "Histórico", "Início"):
            app.nav_buttons[page].invoke()
            self.root.update()
            self.assertTrue(app.sidebar_expanded)
            self.assertEqual(app.nav_buttons[page].cget("text"), page)
        app.menu_toggle.invoke()
        self.wait_animation()
        self.assertEqual(app.sidebar.winfo_width(), 56)
        for page in ("Avaliação", "Retificação", "Histórico", "Início"):
            app.nav_buttons[page].invoke()
            self.root.update()
            self.assertFalse(app.sidebar_expanded)
            self.assertTrue(app.menu_toggle.winfo_viewable())
            self.assertEqual(app.nav_buttons[page].cget("text"), "")
        app.menu_toggle.invoke()
        app.menu_toggle.invoke()
        self.wait_animation()
        self.assertEqual(app.sidebar.winfo_width(), 56)
        self.assertIn('"sidebar_expanded": false', app.settings_path.read_text())

    def test_scroll_stays_at_top_when_content_fits(self):
        area = self.app.scroll
        wheel = SimpleNamespace(delta=-120, num=0)
        self.assertLess(area.body.winfo_reqheight(), area.canvas.winfo_height())
        for _ in range(20):
            area.wheel(wheel)
        self.assertEqual(area.canvas.yview(), (0, 1))
        self.assertEqual(area.canvas.canvasy(0), 0)
        self.app.nav_buttons["Avaliação"].invoke()
        self.root.geometry("640x420")
        self.root.update()
        area = self.app.scroll
        area.wheel(wheel)
        self.assertGreater(area.canvas.yview()[0], 0)
        area.wheel(SimpleNamespace(delta=120, num=0))
        self.assertEqual(area.canvas.canvasy(0), 0)
        self.assertTrue(self.app.generate_button.winfo_viewable())

    def test_history_opens_side_details_and_can_edit_saved_record(self):
        model = next(model for model in self.store.models() if model.kind == "desdobro")
        rid, _, _ = self.store.save_record(model, {"NomeProprietario": "PESSOA TESTE", "Lote": "02"}, [
            {"lote": "02-A", "area": "180", "endereco": "Rua do Teste", "numero": "6"}], b"%PDF", "012")
        self.app.history()
        self.root.geometry("640x420")
        self.root.update()
        table = next(w for w in descendants(self.app.main) if isinstance(w, ttk.Treeview))
        table.selection_set(table.get_children()[0])
        self.root.update()
        def button(caption):
            return next(w for w in descendants(self.app.main)
                        if isinstance(w, ttk.Button) and w.cget("text") == caption)
        self.assertTrue(button("Detalhes").winfo_viewable())
        button("Detalhes").invoke()
        self.root.update()
        self.assertTrue(self.app.history_scroll.winfo_viewable())
        captions = [w.cget("text") for w in descendants(self.app.history_scroll) if isinstance(w, tk.Label)]
        self.assertIn("Lote 02-A", captions)
        self.assertIn("Lote 02", captions)
        button("Editar").invoke()
        self.root.update()
        self.assertEqual(self.app.form_vars["NomeProprietario"].get(), "PESSOA TESTE")
        self.assertEqual(self.app.document_number.get(), "012")
        self.assertIsNotNone(self.store.record(rid, 1))


    def test_automatic_reference_reaches_word_preview_history_and_filename(self):
        app = self.app
        app.nav_buttons["Avaliação"].invoke()
        self.assertEqual(app.document_number.get(), "001")
        self.assertIs(app.document_number, app.form_vars["NumeroDocumento"])
        for key, value in {"NomeProprietario": "PESSOA TESTE", "ValorAvaliacao": "450000"}.items():
            app.form_vars[key].set(value)

        class SavePreview(tk.Toplevel):
            def __init__(self, parent, data, title, destination, suggested, details, save, controls):
                super().__init__(parent)
                self.saved = save()
                self.after_idle(self.destroy)

        errors = []
        with patch("tributos.ui.Preview", SavePreview), \
             patch("tributos.ui.export_pdf", return_value=b"%PDF-ficticio") as export, \
             patch.object(app, "error", side_effect=errors.append):
            app.generate_button.invoke()
            end = time.monotonic() + 2
            while app.busy and time.monotonic() < end:
                self.root.update()
                time.sleep(.01)
        self.assertFalse(app.busy)
        self.assertEqual(errors, [])
        export.assert_called_once()
        record, = app.store.history()
        self.assertEqual(record["reference"], "001")
        self.assertEqual(record["path"].name, "AVALIAÇÃO 001 PESSOA TESTE.pdf")
        self.assertTrue(any(isinstance(w, tk.Label) and w.cget("text") == "PDF arquivado"
                            for w in descendants(app.main)))
        app.nav_buttons["Avaliação"].invoke()
        self.assertEqual(app.document_number.get(), "002")
        app.nav_buttons["Desdobro"].invoke()
        self.assertEqual(app.document_number.get(), "001")

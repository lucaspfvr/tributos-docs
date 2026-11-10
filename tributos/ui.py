"""Offline desktop interface for generating and managing municipal documents."""

import json
import os
import subprocess
import sys
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import filedialog, messagebox, ttk
from uuid import uuid4

from .docx import export_pdf
from .models import FRACTION_FIELDS, MAX_FRACTIONS, REQUIRED, Model, defaults, prepare
from .storage import Storage, matches_history

NAV = "#112836"
NAV_ACTIVE = "#214552"
BG = "#f1f5f6"
PAPER = "#ffffff"
INK = "#182e38"
MUTED = "#4d656d"
LINE = "#d5e2e5"
ACCENT = "#076b5c"
ACCENT_BG = "#e0f2ec"
FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"


def label(parent, text, size=10, *, bold=False, fg=INK, bg=PAPER, anchor="w", **kwargs):
    return tk.Label(parent, text=text, fg=fg, bg=bg, anchor=anchor,
                    font=(FONT, size, "bold" if bold else "normal"), **kwargs)


def open_path(path):
    if sys.platform == "win32":
        os.startfile(str(path))
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(path)])


def data_directory():
    custom = os.environ.get("TRIBUTOS_DOCS_DATA")
    if custom:
        return Path(custom).expanduser()
    if sys.platform == "win32":
        return Path(os.environ["LOCALAPPDATA"]) / "Tributos Docs"
    return Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share")) / "tributos-docs"


def archive_directory():
    custom = os.environ.get("TRIBUTOS_DOCS_ARCHIVE")
    if custom:
        return Path(custom).expanduser()
    folder = "Documents" if sys.platform == "win32" else "Documentos"
    return Path.home() / folder / "Tributos Docs"


class ScrollArea(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.body = tk.Frame(self.canvas, bg=BG)
        window = self.canvas.create_window(0, 0, anchor="nw", window=self.body)
        self.body.bind("<Configure>", lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(window, width=e.width))
        self.canvas.configure(yscrollcommand=bar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")

    def wheel(self, event):
        units = -int(event.delta / 120) if event.delta else (-1 if event.num == 4 else 1)
        self.canvas.yview_scroll(units * 3, "units")


class Preview(tk.Toplevel):
    def __init__(self, parent, data, title, destination, save):
        import pypdfium2
        from PIL import ImageTk

        self.pdf = pypdfium2.PdfDocument(data)
        self.image_tk = ImageTk
        self.index = 0
        self.saved = None
        self.save_document = save
        self.document_title = title
        super().__init__(parent)
        self.title(f"{title} · Prévia")
        width, height = self.maxsize()
        self.geometry(f"{min(900, width - 25)}x{min(710, height - 55)}+30+25")
        self.minsize(550, 320)
        self.transient(parent)

        header = tk.Frame(self, bg=PAPER, padx=16, pady=10)
        header.pack(fill="x")
        label(header, "REVISAR DOCUMENTO", 9, bold=True, fg=ACCENT).pack(anchor="w")
        label(header, title, 15, bold=True).pack(anchor="w", pady=(2, 0))
        toolbar = tk.Frame(self, bg=PAPER, padx=16, pady=8)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Voltar", command=self.close).pack(side="left")
        ttk.Button(toolbar, text="Salvar PDF", style="Accent.TButton", command=self.confirm).pack(side="right")
        ttk.Button(toolbar, text="Imprimir", command=self.print_document).pack(side="right", padx=8)
        label(self, f"Destino do PDF  ·  {destination}", 9, fg=MUTED, padx=16, pady=7).pack(fill="x")
        pages = tk.Frame(self, bg=PAPER, padx=16, pady=5)
        pages.pack(fill="x")
        self.previous = ttk.Button(pages, text="‹", width=4, command=lambda: self.show_page(-1))
        self.previous.pack(side="left")
        self.counter = label(pages, "", fg=MUTED, width=12, anchor="center")
        self.counter.pack(side="left")
        self.next = ttk.Button(pages, text="›", width=4, command=lambda: self.show_page(1))
        self.next.pack(side="left")
        self.bind("<Control-s>", lambda _: self.confirm())
        self.bind("<Control-p>", lambda _: self.print_document())
        body = tk.Frame(self)
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, bg="#dbe4e6", highlightthickness=0)
        ybar = ttk.Scrollbar(body, command=self.canvas.yview)
        xbar = ttk.Scrollbar(body, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-int(e.delta / 120) * 3, "units"))
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", lambda _: self.close())
        try:
            self.show_page(0)
        except Exception:
            self.close()
            raise
        self.grab_set()

    def show_page(self, offset):
        target = self.index + offset
        if not 0 <= target < len(self.pdf):
            return
        self.index = target
        page = self.pdf[self.index]
        try:
            bitmap = page.render(scale=min(1.3, 1000 / max(page.get_size())))
            try:
                self.photo = self.image_tk.PhotoImage(bitmap.to_pil().copy(), master=self)
            finally:
                bitmap.close()
        finally:
            page.close()
        self.canvas.delete("all")
        self.canvas.create_image(12, 12, anchor="nw", image=self.photo)
        self.canvas.configure(scrollregion=(0, 0, self.photo.width() + 24, self.photo.height() + 24))
        self.canvas.yview_moveto(0)
        self.counter.configure(text=f"{self.index + 1} de {len(self.pdf)}")
        self.previous.state(["disabled"] if self.index == 0 else ["!disabled"])
        self.next.state(["disabled"] if self.index + 1 == len(self.pdf) else ["!disabled"])

    def confirm(self):
        try:
            self.saved = self.save_document()
        except Exception as exc:
            messagebox.showerror("Não foi possível salvar o PDF", str(exc), parent=self)
            return
        self.close()

    def print_document(self):
        if sys.platform != "win32":
            messagebox.showinfo("Imprimir", "Salve o PDF e abra-o em um leitor para imprimir.", parent=self)
            return
        from .printing import print_pdf

        try:
            print_pdf(self.pdf, self.winfo_id(), self.document_title)
        except Exception as exc:
            messagebox.showerror("Impressão", str(exc), parent=self)

    def close(self):
        self.pdf.close()
        self.destroy()


class App:
    def __init__(self, root, store):
        self.root = root
        self.store = store
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.busy = False
        self.scroll = None
        self.nav_buttons = {}
        self.root.title("Tributos Docs | Prefeitura Municipal de Biquinhas")
        self.root.configure(bg=BG)
        self.root.geometry("1190x790")
        self.root.minsize(930, 650)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.style()
        self.sidebar()
        self.main = tk.Frame(root, bg=BG)
        self.main.pack(side="left", fill="both", expand=True)
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.root.bind(event, self.wheel)
        self.dashboard()

    def style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=(FONT, 10), foreground=INK)
        style.configure("TButton", padding=(12, 7), background=PAPER, bordercolor=LINE, focuscolor=ACCENT)
        style.map("TButton", background=[("active", ACCENT_BG)])
        style.configure("Accent.TButton", background=ACCENT, foreground=PAPER, bordercolor=ACCENT)
        style.map("Accent.TButton", background=[("disabled", "#6d8883"), ("active", "#07574c")],
                  foreground=[("disabled", PAPER), ("active", PAPER)])
        style.configure("TEntry", padding=(8, 6), fieldbackground="#fbfdfd", bordercolor=LINE)
        style.configure("TSpinbox", padding=(8, 6), fieldbackground="#fbfdfd", bordercolor=LINE)
        style.configure("Treeview", rowheight=35, background=PAPER, fieldbackground=PAPER, bordercolor=LINE)
        style.configure("Treeview.Heading", padding=(8, 9), background="#e8eff1", font=(FONT, 9, "bold"))
        style.map("Treeview", background=[("selected", NAV_ACTIVE)], foreground=[("selected", PAPER)])

    def sidebar(self):
        side = tk.Frame(self.root, width=225, bg=NAV)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        mark = label(side, "TD", 13, bold=True, fg=PAPER, bg=ACCENT, padx=11, pady=7)
        mark.pack(anchor="w", padx=20, pady=(26, 9))
        label(side, "Tributos Docs", 17, bold=True, fg=PAPER, bg=NAV).pack(anchor="w", padx=20)
        label(side, "PREFEITURA DE BIQUINHAS", 8, bold=True, fg="#a7c3cb", bg=NAV).pack(anchor="w", padx=20, pady=(5, 30))
        label(side, "ÁREA DE TRABALHO", 8, bold=True, fg="#a7c3cb", bg=NAV).pack(anchor="w", padx=20, pady=(0, 10))
        links = [("Início", self.dashboard), ("Novo documento", self.catalog),
                 ("Histórico", self.history), ("Modelos Word", self.models),
                 ("Backup e dados", self.backups)]
        for index, (name, action) in enumerate(links, 1):
            button = tk.Button(side, text=f"{index:02d}    {name}", anchor="w", padx=13, pady=10,
                               bg=NAV, fg=PAPER, activebackground=NAV_ACTIVE, activeforeground=PAPER,
                               font=(FONT, 10), relief="flat", highlightthickness=0,
                               command=lambda callback=action: self.navigate(callback))
            button.pack(fill="x", padx=11, pady=2)
            self.nav_buttons[name] = button
        bottom = tk.Frame(side, bg=NAV)
        bottom.pack(side="bottom", fill="x", padx=20, pady=20)
        label(bottom, "●  Operação local", 9, bold=True, fg="#9ee0c7", bg=NAV).pack(anchor="w")
        label(bottom, "Setor de Tributos e Arrecadação", 8, fg="#a7c3cb", bg=NAV, wraplength=176, justify="left").pack(anchor="w", pady=(5, 0))

    def navigate(self, action):
        if not self.busy:
            action()

    def wheel(self, event):
        if self.scroll and str(event.widget).startswith(str(self.scroll)):
            self.scroll.wheel(event)

    def page(self, active, kicker, title, description):
        for child in self.main.winfo_children():
            child.destroy()
        self.scroll = None
        for name, button in self.nav_buttons.items():
            button.configure(bg=NAV_ACTIVE if name == active else NAV)
        header = tk.Frame(self.main, bg=PAPER, padx=27, pady=16)
        header.pack(fill="x")
        label(header, kicker.upper(), 8, bold=True, fg=ACCENT).pack(anchor="w")
        label(header, title, 22, bold=True).pack(anchor="w", pady=(4, 3))
        label(header, description, 10, fg=MUTED, wraplength=655, justify="left").pack(anchor="w")
        tk.Frame(self.main, bg=LINE, height=1).pack(fill="x")

    def scroll_page(self):
        self.scroll = ScrollArea(self.main)
        self.scroll.pack(fill="both", expand=True, padx=(26, 12), pady=17)
        return self.scroll.body

    def card(self, parent, title, note=""):
        box = tk.Frame(parent, bg=PAPER, highlightbackground=LINE, highlightthickness=1, padx=16, pady=13)
        box.pack(fill="x", padx=(0, 10), pady=(0, 11))
        label(box, title, 12, bold=True).pack(anchor="w")
        if note:
            label(box, note, fg=MUTED, wraplength=630, justify="left").pack(anchor="w", pady=(4, 8))
        return box

    def dashboard(self):
        self.page("Início", "Painel inicial", "Documentos municipais",
                  "Preencha um modelo, confira o PDF e arquive o documento na mesma estação.")
        body = self.scroll_page()
        summary = tk.Frame(body, bg=ACCENT_BG, padx=18, pady=17)
        summary.pack(fill="x", padx=(0, 10), pady=(0, 14))
        label(summary, "EMISSÃO DE DOCUMENTOS", 9, bold=True, fg=ACCENT, bg=ACCENT_BG).pack(anchor="w")
        label(summary, f"{self.store.model_count()} modelos ativos",
              14, bold=True, bg=ACCENT_BG).pack(anchor="w", pady=(5, 11))
        ttk.Button(summary, text="Novo documento", style="Accent.TButton", command=self.catalog).pack(anchor="w")
        box = self.card(body, "Arquivo de documentos", "Consulte os PDFs que estão nas pastas de trabalho.")
        ttk.Button(box, text="Consultar histórico", command=self.history).pack(anchor="w", pady=(8, 0))

    def catalog(self):
        self.page("Novo documento", "Emissão", "Escolha um documento",
                  "O formulário muda de acordo com as chaves do modelo Word.")
        body = self.scroll_page()
        models = self.store.models()
        if not models:
            box = self.card(body, "Cadastre um modelo para começar")
            ttk.Button(box, text="Gerenciar modelos", command=self.models).pack(anchor="w", pady=(7, 0))
        for model in models:
            note = f"{sum(map(len, model.groups().values()))} campos de preenchimento"
            if model.has_fractions():
                note += "  •  Frações de quantidade variável"
            box = self.card(body, model.title, note)
            ttk.Button(box, text="Preencher documento", command=lambda m=model: self.form(m)).pack(anchor="w")

    def form(self, model, previous=None):
        record_id = previous["record_id"] if previous else None
        self.page("Novo documento", "Revisão do histórico" if record_id else "Preenchimento",
                  model.title, "Os campos com * são obrigatórios. Confira o PDF antes de salvar.")
        footer = tk.Frame(self.main, bg=PAPER, padx=26, pady=10)
        footer.pack(side="bottom", fill="x")
        self.form_status = label(footer, "", fg=MUTED)
        self.form_status.pack(side="left")
        body = self.scroll_page()
        values = json.loads(previous["form"]) if previous else defaults()
        self.form_vars = {}
        reference = tk.StringVar(value=previous["reference"] if previous else "")
        groups = model.groups()
        if not any("NumeroDocumento" in group for group in groups.values()):
            box = self.card(body, "Identificação no histórico")
            label(box, "Número ou referência do processo (opcional)", fg=MUTED).pack(anchor="w", pady=(6, 4))
            ttk.Entry(box, textvariable=reference, width=36).pack(anchor="w")
        for group, fields in groups.items():
            box = self.card(body, "Imóvel de origem" if group == "Imóvel" and model.kind == "desdobro" else group)
            grid = tk.Frame(box, bg=PAPER)
            grid.pack(fill="x", pady=(7, 0))
            for col in range(3):
                grid.columnconfigure(col, weight=1, uniform="form")
            for index, (key, caption) in enumerate(fields.items()):
                row, col = divmod(index, 3)
                variable = tk.StringVar(value=values.get(key, ""))
                self.form_vars[key] = variable
                cell = tk.Frame(grid, bg=PAPER)
                cell.grid(row=row, column=col, sticky="ew", padx=(0, 11), pady=(0, 8))
                if key == "LogradouroImovel" and model.kind == "retificacao":
                    caption = "Nome da rua"
                required = " *" if key in REQUIRED.get(model.kind, ()) else ""
                label(cell, caption + required, 9, bold=True, wraplength=195, justify="left").pack(anchor="w", pady=(0, 4))
                ttk.Entry(cell, textvariable=variable, width=10).pack(fill="x")
        self.fraction_vars = []
        self.active_fractions = 0
        if model.has_fractions():
            rows = json.loads(previous["fractions"]) if previous else [{}, {}]
            box = self.card(body, "Frações resultantes", "Informe a quantidade e preencha o lote, a área e o endereço de cada fração.")
            controls = tk.Frame(box, bg=PAPER)
            controls.pack(fill="x", pady=(2, 10))
            label(controls, "Quantidade de frações", 9, bold=True).pack(side="left")
            self.count = tk.StringVar(value=str(len(rows)))
            self.fraction_table = tk.Frame(box, bg=PAPER)
            spin = ttk.Spinbox(controls, from_=1, to=MAX_FRACTIONS, width=5, textvariable=self.count)
            spin.pack(side="left", padx=8)
            self.count.trace_add("write", self.update_fractions)
            for row in rows:
                self.fraction_vars.append({key: tk.StringVar(value=row.get(key, "")) for key in FRACTION_FIELDS})
            self.fraction_table.pack(fill="x")
            self.update_fractions()
        self.generate_button = ttk.Button(footer, text="Gerar documento", style="Accent.TButton",
            command=lambda: self.generate(model, reference.get(), record_id))
        self.generate_button.pack(side="right")
        ttk.Button(footer, text="Voltar", command=lambda: self.navigate(self.history if record_id else self.catalog)).pack(side="right", padx=8)

    def update_fractions(self, *_):
        try:
            quantity = int(self.count.get())
        except ValueError:
            return
        if not 1 <= quantity <= MAX_FRACTIONS or quantity == self.active_fractions:
            return
        self.active_fractions = quantity
        for child in self.fraction_table.winfo_children():
            child.destroy()
        for col, caption in enumerate(["#", *FRACTION_FIELDS.values()]):
            if col:
                self.fraction_table.columnconfigure(col, weight=2 if col == 3 else 1, uniform="fractions")
            label(self.fraction_table, caption, 9, bold=True, bg=BG, padx=6, pady=4).grid(row=0, column=col, sticky="ew")
        while len(self.fraction_vars) < quantity:
            self.fraction_vars.append({key: tk.StringVar() for key in FRACTION_FIELDS})
        for index, variables in enumerate(self.fraction_vars[:quantity], 1):
            label(self.fraction_table, f"{index:02d}", 9, bold=True, fg=ACCENT, padx=6).grid(row=index, column=0, sticky="w")
            for col, variable in enumerate(variables.values(), 1):
                ttk.Entry(self.fraction_table, textvariable=variable, width=8).grid(row=index, column=col, sticky="ew", padx=3, pady=3)

    def generate(self, model, reference, record_id):
        try:
            if model.has_fractions() and (not self.count.get().isdigit() or int(self.count.get()) != self.active_fractions):
                raise ValueError(f"Informe entre 1 e {MAX_FRACTIONS} frações.")
            values = {key: var.get().strip() for key, var in self.form_vars.items()}
            fractions = [{key: var.get().strip() for key, var in row.items()}
                         for row in self.fraction_vars[:self.active_fractions]]
            filled = prepare(model, values, fractions)
        except Exception as exc:
            self.error(exc)
            return
        self.busy = True
        self.generate_button.configure(state="disabled")
        self.form_status.configure(text="Gerando o PDF…")
        future = self.executor.submit(export_pdf, filled)
        def complete():
            if not future.done():
                self.root.after(90, complete)
                return
            try:
                data = future.result()
                self.form_status.configure(text="Prévia pronta para conferência.")
                def archive():
                    return self.store.save_record(
                        model, values, fractions, data, values.get("NumeroDocumento", reference).strip(), record_id)

                preview = Preview(self.root, data, model.title, self.store.destination(model.kind), archive)
                self.root.wait_window(preview)
                if preview.saved:
                    rid, version, path = preview.saved
                    self.history()
                    messagebox.showinfo("PDF salvo", f"TD-{rid:06d} · revisão {version:02d}\n\n{path}",
                                        parent=self.root)
            except Exception as exc:
                self.error(exc)
            finally:
                self.busy = False
                if self.generate_button.winfo_exists():
                    self.generate_button.configure(state="normal")
                    self.form_status.configure(text="")
        self.root.after(90, complete)

    def history(self):
        self.page("Histórico", "Arquivo local", "Histórico de documentos",
                  "PDFs encontrados nas pastas do repositório. Pesquise pelo nome do arquivo, tipo, proprietário ou referência.")
        body = tk.Frame(self.main, bg=BG, padx=26, pady=16)
        body.pack(fill="both", expand=True)
        query = tk.StringVar()
        searchbar = tk.Frame(body, bg=BG)
        searchbar.pack(fill="x", pady=(0, 10))
        label(searchbar, "Pesquisar", 9, bold=True, bg=BG).pack(side="left")
        ttk.Entry(searchbar, textvariable=query, width=35).pack(side="left", padx=8)
        ttk.Button(searchbar, text="Abrir pastas", command=lambda: self.safe(
            lambda: open_path(self.store.archive))).pack(side="right")
        ttk.Button(searchbar, text="Atualizar", command=lambda: reload_archive()).pack(side="right", padx=5)
        panel = tk.Frame(body, bg=BG)
        panel.pack(fill="both", expand=True)
        columns = [("code", "Código / rev.", 121), ("reference", "Nº / referência", 113),
                   ("owner", "Proprietário", 166), ("model", "Arquivo / tipo", 210), ("date", "Emissão", 100)]
        table = ttk.Treeview(panel, columns=[x[0] for x in columns], show="headings", selectmode="browse")
        bar = ttk.Scrollbar(panel, command=table.yview)
        table.configure(yscrollcommand=bar.set)
        bar.pack(fill="y", side="right")
        table.pack(fill="both", expand=True)
        for key, heading, width in columns:
            table.heading(key, text=heading, anchor="w")
            table.column(key, width=width, minwidth=70, stretch=key in {"owner", "model"})
        status = label(body, "", fg=MUTED, bg=BG)
        status.pack(anchor="w", pady=(8, 0))

        rows, documents = {}, []

        def reload_archive():
            try:
                documents[:] = self.store.history()
            except Exception as exc:
                status.configure(text=str(exc))
                return
            refresh()

        def refresh(*_):
            table.delete(*table.get_children())
            rows.clear()
            try:
                for index, row in enumerate(documents):
                    if not matches_history(row, query.get()):
                        continue
                    rid, version = row["record_id"], row["version"]
                    rows[str(index)] = row
                    table.insert("", "end", iid=str(index), values=(
                        f"TD-{rid:06d} / {version:02d}" if rid else "Arquivo da pasta",
                        row["reference"] or "—", row["owner"] or "—",
                        row["title"] if rid else row["path"].name,
                        date.fromisoformat(row["created"][:10]).strftime("%d/%m/%Y")))
                status.configure(text=f"{len(rows)} PDF(s) encontrado(s)")
            except Exception as exc:
                status.configure(text=str(exc))

        def selected(action):
            selection = table.selection()
            if not selection:
                messagebox.showinfo("Histórico", "Selecione um documento na lista.", parent=self.root)
                return
            try:
                action(rows[selection[0]])
            except Exception as exc:
                self.error(exc)

        def correct(row):
            if row["record_id"] is None:
                raise ValueError("Este PDF foi adicionado pela pasta e não possui formulário registrado.")
            record = self.store.record(row["record_id"], row["version"])
            self.form(Model(record["model_id"], record["title"], record["kind"], record["document"]), record)

        def show_pdf(row):
            if not row["path"].is_file():
                raise FileNotFoundError("O PDF não está mais na pasta. Clique em Atualizar.")
            open_path(row["path"])

        query.trace_add("write", refresh)
        table.bind("<Double-1>", lambda _: selected(show_pdf))
        actions = tk.Frame(body, bg=BG)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Abrir PDF", command=lambda: selected(show_pdf)).pack(side="left")
        ttk.Button(actions, text="Corrigir e gerar nova revisão", style="Accent.TButton",
                   command=lambda: selected(correct)).pack(side="left", padx=8)
        reload_archive()

    def models(self):
        self.page("Modelos Word", "Gerenciamento", "Modelos Word",
                  "Edite o texto no Word ou Writer. As chaves entre < e > criam campos no formulário.")
        body = self.scroll_page()
        action = tk.Frame(body, bg=BG)
        action.pack(fill="x", padx=(0, 10), pady=(0, 10))
        ttk.Button(action, text="Cadastrar modelo .docx", style="Accent.TButton",
                   command=lambda: self.model_editor()).pack(side="left")
        for model in self.store.models():
            box = self.card(body, model.title, f"{len(model.keys())} chaves · arquivo Word editável")
            row = tk.Frame(box, bg=PAPER)
            row.pack(fill="x")
            ttk.Button(row, text="Alterar modelo", command=lambda m=model: self.model_editor(m)).pack(side="left")
            ttk.Button(row, text="Excluir", command=lambda m=model: self.delete_model(m)).pack(side="left", padx=8)

    def delete_model(self, model):
        if messagebox.askyesno("Excluir modelo", f"Remover «{model.title}» da lista de novos documentos?",
                               parent=self.root):
            try:
                self.store.delete_model(model.id)
                self.models()
            except Exception as exc:
                self.error(exc)

    def model_editor(self, model=None):
        dialog = tk.Toplevel(self.root)
        dialog.title("Alterar modelo" if model else "Cadastrar modelo")
        dialog.geometry("690x510")
        dialog.minsize(560, 410)
        dialog.transient(self.root)
        dialog.grab_set()
        frame = tk.Frame(dialog, bg=PAPER, padx=20, pady=16)
        frame.pack(fill="both", expand=True)
        label(frame, "Modelo do documento", 15, bold=True).pack(anchor="w")
        label(frame, "Nome do modelo", 9, bold=True).pack(anchor="w", pady=(12, 4))
        title = tk.StringVar(value=model.title if model else "")
        ttk.Entry(frame, textvariable=title).pack(fill="x")
        label(frame, "Edite frases e formatação no Word ou Writer. Para criar um campo, digite uma chave como <CimImovel>.",
              fg=MUTED, wraplength=530, justify="left").pack(anchor="w", pady=(10, 6))
        working = TemporaryDirectory(prefix="modelo-", dir=self.store.root, ignore_cleanup_errors=True)
        editable = Path(working.name) / "Modelo.docx"
        if model:
            editable.write_bytes(model.document)
        source = [editable if model else None]

        def close():
            working.cleanup()
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", close)
        box = tk.Frame(frame, bg=PAPER)
        box.pack(fill="x")
        status = label(frame, "", fg=MUTED, wraplength=530, justify="left")
        footer = tk.Frame(frame, bg=PAPER)
        footer.pack(side="bottom", fill="x", pady=(8, 0))
        listframe = tk.Frame(frame, bg=PAPER)
        keys = tk.Listbox(listframe, height=7, bg="#f5f9fa", fg=INK, relief="flat",
                          highlightthickness=1, highlightbackground=LINE, exportselection=False, font=(FONT, 10))
        bar = ttk.Scrollbar(listframe, command=keys.yview)
        bar.pack(side="right", fill="y")
        keys.configure(yscrollcommand=bar.set)
        keys.pack(fill="both", expand=True)

        def current():
            if source[0] is None:
                raise ValueError("Selecione um arquivo .docx.")
            return Model(model.id if model else uuid4().hex, title.get(),
                         model.kind if model else "custom", source[0].read_bytes())

        def refresh():
            keys.delete(0, "end")
            if source[0] is None:
                status.configure(text="Selecione um arquivo do Word com as chaves de preenchimento.")
                return
            try:
                fields = current().keys()
                for key in fields:
                    keys.insert("end", f"<{key}>")
                status.configure(text=f"{len(fields)} chaves encontradas no arquivo selecionado.")
            except Exception as exc:
                self.error(exc, dialog)

        def select_file():
            path = filedialog.askopenfilename(parent=dialog, filetypes=[("Modelo Word", "*.docx")])
            if path:
                source[0] = Path(path)
                if not title.get().strip():
                    title.set(Path(path).stem)
                refresh()

        def edit():
            if source[0] is None:
                self.error(ValueError("Selecione um arquivo .docx primeiro."), dialog)
                return
            try:
                if source[0] != editable:
                    editable.write_bytes(source[0].read_bytes())
                    source[0] = editable
                open_path(editable)
                status.configure(text="Salve no Word ou Writer, feche o editor e clique em Aplicar alterações.")
            except Exception as exc:
                self.error(exc, dialog)

        def apply():
            try:
                self.store.save_model(current())
            except Exception as exc:
                self.error(exc, dialog)
                return
            close()
            self.models()

        def copy_key():
            selection = keys.curselection()
            if selection:
                dialog.clipboard_clear()
                dialog.clipboard_append(keys.get(selection[0]))

        ttk.Button(box, text="Selecionar .docx", command=select_file).pack(side="left")
        ttk.Button(box, text="Editar no Word ou Writer", command=edit).pack(side="left", padx=8)
        ttk.Button(box, text="Reler chaves", command=refresh).pack(side="left")
        status.pack(anchor="w", pady=(8, 5))
        listframe.pack(fill="both", expand=True)
        ttk.Button(footer, text="Copiar chave", command=copy_key).pack(side="left")
        ttk.Button(footer, text="Aplicar alterações", style="Accent.TButton", command=apply).pack(side="right")
        ttk.Button(footer, text="Cancelar", command=close).pack(side="right", padx=8)
        refresh()

    def backups(self):
        self.page("Backup e dados", "Estação de trabalho", "Backup e dados",
                  "Os arquivos e o histórico ficam neste computador e usam o perfil local do Windows.")
        body = self.scroll_page()
        box = self.card(body, "Pasta de trabalho", str(self.store.root))
        ttk.Button(box, text="Abrir pasta local", command=lambda: self.safe(lambda: open_path(self.store.root))).pack(anchor="w")
        box = self.card(body, "Criar backup", "O arquivo inclui modelos, formulários, revisões e os PDFs arquivados.")
        ttk.Button(box, text="Salvar backup", style="Accent.TButton", command=self.make_backup).pack(anchor="w")
        box = self.card(body, "Restaurar backup", "Restaure um backup criado pelo Tributos Docs. A operação substitui os modelos e o histórico desta estação.")
        ttk.Button(box, text="Restaurar arquivo", command=self.restore_backup).pack(anchor="w")

    def make_backup(self):
        filename = filedialog.asksaveasfilename(parent=self.root, defaultextension=".sqlite3",
                    initialfile=f"Tributos-Docs-{date.today().isoformat()}.sqlite3",
                    filetypes=[("Backup Tributos Docs", "*.sqlite3")])
        if filename:
            def save():
                self.store.backup(filename)
                messagebox.showinfo("Backup", "Backup salvo.", parent=self.root)
            self.safe(save)

    def restore_backup(self):
        filename = filedialog.askopenfilename(parent=self.root, filetypes=[("Backup Tributos Docs", "*.sqlite3")])
        if filename and messagebox.askyesno("Restaurar backup", "Substituir os modelos e o histórico atuais?",
                                             parent=self.root):
            def restore():
                self.store.backup(self.store.root / f"antes-de-restaurar-{uuid4().hex[:8]}.sqlite3")
                self.store.restore(filename)
                messagebox.showinfo("Backup", "Backup restaurado.", parent=self.root)
                self.dashboard()
            self.safe(restore)

    def safe(self, action):
        try:
            action()
        except Exception as exc:
            self.error(exc)

    def error(self, exc, parent=None):
        messagebox.showerror("Tributos Docs", str(exc), parent=parent or self.root)

    def close(self):
        if not self.busy:
            self.executor.shutdown(wait=False)
            self.store.connection.close()
            self.root.destroy()


def main():
    root = tk.Tk()
    try:
        App(root, Storage(data_directory(), archive_directory()))
    except Exception as exc:
        root.withdraw()
        messagebox.showerror("Tributos Docs", str(exc), parent=root)
        root.destroy()
        return
    root.mainloop()

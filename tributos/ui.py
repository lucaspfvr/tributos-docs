"""Offline desktop interface for generating and managing municipal documents."""

import json
import os
import subprocess
import sys
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from dataclasses import astuple
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import filedialog, font, messagebox, ttk
from uuid import uuid4

from .appearance import DEFAULT_THEME, PALETTES
from .controls import CommandControls
from .docx import export_pdf
from .models import (ADDRESS_PARTS, FRACTION_FIELDS, MAX_FRACTIONS, REQUIRED, TAX_ID_FIELDS, Model, defaults,
                     format_tax_id, normalize_fields, prepare)
from .storage import Storage, atomic_write, matches_history, pdf_stem

T = PALETTES[DEFAULT_THEME]
FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
APP_TITLE = "Setor de Tributos e Arrecadação - Documentos"


def label(parent, text, size=10, *, bold=False, fg=None, bg=None, anchor="w", **kwargs):
    return tk.Label(parent, text=text, fg=fg or T.ink, bg=bg or T.paper, anchor=anchor,
                    font=(FONT, size, "bold" if bold else "normal"), **kwargs)



def mask_tax_id(entry, variable):
    pending = None

    def update():
        nonlocal pending
        value = variable.get()
        formatted = format_tax_id(value, partial=True)
        if formatted != value:
            cursor = entry.index("insert")
            digits = sum("0" <= char <= "9" for char in value[:cursor])
            position = 0
            while digits and position < len(formatted):
                digits -= "0" <= formatted[position] <= "9"
                position += 1
            variable.set(formatted)
            entry.icursor(len(formatted) if cursor == len(value) else position)
        pending = None

    def schedule(*_):
        nonlocal pending
        if pending is None:
            pending = entry.after_idle(update)

    def cleanup(_):
        variable.trace_remove("write", trace)
        if pending is not None:
            entry.after_cancel(pending)

    trace = variable.trace_add("write", schedule)
    entry.bind("<Destroy>", cleanup, add="+")
    schedule()


def workflow(parent, active):
    strip = tk.Frame(parent, bg=T.nav, pady=9)
    center = tk.Frame(strip, bg=T.nav)
    center.pack()
    for index, step in enumerate(("Preencher", "Revisar", "Arquivar"), 1):
        if index > 1:
            label(center, "─────", 9, fg=T.line, bg=T.nav).pack(side="left", padx=9)
        current = index == active
        color = T.accent if index <= active else T.muted
        number = "✓" if index < active else f"{index:02d}"
        label(center, f"{number}   {step}", 10, bold=current, fg=color,
              bg=T.active if current else T.nav, padx=12, pady=6).pack(side="left")
    return strip


def review_details(model, fields, fractions, reference):
    street = fields.get("LogradouroImovel")
    address = ", ".join(filter(None, (street, fields.get("NumeroImovel"))))
    if not address and fractions:
        addresses = dict.fromkeys(", ".join(filter(None, (row.get("endereco"), row.get("numero"))))
                                  for row in fractions)
        address = "\n".join(filter(None, addresses))
    lots = ", ".join(row["lote"] for row in fractions)
    area = fields.get("AreaTerreno") or fields.get("AreaMensurada")
    if fractions:
        total = sum(Decimal(row["area"].replace(".", "").replace(",", ".")) for row in fractions)
        area = f"{total:,.2f}".translate(str.maketrans(",.", ".,")) + " m²"
    elif area:
        area += " m²"
    return [
        ("Nº do documento", reference),
        ("Proprietário", fields.get("NomeProprietario")),
        ("Endereço dos lotes" if fractions and not street else "Endereço", address),
        ("Lotes resultantes", f"{len(fractions)} · {lots}" if lots else None),
        ("Soma das áreas dos lotes" if fractions else "Área do imóvel", area),
    ]

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
    def __init__(self, parent, background=None):
        background = background or T.bg
        super().__init__(parent, bg=background)
        self.canvas = tk.Canvas(self, bg=background, highlightthickness=0)
        bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.body = tk.Frame(self.canvas, bg=background)
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
    def __init__(self, parent, data, title, destination, suggested, details, save, controls):
        import pypdfium2
        from PIL import ImageTk

        self.pdf = pypdfium2.PdfDocument(data)
        self.image_tk = ImageTk
        self.index = 0
        self.saved = None
        self.save_document = save
        self.document_title = title
        self.zoom = 1.0
        self.fit_job = None
        super().__init__(parent)
        self.title(f"{title} · Prévia")
        width, height = self.maxsize()
        self.geometry(f"{min(1130, width - 25)}x{min(760, height - 55)}+18+22")
        self.minsize(610, 340)
        self.transient(parent)

        header = tk.Frame(self, bg=T.paper, padx=18, pady=10)
        header.pack(fill="x")
        label(header, "REVISÃO DO DOCUMENTO", 9, bold=True, fg=T.accent).pack(anchor="w")
        label(header, title, 17, bold=True).pack(anchor="w", pady=(2, 0))
        workflow(self, 2).pack(fill="x")

        toolbar = tk.Frame(self, bg=T.paper, padx=18, pady=8)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Voltar", command=self.close).pack(side="left")
        controls.button(toolbar, "Salvar PDF", "save", self.confirm, primary=True).pack(side="right")
        controls.button(toolbar, "Imprimir", "print", self.print_document).pack(side="right", padx=8)
        self.bind("<Control-s>", lambda _: self.confirm())
        self.bind("<Control-p>", lambda _: self.print_document())

        content = tk.Frame(self, bg=T.bg, padx=12, pady=10)
        content.pack(fill="both", expand=True)
        side = tk.Frame(content, bg=T.paper, width=267, highlightbackground=T.line, highlightthickness=1)
        side.pack(side="right", fill="y", padx=(10, 0))
        side.pack_propagate(False)
        label(side, "Resumo da revisão", 13, bold=True, padx=13, pady=9).pack(fill="x")
        summary = ScrollArea(side, background=T.paper)
        summary.pack(fill="both", expand=True, padx=9, pady=(0, 7))
        for caption, value in details:
            if caption == "Lotes resultantes" and not value:
                continue
            cell = tk.Frame(summary.body, bg=T.paper, padx=5, pady=5)
            cell.pack(fill="x")
            label(cell, caption, 9, fg=T.muted).pack(anchor="w")
            label(cell, value or "Não informado", 10, bold=bool(value), wraplength=220,
                  justify="left").pack(anchor="w", pady=(2, 0))
            tk.Frame(summary.body, bg=T.line, height=1).pack(fill="x", padx=5)
        label(summary.body, "ARQUIVO PDF", 8, bold=True, fg=T.accent,
              padx=5, pady=6).pack(fill="x")
        label(summary.body, suggested, 9, wraplength=220, padx=5, justify="left").pack(fill="x")
        label(summary.body, str(destination), 8, fg=T.muted, wraplength=220,
              padx=5, pady=4, justify="left").pack(fill="x")
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind(event, lambda event: summary.wheel(event)
                      if str(event.widget).startswith(str(summary)) else None, add="+")

        viewer = tk.Frame(content, bg=T.paper, highlightbackground=T.line, highlightthickness=1)
        viewer.pack(side="left", fill="both", expand=True)
        pages = tk.Frame(viewer, bg=T.paper, padx=6, pady=4)
        pages.pack(fill="x")
        self.previous = ttk.Button(pages, text="<", width=3, command=lambda: self.show_page(-1))
        self.previous.pack(side="left")
        self.counter = label(pages, "", fg=T.muted, width=8, anchor="center")
        self.counter.pack(side="left")
        self.next = ttk.Button(pages, text=">", width=3, command=lambda: self.show_page(1))
        self.next.pack(side="left")
        ttk.Button(pages, text="+", width=3, command=lambda: self.set_zoom(1.2)).pack(side="right")
        self.zoom_label = label(pages, "100%", 9, fg=T.muted, padx=4)
        self.zoom_label.pack(side="right")
        ttk.Button(pages, text="-", width=3, command=lambda: self.set_zoom(1 / 1.2)).pack(side="right")
        body = tk.Frame(viewer, bg=T.canvas)
        body.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(body, bg=T.canvas, highlightthickness=0)
        ybar = ttk.Scrollbar(body, command=self.canvas.yview)
        xbar = ttk.Scrollbar(body, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.canvas.bind("<Configure>", self.resize_preview)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-int(e.delta / 120) * 3, "units"))
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", lambda _: self.close())
        try:
            self.update_idletasks()
            self.show_page(0)
        except Exception:
            self.close()
            raise
        self.grab_set()

    def resize_preview(self, _):
        if self.fit_job is not None:
            self.after_cancel(self.fit_job)
        self.fit_job = self.after(130, self.render_resize)

    def render_resize(self):
        self.fit_job = None
        self.show_page(0, reset_scroll=False)

    def set_zoom(self, factor):
        self.zoom = min(3.0, max(0.55, self.zoom * factor))
        self.zoom_label.configure(text=f"{self.zoom:.0%}")
        self.show_page(0, reset_scroll=False)

    def show_page(self, offset, reset_scroll=True):
        target = self.index + offset
        if not 0 <= target < len(self.pdf):
            return
        self.index = target
        page = self.pdf[self.index]
        try:
            w, h = page.get_size()
            available = max(100, self.canvas.winfo_width() - 32)
            bitmap = page.render(scale=min(available * self.zoom / w, 2100 / max(w, h)))
            try:
                self.photo = self.image_tk.PhotoImage(bitmap.to_pil().copy(), master=self)
            finally:
                bitmap.close()
        finally:
            page.close()
        self.canvas.delete("all")
        area = max(self.canvas.winfo_width(), self.photo.width() + 24)
        self.canvas.create_image((area - self.photo.width()) // 2, 12, anchor="nw", image=self.photo)
        self.canvas.configure(scrollregion=(0, 0, area, self.photo.height() + 24))
        if reset_scroll:
            self.canvas.yview_moveto(0)
            self.canvas.xview_moveto(0)
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
        if self.fit_job is not None:
            self.after_cancel(self.fit_job)
        self.pdf.close()
        self.destroy()


class App:
    def __init__(self, root, store):
        global T
        self.root = root
        self.store = store
        self.settings_path = store.root / "appearance.json"
        try:
            theme = json.loads(self.settings_path.read_text(encoding="utf-8"))["theme"]
            self.theme = theme if theme in PALETTES else DEFAULT_THEME
        except (OSError, ValueError, KeyError, TypeError):
            self.theme = DEFAULT_THEME
        T = PALETTES[self.theme]
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.busy = False
        self.scroll = None
        self.nav_buttons = {}
        self.root.title(APP_TITLE)
        self.root.configure(bg=T.bg)
        width, height = self.root.maxsize()
        self.root.geometry(f"{min(1190, width - 24)}x{min(770, height - 38)}+12+12")
        self.root.minsize(640, 420)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.controls = CommandControls(root)
        self.style()
        self.topbar()
        footer = tk.Frame(root, bg=T.nav, padx=24, pady=6)
        footer.pack(side="bottom", fill="x")
        label(footer, "Desenvolvido por Lucas Cézar", 9, fg=T.muted, bg=T.nav).pack(side="right")
        self.main = tk.Frame(root, bg=T.bg)
        self.main.pack(fill="both", expand=True)
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.root.bind(event, self.wheel)
        self.dashboard()

    def style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        self.controls.style(style, T)
        style.configure(".", font=(FONT, 10), foreground=T.ink)
        style.configure("TButton", padding=(13, 8), foreground=T.accent, font=(FONT, 10, "bold"),
                        focuscolor=T.accent)
        style.configure("Nav.TButton", padding=(9, 7), width=-1, foreground=T.ink, font=(FONT, 10))
        style.configure("Selected.Nav.TButton", foreground=T.accent, font=(FONT, 10, "bold"))
        style.map("TButton", foreground=[("disabled", T.muted)])
        style.configure("Accent.TButton", foreground="#ffffff", focuscolor="#ffffff")
        style.map("Accent.TButton", foreground=[("disabled", "#ffffff"), ("active", "#ffffff")])
        style.configure("TEntry", padding=(8, 6), fieldbackground=T.paper, bordercolor=T.line)
        style.configure("Form.TEntry", padding=(7, 4))
        style.configure("TCheckbutton", background=T.paper)
        style.map("TCheckbutton", background=[("active", T.tint)])
        style.configure("TSpinbox", padding=(8, 6), fieldbackground=T.paper, bordercolor=T.line)
        style.configure("TRadiobutton", background=T.paper, foreground=T.ink)
        style.map("TRadiobutton", background=[("active", T.tint)])
        line_height = font.nametofont("TkDefaultFont", root=self.root).metrics("linespace")
        style.configure("Treeview", rowheight=max(31, line_height + 12), background=T.paper,
                        fieldbackground=T.paper, bordercolor=T.line)
        style.configure("Treeview.Heading", padding=(8, 9), background=T.tint, font=(FONT, 9, "bold"))
        style.map("Treeview", background=[("selected", T.active)], foreground=[("selected", T.ink)])

    def topbar(self):
        header = self.header = tk.Frame(self.root, bg=T.paper, padx=18, pady=7)
        header.pack(fill="x")
        header.columnconfigure(1, weight=1)
        logo = Path(__file__).with_name("assets") / "municipio.png"
        self.logo = tk.PhotoImage(file=str(logo))
        self.root.iconphoto(True, self.logo)
        if sys.platform == "win32":
            try:
                self.root.iconbitmap(default=str(logo.with_suffix(".ico")))
            except tk.TclError:
                pass
        tk.Label(header, image=self.logo, bg=T.paper).grid(row=0, column=0, padx=(0, 9))
        brand = tk.Frame(header, bg=T.paper)
        brand.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        title = label(brand, APP_TITLE, 15, bold=True, justify="left", wraplength=435)
        title.pack(fill="x")
        brand.bind("<Configure>", lambda event: title.configure(wraplength=max(80, event.width)))
        label(brand, "PREFEITURA DE BIQUINHAS", 8, fg=T.muted).pack(anchor="w")
        ttk.Button(header, text="⚙  Configurações", command=self.settings).grid(row=0, column=2, sticky="e")

        nav = self.nav = tk.Frame(self.root, bg=T.nav, padx=16, pady=4)
        nav.pack(fill="x")
        links = [("Início", self.dashboard),
                 ("Desdobro", lambda: self.open_kind("desdobro")),
                 ("Retificação", lambda: self.open_kind("retificacao")),
                 ("Avaliação", lambda: self.open_kind("avaliacao")),
                 ("Histórico", self.history)]
        for name, action in links:
            button = ttk.Button(nav, text=name, style="Nav.TButton",
                                command=lambda action=action: self.navigate(action))
            button.pack(side="left", padx=(0, 4))
            self.nav_buttons[name] = button
        ttk.Button(nav, text="Modelos e dados ▾", style="Nav.TButton", command=self.utility_menu).pack(side="right")
        self.nav_rule = tk.Frame(self.root, bg=T.line, height=1)
        self.nav_rule.pack(fill="x")

    def open_kind(self, kind):
        model = next((model for model in self.store.models() if model.kind == kind), None)
        self.form(model) if model else self.catalog()

    def utility_menu(self):
        if self.busy:
            return
        menu = tk.Menu(self.root, tearoff=False)
        menu.add_command(label="Modelos Word", command=self.models)
        menu.add_command(label="Backup e dados", command=self.backups)
        menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())

    def settings(self):
        if self.busy:
            return
        dialog = tk.Toplevel(self.root, bg=T.paper)
        dialog.title("Configurações")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        body = tk.Frame(dialog, bg=T.paper, padx=21, pady=17)
        body.pack(fill="both", expand=True)
        label(body, "Aparência do sistema", 15, bold=True).pack(anchor="w")
        label(body, "Escolha uma paleta para esta estação.", 10, fg=T.muted).pack(anchor="w", pady=(3, 12))
        selection = tk.StringVar(value=self.theme)
        for key, choice in PALETTES.items():
            row = tk.Frame(body, bg=T.paper)
            row.pack(fill="x", pady=3)
            ttk.Radiobutton(row, text=choice.name, value=key, variable=selection, width=21).pack(side="left")
            for color in (choice.bg, choice.active, choice.accent):
                tk.Frame(row, width=19, height=17, bg=color,
                         highlightbackground=T.line, highlightthickness=1).pack(side="left", padx=2)
        actions = tk.Frame(body, bg=T.paper)
        actions.pack(fill="x", pady=(13, 0))

        def apply():
            try:
                atomic_write(self.settings_path, json.dumps({"theme": selection.get()}).encode("utf-8"))
            except OSError as exc:
                self.error(exc, dialog)
                return
            dialog.destroy()
            self.change_theme(selection.get())

        ttk.Button(actions, text="Aplicar", style="Accent.TButton", command=apply).pack(side="right")
        ttk.Button(actions, text="Cancelar", command=dialog.destroy).pack(side="right", padx=8)
        dialog.bind("<Escape>", lambda _: dialog.destroy())
        dialog.wait_visibility()
        dialog.grab_set()

    def change_theme(self, key):
        global T
        old = T
        self.theme = key
        T = PALETTES[key]
        colors = dict(zip(astuple(old)[1:], astuple(T)[1:]))

        def recolor(widget):
            if not isinstance(widget, ttk.Widget):
                options = widget.keys()
                for option in ("background", "foreground", "highlightbackground"):
                    if option in options:
                        current = str(widget.cget(option))
                        if current in colors:
                            widget.configure({option: colors[current]})
            for child in widget.winfo_children():
                recolor(child)

        recolor(self.root)
        self.style()

    def navigate(self, action):
        if not self.busy:
            action()

    def wheel(self, event):
        if self.scroll and str(event.widget).startswith(str(self.scroll)):
            self.scroll.wheel(event)

    def page(self, active, kicker, title, description, *, compact=False):
        for child in self.main.winfo_children():
            child.destroy()
        self.scroll = None
        for name, button in self.nav_buttons.items():
            button.configure(style="Selected.Nav.TButton" if name == active else "Nav.TButton")
        for bar in (self.header, self.nav, self.nav_rule):
            if compact:
                bar.pack_forget()
            elif not bar.winfo_manager():
                bar.pack(fill="x", before=self.main)
        if compact:
            header = tk.Frame(self.main, bg=T.paper, padx=13, pady=7)
            header.pack(fill="x")
            ttk.Button(header, text="Menu ▾", style="Nav.TButton",
                       command=self.form_menu).pack(side="right", padx=(8, 0))
            steps = tk.Frame(header, bg=T.paper)
            steps.pack(side="right")
            label(steps, "01 Preencher", 9, bold=True, fg=T.accent, bg=T.tint,
                  padx=7, pady=4).pack(side="left")
            label(steps, "  02 Revisar   03 Arquivar", 9, fg=T.muted).pack(side="left")
            label(header, title, 14, bold=True, wraplength=285, justify="left").pack(side="left", fill="x")
            tk.Frame(self.main, bg=T.line, height=1).pack(fill="x")
            return
        header = tk.Frame(self.main, bg=T.paper, padx=20, pady=11)
        header.pack(fill="x")
        label(header, kicker.upper(), 8, bold=True, fg=T.accent).pack(anchor="w")
        label(header, title, 22, bold=True).pack(anchor="w", pady=(4, 3))
        description_label = label(header, description, 10, fg=T.muted, wraplength=595, justify="left")
        description_label.pack(fill="x")
        header.bind("<Configure>", lambda event: description_label.configure(
            wraplength=max(120, min(730, event.width - 40))))
        tk.Frame(self.main, bg=T.line, height=1).pack(fill="x")

    def form_menu(self):
        if self.busy:
            return
        menu = tk.Menu(self.root, tearoff=False)
        menu.add_command(label="Início", command=self.dashboard)
        menu.add_command(label="Histórico", command=self.history)
        menu.add_separator()
        menu.add_command(label="Configurações", command=self.settings)
        menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())

    def link_address(self, container, linked, pairs, entries):
        def sync(*_):
            for source, target in pairs:
                if linked.get():
                    self.form_vars[target].set(self.form_vars[source].get())
                entries[target].configure(state="readonly" if linked.get() else "normal")

        traces = [(variable, variable.trace_add("write", sync)) for variable in
                  [linked, *(self.form_vars[source] for source, _ in pairs)]]

        def cleanup(_):
            for variable, trace in traces:
                variable.trace_remove("write", trace)

        container.bind("<Destroy>", cleanup, add="+")
        sync()

    def scroll_page(self):
        self.scroll = ScrollArea(self.main)
        self.scroll.pack(fill="both", expand=True, padx=(26, 12), pady=17)
        return self.scroll.body

    def card(self, parent, title, note=""):
        box = tk.Frame(parent, bg=T.paper, highlightbackground=T.line, highlightthickness=1, padx=16, pady=13)
        box.pack(fill="x", padx=(0, 10), pady=(0, 11))
        label(box, title, 12, bold=True).pack(anchor="w")
        if note:
            label(box, note, fg=T.muted, wraplength=630, justify="left").pack(anchor="w", pady=(4, 8))
        return box

    def dashboard(self):
        self.page("Início", "Painel inicial", "Documentos municipais",
                  "Preencha um modelo, confira o PDF e arquive o documento na mesma estação.")
        body = self.scroll_page()
        heading = tk.Frame(body, bg=T.bg)
        heading.pack(fill="x", padx=(0, 10), pady=(5, 12))
        label(heading, "Começar um documento", 16, bold=True, bg=T.bg).pack(side="left")
        ttk.Button(heading, text="Todos os modelos", command=self.catalog).pack(side="right")
        quick = tk.Frame(body, bg=T.bg)
        quick.pack(fill="x", padx=(0, 10), pady=(0, 16))
        for index, (key, title, description) in enumerate((
            ("desdobro", "Desdobro", "Identifique o imóvel e os lotes resultantes."),
            ("retificacao", "Retificação", "Registre os dados da área retificada."),
            ("avaliacao", "Avaliação", "Preencha os dados e o valor do imóvel."),
        )):
            quick.columnconfigure(index, weight=1, uniform="document")
            tile = tk.Frame(quick, bg=T.paper, highlightbackground=T.line, highlightthickness=1,
                            padx=15, pady=13)
            tile.grid(row=0, column=index, sticky="nsew", padx=(0, 8) if index < 2 else 0)
            label(tile, f"0{index + 1}", 9, bold=True, fg=T.accent, bg=T.tint,
                  padx=8, pady=4).pack(anchor="w")
            label(tile, title, 14, bold=True).pack(anchor="w", pady=(8, 3))
            label(tile, description, 9, fg=T.muted, wraplength=185,
                  justify="left").pack(anchor="w", pady=(0, 10))
            ttk.Button(tile, text="Preencher  ›", command=lambda key=key: self.open_kind(key)).pack(anchor="w")
        box = self.card(body, "Arquivo de documentos", "Consulte e abra os PDFs que estão nas pastas de trabalho.")
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
        active = {"desdobro": "Desdobro", "retificacao": "Retificação", "avaliacao": "Avaliação"}.get(model.kind)
        self.page(active, "Emissão de documentos", model.title, "", compact=True)
        footer = tk.Frame(self.main, bg=T.paper, padx=14, pady=6)
        footer.pack(side="bottom", fill="x")
        self.form_status = label(footer, "Campos com * são obrigatórios.", 9, fg=T.muted)
        self.form_status.pack(side="left")
        self.scroll = ScrollArea(self.main)
        self.scroll.pack(fill="both", expand=True, padx=(13, 9), pady=7)
        body = self.scroll.body
        values = json.loads(previous["form"]) if previous else defaults()
        self.form_vars = {}
        entries = {}
        reference = tk.StringVar(value=previous["reference"] if previous else "")
        groups = model.groups()
        available = {key for group in groups.values() for key in group}
        pairs = [(f"{part}Residencial", f"{part}Imovel") for part in ADDRESS_PARTS
                 if {f"{part}Residencial", f"{part}Imovel"} <= available]
        self.address_link = tk.BooleanVar(value=values.get("_EnderecoCompartilhado") == "1") if pairs else None
        if "NumeroDocumento" not in available:
            groups = dict(groups)
            groups["Emissão"] = {"_Referencia": "Número do documento (opcional)", **groups.get("Emissão", {})}
            groups = {"Emissão": groups.pop("Emissão"), **groups}
        for group, fields in groups.items():
            title = "Imóvel de origem" if group == "Imóvel" and model.kind == "desdobro" else group
            box = tk.Frame(body, bg=T.paper, highlightbackground=T.line, highlightthickness=1, padx=11, pady=6)
            box.pack(fill="x", padx=(0, 5), pady=(0, 5))
            label(box, title, 11, bold=True).pack(anchor="w")
            if group == "Endereço residencial" and self.address_link is not None:
                ttk.Checkbutton(box, text="Usar o endereço do contribuinte no imóvel",
                                variable=self.address_link).pack(anchor="w", pady=(3, 0))
            grid = tk.Frame(box, bg=T.paper)
            grid.pack(fill="x", pady=(4, 0))
            for col in range(3):
                grid.columnconfigure(col, weight=1, uniform="form")
            for index, (key, caption) in enumerate(fields.items()):
                row, col = divmod(index, 3)
                variable = reference if key == "_Referencia" else tk.StringVar(value=values.get(key, ""))
                if key != "_Referencia":
                    self.form_vars[key] = variable
                cell = tk.Frame(grid, bg=T.paper)
                cell.grid(row=row, column=col, sticky="ew", padx=(0, 7), pady=(0, 3))
                if key == "LogradouroImovel" and model.kind == "retificacao":
                    caption = "Nome da rua"
                required = " *" if key in REQUIRED.get(model.kind, ()) else ""
                caption_label = label(cell, caption + required, 9, bold=True, wraplength=195, justify="left")
                caption_label.pack(fill="x", pady=(0, 2))
                cell.bind("<Configure>", lambda event, caption=caption_label:
                          caption.configure(wraplength=max(50, event.width)))
                entry = ttk.Entry(cell, textvariable=variable, width=9, style="Form.TEntry")
                entries[key] = entry
                entry.pack(fill="x")
                if key in TAX_ID_FIELDS:
                    mask_tax_id(entry, variable)
        if pairs:
            self.link_address(body, self.address_link, pairs, entries)
        self.fraction_vars = []
        self.active_fractions = 0
        if model.has_fractions():
            rows = json.loads(previous["fractions"]) if previous else [{}, {}]
            box = self.card(body, "Frações resultantes", "Informe a quantidade e preencha o lote, a área e o endereço de cada fração.")
            controls = tk.Frame(box, bg=T.paper)
            controls.pack(fill="x", pady=(2, 10))
            label(controls, "Quantidade de frações", 9, bold=True).pack(side="left")
            self.count = tk.StringVar(value=str(len(rows)))
            self.fraction_table = tk.Frame(box, bg=T.paper)
            spin = ttk.Spinbox(controls, from_=1, to=MAX_FRACTIONS, width=5, textvariable=self.count)
            spin.pack(side="left", padx=8)
            self.count.trace_add("write", self.update_fractions)
            for row in rows:
                self.fraction_vars.append({key: tk.StringVar(value=row.get(key, "")) for key in FRACTION_FIELDS})
            self.fraction_table.pack(fill="x")
            self.update_fractions()
        self.generate_button = ttk.Button(footer, text="Revisar documento  ›", style="Accent.TButton",
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
            label(self.fraction_table, caption, 9, bold=True, bg=T.bg, padx=6, pady=4).grid(row=0, column=col, sticky="ew")
        while len(self.fraction_vars) < quantity:
            self.fraction_vars.append({key: tk.StringVar() for key in FRACTION_FIELDS})
        for index, variables in enumerate(self.fraction_vars[:quantity], 1):
            label(self.fraction_table, f"{index:02d}", 9, bold=True, fg=T.accent, padx=6).grid(row=index, column=0, sticky="w")
            for col, variable in enumerate(variables.values(), 1):
                ttk.Entry(self.fraction_table, textvariable=variable, width=8).grid(row=index, column=col, sticky="ew", padx=3, pady=3)

    def generate(self, model, reference, record_id):
        try:
            if model.has_fractions() and (not self.count.get().isdigit() or int(self.count.get()) != self.active_fractions):
                raise ValueError(f"Informe entre 1 e {MAX_FRACTIONS} frações.")
            values = normalize_fields({key: var.get() for key, var in self.form_vars.items()})
            if self.address_link is not None:
                values["_EnderecoCompartilhado"] = "1" if self.address_link.get() else "0"
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
                number = values.get("NumeroDocumento", reference).strip()
                owner = values.get("NomeProprietario", "")
                stem = pdf_stem(model.kind, model.title, number, owner, record_id)

                def archive():
                    return self.store.save_record(model, values, fractions, data, number, record_id)

                preview = Preview(self.root, data, model.title, self.store.destination(model.kind),
                                  f"{stem}.pdf" if number else "PDF com número automático",
                                  review_details(model, values, fractions, number), archive, self.controls)
                self.root.wait_window(preview)
                if preview.saved:
                    rid, version, path = preview.saved
                    self.archive_done(rid, version, path)
            except Exception as exc:
                self.error(exc)
            finally:
                self.busy = False
                if self.generate_button.winfo_exists():
                    self.generate_button.configure(state="normal")
                    self.form_status.configure(text="")
        self.root.after(90, complete)

    def archive_done(self, record, version, path):
        self.page(None, "Emissão de documentos", "PDF arquivado",
                  "O documento está disponível no histórico e na pasta da categoria.")
        workflow(self.main, 3).pack(fill="x")
        body = self.scroll_page()
        card = self.card(body, path.name, f"TD-{record:06d} · revisão {version:02d}\n{path.parent}")
        row = tk.Frame(card, bg=T.paper)
        row.pack(fill="x")
        ttk.Button(row, text="Abrir PDF", style="Accent.TButton",
                   command=lambda: self.safe(lambda: open_path(path))).pack(side="left")
        ttk.Button(row, text="Abrir pasta", command=lambda: self.safe(
            lambda: open_path(path.parent))).pack(side="left", padx=8)
        ttk.Button(row, text="Ver histórico", command=self.history).pack(side="left")

    def history(self):
        self.page("Histórico", "Arquivo local", "Histórico de documentos",
                  "Pesquise no repositório e selecione um documento para abrir ou corrigir.")
        body = tk.Frame(self.main, bg=T.bg, padx=20, pady=9)
        body.pack(fill="both", expand=True)
        query = tk.StringVar()
        searchbar = tk.Frame(body, bg=T.bg)
        searchbar.pack(fill="x", pady=(0, 6))
        label(searchbar, "Pesquisar", 9, bold=True, bg=T.bg).pack(side="left")
        self.controls.button(searchbar, "Abrir pastas", "folder", lambda: self.safe(
            lambda: open_path(self.store.archive))).pack(side="right")
        self.controls.button(searchbar, "Atualizar", "refresh", lambda: reload_archive()).pack(side="right", padx=5)
        ttk.Entry(searchbar, textvariable=query, width=16).pack(side="left", padx=8, fill="x", expand=True)

        actions = tk.Frame(body, bg=T.bg)
        actions.pack(fill="x", pady=(0, 7))
        open_button = self.controls.button(actions, "Abrir PDF", "open", lambda: selected(show_pdf),
                                           state="disabled")
        open_button.pack(side="left")
        edit_button = self.controls.button(actions, "Corrigir documento", "edit",
                                           lambda: selected(correct), primary=True, state="disabled")
        edit_button.pack(side="left", padx=8)
        status = label(actions, "", 9, fg=T.muted, bg=T.bg)
        status.pack(side="right", padx=(4, 0))

        panel = tk.Frame(body, bg=T.bg)
        panel.pack(fill="both", expand=True)
        columns = [("code", "Código / rev.", 112), ("reference", "Nº / referência", 105),
                   ("owner", "Proprietário", 138), ("model", "Arquivo / tipo", 156), ("date", "Emissão", 85)]
        table = ttk.Treeview(panel, columns=[x[0] for x in columns], show="headings",
                             selectmode="browse", height=3)
        vertical = ttk.Scrollbar(panel, orient="vertical", command=table.yview)
        horizontal = ttk.Scrollbar(panel, orient="horizontal", command=table.xview)
        table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(0, weight=1)
        table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        for key, heading, width in columns:
            table.heading(key, text=heading, anchor="w")
            table.column(key, width=width, minwidth=70, stretch=key in {"owner", "model"})

        rows, documents = {}, []

        def sync_actions(_=None):
            selection = table.selection()
            row = rows.get(selection[0]) if selection else None
            open_button.configure(state="normal" if row else "disabled")
            edit_button.configure(state="normal" if row and row["record_id"] is not None else "disabled")

        def reload_archive():
            try:
                documents[:] = self.store.history()
            except Exception as exc:
                status.configure(text=str(exc))
                return
            refresh()

        def refresh(*_):
            selection = table.selection()
            previous = rows.get(selection[0]) if selection else None
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
                    if previous and previous["path"] == row["path"]:
                        table.selection_set(str(index))
                        table.focus(str(index))
                status.configure(text=f"{len(rows)} PDF(s)")
            except Exception as exc:
                status.configure(text=str(exc))
            sync_actions()

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
            if record is None:
                raise ValueError("A revisão não está mais disponível. Clique em Atualizar.")
            self.form(Model(record["model_id"], record["title"], record["kind"], record["document"]), record)

        def show_pdf(row):
            if not row["path"].is_file():
                raise FileNotFoundError("O PDF não está mais na pasta. Clique em Atualizar.")
            open_path(row["path"])

        query.trace_add("write", refresh)
        table.bind("<<TreeviewSelect>>", sync_actions)
        table.bind("<Return>", lambda _: selected(show_pdf))
        table.bind("<Double-1>", lambda event: selected(show_pdf) if table.identify_row(event.y) else None)
        reload_archive()

    def models(self):
        self.page("Modelos Word", "Gerenciamento", "Modelos Word",
                  "Edite o texto no Word ou Writer. As chaves entre < e > criam campos no formulário.")
        body = self.scroll_page()
        action = tk.Frame(body, bg=T.bg)
        action.pack(fill="x", padx=(0, 10), pady=(0, 10))
        ttk.Button(action, text="Cadastrar modelo .docx", style="Accent.TButton",
                   command=lambda: self.model_editor()).pack(side="left")
        for model in self.store.models():
            box = self.card(body, model.title, f"{len(model.keys())} chaves · arquivo Word editável")
            row = tk.Frame(box, bg=T.paper)
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
        frame = tk.Frame(dialog, bg=T.paper, padx=20, pady=16)
        frame.pack(fill="both", expand=True)
        label(frame, "Modelo do documento", 15, bold=True).pack(anchor="w")
        label(frame, "Nome do modelo", 9, bold=True).pack(anchor="w", pady=(12, 4))
        title = tk.StringVar(value=model.title if model else "")
        ttk.Entry(frame, textvariable=title).pack(fill="x")
        label(frame, "Edite frases e formatação no Word ou Writer. Para criar um campo, digite uma chave como <CimImovel>.",
              fg=T.muted, wraplength=530, justify="left").pack(anchor="w", pady=(10, 6))
        working = TemporaryDirectory(prefix="modelo-", dir=self.store.root, ignore_cleanup_errors=True)
        editable = Path(working.name) / "Modelo.docx"
        if model:
            editable.write_bytes(model.document)
        source = [editable if model else None]

        def close():
            working.cleanup()
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", close)
        box = tk.Frame(frame, bg=T.paper)
        box.pack(fill="x")
        status = label(frame, "", fg=T.muted, wraplength=530, justify="left")
        footer = tk.Frame(frame, bg=T.paper)
        footer.pack(side="bottom", fill="x", pady=(8, 0))
        listframe = tk.Frame(frame, bg=T.paper)
        keys = tk.Listbox(listframe, height=7, bg="#f5f9fa", fg=T.ink, relief="flat",
                          highlightthickness=1, highlightbackground=T.line, exportselection=False, font=(FONT, 10))
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

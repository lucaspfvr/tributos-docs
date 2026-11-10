from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tributos.docx import export_pdf
from tributos.models import ASSETS, FRACTION_FIELDS, Model, defaults, prepare

ASSET_DIR = Path(__file__).parent / "tributos" / "assets"


def main():
    window = tk.Tk()
    window.title("Tributos Docs")
    window.geometry("700x660")
    selected = tk.StringVar(value=ASSETS["avaliacao"])
    selector = ttk.Combobox(window, textvariable=selected, values=list(ASSETS.values()), state="readonly")
    selector.pack(fill="x", padx=16, pady=14)

    content = ttk.Frame(window)
    content.pack(fill="both", expand=True, padx=16)
    canvas = tk.Canvas(content, highlightthickness=0)
    bar = ttk.Scrollbar(content, command=canvas.yview)
    canvas.configure(yscrollcommand=bar.set)
    bar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    form = ttk.Frame(canvas, padding=8)
    pane = canvas.create_window((0, 0), window=form, anchor="nw")
    form.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(pane, width=event.width))
    actions = ttk.Frame(window, padding=16)
    actions.pack(fill="x")

    def show_form(_=None):
        for child in (*form.winfo_children(), *actions.winfo_children()):
            child.destroy()
        kind = next(key for key, title in ASSETS.items() if title == selected.get())
        model = Model(kind, ASSETS[kind], kind, (ASSET_DIR / f"{kind}.docx").read_bytes())
        initial = defaults()
        variables = {}
        fraction_rows = []

        def add_field(parent, key, caption, target, value=""):
            row = ttk.Frame(parent)
            row.pack(fill="x", padx=8, pady=2)
            ttk.Label(row, text=caption).pack(side="left")
            target[key] = tk.StringVar(value=value)
            ttk.Entry(row, textvariable=target[key], width=28).pack(side="right")

        for title, fields in model.groups().items():
            group = ttk.LabelFrame(form, text=title, padding=6)
            group.pack(fill="x", pady=4)
            for key, caption in fields.items():
                add_field(group, key, caption, variables, initial.get(key, ""))

        if model.has_fractions():
            group = ttk.LabelFrame(form, text="Frações do desdobro", padding=6)
            group.pack(fill="x", pady=4)

            def add_fraction():
                if len(fraction_rows) == 99:
                    return
                row = ttk.LabelFrame(group, text=f"Fração {len(fraction_rows) + 1}")
                row.pack(fill="x", pady=4, before=add_button)
                values = {}
                for key, caption in FRACTION_FIELDS.items():
                    add_field(row, key, caption, values)
                fraction_rows.append(values)

            add_button = ttk.Button(group, text="Adicionar fração", command=add_fraction)
            add_button.pack(anchor="w", pady=6)
            add_fraction()

        def save(kind):
            try:
                fields = {key: variable.get() for key, variable in variables.items()}
                fractions = [{key: value.get() for key, value in row.items()} for row in fraction_rows]
                filled = prepare(model, fields, fractions)
                target = filedialog.asksaveasfilename(
                    parent=window, defaultextension=".docx", filetypes=[("Documento Word", "*.docx")]
                )
                if target:
                    Path(target).write_bytes(filled)
                    messagebox.showinfo("Tributos Docs", "Documento salvo.", parent=window)
            except (OSError, ValueError) as exc:
                messagebox.showerror("Tributos Docs", str(exc), parent=window)

        ttk.Button(actions, text="Salvar Word", command=lambda: save("docx")).pack(side="left")
        ttk.Button(actions, text="Salvar PDF", command=lambda: save("pdf")).pack(side="left", padx=8)
        canvas.yview_moveto(0)

    selector.bind("<<ComboboxSelected>>", show_form)
    show_form()
    window.mainloop()


if __name__ == "__main__":
    main()

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tributos.models import Model, defaults, prepare


def main():
    model_path = Path(__file__).parent / "tributos" / "assets" / "avaliacao.docx"
    model = Model("avaliacao", "Avaliação de imóvel", "avaliacao", model_path.read_bytes())
    window = tk.Tk()
    window.title("Tributos Docs")
    window.geometry("550x385")
    panel = ttk.Frame(window, padding=24)
    panel.pack(fill="both", expand=True)
    ttk.Label(panel, text=model.title, font=("Segoe UI", 20, "bold")).pack(anchor="w", pady=(0, 14))
    fields = {
        "NumeroDocumento": ("Nº do documento", "001"),
        "NomeProprietario": ("Nome do proprietário", "PESSOA DE TESTE"),
        "ValorAvaliacao": ("Valor da avaliação (R$)", "1500,00"),
        "ValorExtenso": ("Valor por extenso", "Mil e quinhentos reais"),
    }
    variables = {}
    for key, (caption, value) in fields.items():
        ttk.Label(panel, text=caption).pack(anchor="w")
        variables[key] = tk.StringVar(value=value)
        ttk.Entry(panel, textvariable=variables[key]).pack(fill="x", pady=(0, 9))

    def save():
        values = defaults()
        values.update({key: var.get() for key, var in variables.items()})
        try:
            filled = prepare(model, values, [])
            target = filedialog.asksaveasfilename(
                parent=window, defaultextension=".docx", filetypes=[("Documento Word", "*.docx")]
            )
            if target:
                Path(target).write_bytes(filled)
                messagebox.showinfo("Tributos Docs", "Documento Word salvo.", parent=window)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Tributos Docs", str(exc), parent=window)

    ttk.Button(panel, text="Salvar Word", command=save).pack(anchor="w", pady=12)
    window.mainloop()


if __name__ == "__main__":
    main()

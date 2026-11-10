import tkinter as tk
from tkinter import ttk


def main():
    window = tk.Tk()
    window.title("Tributos Docs")
    window.geometry("520x280")
    panel = ttk.Frame(window, padding=24)
    panel.pack(fill="both", expand=True)
    ttk.Label(panel, text="Tributos Docs", font=("Segoe UI", 22, "bold")).pack(anchor="w")
    ttk.Label(panel, text="Automação de documentos tributários").pack(anchor="w", pady=(4, 20))
    ttk.Label(panel, text="Base da aplicação desktop").pack(anchor="w")
    ttk.Button(panel, text="Fechar", command=window.destroy).pack(anchor="w", pady=16)
    window.mainloop()


if __name__ == "__main__":
    main()

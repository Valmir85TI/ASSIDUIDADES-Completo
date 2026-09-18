"""Componentes visuais reutilizados pelo aplicativo."""

import tkinter as tk
from tkinter import ttk


def create_loading_dialog(parent: tk.Misc) -> tuple[tk.Toplevel, ttk.Progressbar]:
    """Cria uma janela de progresso consistente com a tela principal."""
    dialog = tk.Toplevel(parent)
    dialog.title("Gerando documentos")
    dialog.configure(bg="#f8fafc")
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()

    width, height = 470, 245
    x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    dialog.geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")

    header = tk.Frame(dialog, bg="#0f2a48", height=86)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="Gerando documentos", bg="#0f2a48", fg="white",
             font=("Segoe UI", 18, "bold")).pack(anchor="w", padx=28, pady=(18, 2))
    tk.Label(header, text="Aguarde enquanto os arquivos são preparados", bg="#0f2a48",
             fg="#cbd5e1", font=("Segoe UI", 10)).pack(anchor="w", padx=28)

    content = tk.Frame(dialog, bg="#f8fafc")
    content.pack(fill="both", expand=True, padx=28, pady=22)
    tk.Label(content, text="Progresso da geração", bg="#f8fafc", fg="#475569",
             font=("Segoe UI", 10, "bold")).pack(anchor="w")

    style = ttk.Style(dialog)
    style.configure("Generation.Horizontal.TProgressbar", troughcolor="#e2e8f0",
                    background="#1f6fc9", bordercolor="#e2e8f0", lightcolor="#1f6fc9",
                    darkcolor="#1f6fc9", thickness=18)
    progress = ttk.Progressbar(content, orient="horizontal", mode="determinate",
                               style="Generation.Horizontal.TProgressbar", maximum=1)
    progress.pack(fill="x", pady=(10, 12))

    tk.Label(content, text="Não feche esta janela durante o processo.", bg="#f8fafc",
             fg="#64748b", font=("Segoe UI", 9)).pack(anchor="w")
    dialog.update_idletasks()
    return dialog, progress

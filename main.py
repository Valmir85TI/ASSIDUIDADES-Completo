"""Aplicativo para geração de documentos de assiduidade e benefícios."""

from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from Assiduidade import assi
from Assiduidade_Castanhao import castanhao
from Autorizacao_Trabalho import trabalho
from Vale_Natal import natal
from Vale_Natal_Castanhao import natalCastanhao


def application_directory() -> Path:
    """Retorna a pasta que contém as planilhas e documentos do usuário."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_directory() -> Path:
    """Retorna a pasta de recursos, inclusive quando executado pelo PyInstaller."""
    return Path(getattr(sys, "_MEIPASS", application_directory()))


class Aplicacao(assi, trabalho, natal, castanhao, natalCastanhao):
    NAVY = "#0f2a48"
    BACKGROUND = "#f1f5f9"
    PRIMARY = "#1f6fc9"
    MUTED = "#64748b"

    def __init__(self) -> None:
        self.base_dir = application_directory()
        self.resource_dir = resource_directory()
        os.chdir(self.base_dir)

        self.janela = tk.Tk()
        self._configure_window()
        self._configure_style()
        self._build_interface()
        self.janela.mainloop()

    def _configure_window(self) -> None:
        self.janela.title("Assiduidades | Castanha")
        self.janela.configure(bg=self.BACKGROUND)
        self.janela.resizable(False, False)
        self._center_window(760, 620)

        logo_path = self.resource_dir / "Logo Castanha2.png"
        if logo_path.exists():
            self.logo = tk.PhotoImage(file=logo_path)
            self.janela.iconphoto(True, self.logo)

    def _configure_style(self) -> None:
        style = ttk.Style(self.janela)
        style.theme_use("clam")
        style.configure("Action.TButton", font=("Segoe UI", 11, "bold"),
                        foreground="white", background=self.PRIMARY,
                        padding=(18, 12), borderwidth=0)
        style.map("Action.TButton",
                  background=[("active", "#1859a2"), ("disabled", "#94a3b8")])
        style.configure("Status.TLabel", font=("Segoe UI", 10),
                        foreground=self.MUTED, background="white")

    def _build_interface(self) -> None:
        header = tk.Frame(self.janela, bg=self.NAVY, height=142)
        header.pack(fill="x")
        header.pack_propagate(False)
        header_content = tk.Frame(header, bg=self.NAVY)
        header_content.pack(fill="both", expand=True, padx=42)

        brand = tk.Frame(header_content, bg=self.NAVY)
        brand.pack(side="right", padx=(24, 0), pady=24)
        if hasattr(self, "logo"):
            tk.Label(brand, image=self.logo, bg=self.NAVY).pack()

        heading = tk.Frame(header_content, bg=self.NAVY)
        heading.pack(side="left", fill="y")
        tk.Label(heading, text="Gerador de documentos", bg=self.NAVY, fg="white",
                 font=("Segoe UI", 25, "bold")).pack(anchor="w", pady=(32, 3))
        tk.Label(heading, text="Assiduidade, vale natalino e autorizações de trabalho",
                 bg=self.NAVY, fg="#cbd5e1", font=("Segoe UI", 12)).pack(anchor="w")

        content = tk.Frame(self.janela, bg="white", highlightbackground="#e2e8f0",
                           highlightthickness=1)
        content.pack(fill="both", expand=True, padx=38, pady=(24, 20))

        tk.Label(content, text="SELECIONE O DOCUMENTO", bg="white", fg=self.MUTED,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=28, pady=(24, 10))

        actions = (
            ("Assiduidade Castanha", "ASSIDUIDADE.xlsx", "ASSIDUIDADE.docx", self.funcao_a_executar),
            ("Assiduidade Castanhão", "ASSIDUIDADE_CASTANHAO.xlsx", "ASSIDUIDADE_CASTANHAO.docx", self.funcao_a_executar4),
            ("Vale Natal", "VALENATAL.xlsx", "VALENATAL.docx", self.funcao_a_executar2),
            ("Vale Natal Castanhão", "VALENATAL_CASTANHAO.xlsx", "VALENATAL_CASTANHAO.docx", self.funcao_a_executar5),
            ("Autorização de Trabalho", "ADITAMENTO.xlsx", "ADITAMENTO.docx", self.funcao_a_executar3),
        )
        self.action_buttons: list[ttk.Button] = []
        for label, input_name, output_name, callback in actions:
            button = ttk.Button(
                content,
                text=label,
                style="Action.TButton",
                command=lambda i=input_name, o=output_name, c=callback, l=label: self._run_generator(l, i, o, c),
            )
            button.pack(fill="x", padx=28, pady=5)
            self.action_buttons.append(button)

        self.status = ttk.Label(content, text="As planilhas devem estar na mesma pasta do programa.",
                                style="Status.TLabel")
        self.status.pack(anchor="w", padx=28, pady=(16, 20))

    def _run_generator(self, label: str, input_name: str, output_name: str, callback) -> None:
        source = self.base_dir / input_name
        if not source.exists():
            messagebox.showerror(
                "Arquivo não encontrado",
                f"Não foi encontrada a planilha necessária:\n{source.name}\n\n"
                "Coloque o arquivo na mesma pasta do programa e tente novamente.",
                parent=self.janela,
            )
            return

        self._set_busy(True, f"Gerando {label}...")
        self.janela.update_idletasks()
        try:
            callback()
            output = self.base_dir / output_name
            if output.exists():
                self.status.configure(text=f"Concluído: {output.name}")
        except Exception as error:
            self.status.configure(text="Não foi possível gerar o documento.")
            messagebox.showerror(
                "Erro na geração",
                f"Não foi possível gerar {label}.\n\nDetalhes: {error}",
                parent=self.janela,
            )
        finally:
            self._set_busy(False)

    def _set_busy(self, busy: bool, text: str | None = None) -> None:
        state = "disabled" if busy else "normal"
        for button in self.action_buttons:
            button.configure(state=state)
        if text:
            self.status.configure(text=text)

    def _center_window(self, width: int, height: int) -> None:
        x = (self.janela.winfo_screenwidth() - width) // 2
        y = (self.janela.winfo_screenheight() - height) // 2
        self.janela.geometry(f"{width}x{height}+{x}+{y}")


if __name__ == "__main__":
    Aplicacao()

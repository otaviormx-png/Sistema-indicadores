from __future__ import annotations

import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


def _run_editor() -> None:
    from aps_clonador_interativo import launch_editor

    root = tk.Tk()
    root.withdraw()
    win = launch_editor(master=root)
    if win is None:
        root.destroy()
        return
    win.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()


def _run_dashboard() -> None:
    from aps_dashboard import launch_dashboard

    root = tk.Tk()
    root.withdraw()
    launch_dashboard()
    root.mainloop()


def _run_aprazador() -> None:
    from aps_aprazamento import main as run_aprazamento_main

    run_aprazamento_main()


def _spawn_tool(tool: str) -> None:
    """Open each module in a separate process to keep windows independent."""
    tool = str(tool or "").strip().lower()
    if tool not in {"dashboard", "editor", "aprazador"}:
        return

    if getattr(sys, "frozen", False):
        cmd = [sys.executable, "--tool", tool]
    else:
        cmd = [sys.executable, str(Path(__file__).resolve()), "--tool", tool]

    subprocess.Popen(cmd)


def _dispatch_tool(tool: str) -> int:
    tool = str(tool or "").strip().lower()
    if tool == "dashboard":
        _run_dashboard()
        return 0
    if tool == "editor":
        _run_editor()
        return 0
    if tool == "aprazador":
        _run_aprazador()
        return 0
    return 1


class APSLiteHub(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("APS 3 em 1")
        self.geometry("720x400")
        self.configure(bg="#EEF4F8")
        self.resizable(False, False)

        try:
            self.iconbitmap(default="APS_Suite.ico")
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self) -> None:
        title_wrap = tk.Frame(self, bg="#EEF4F8")
        title_wrap.pack(fill="x", padx=18, pady=(18, 10))

        tk.Label(
            title_wrap,
            text="APS 3 EM 1",
            bg="#1F4E79",
            fg="white",
            font=("Segoe UI", 16, "bold"),
            pady=12,
        ).pack(fill="x")

        tk.Label(
            self,
            text="Escolha o modulo para abrir:\nDashboard Completo, Editor Completo ou Aprazador Completo.",
            bg="#EEF4F8",
            fg="#1F4E79",
            font=("Segoe UI", 11),
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=18, pady=(6, 8))

        cards = tk.Frame(self, bg="#EEF4F8")
        cards.pack(fill="both", expand=True, padx=18, pady=(4, 12))
        for c in range(3):
            cards.columnconfigure(c, weight=1)

        self._module_card(
            cards,
            0,
            "Dashboard Completo",
            "Painel completo com comparacoes,\ngraficos e exportacoes.",
            "#DCEAF5",
            "dashboard",
        )
        self._module_card(
            cards,
            1,
            "Editor Completo",
            "Abre direto no editor completo\n(sem tela de gerador/clonador).",
            "#FFF4E5",
            "editor",
        )
        self._module_card(
            cards,
            2,
            "Aprazador",
            "Controle completo de\naprazamento e semaforo.",
            "#EAF7EA",
            "aprazador",
        )

        footer = tk.Frame(self, bg="#EEF4F8")
        footer.pack(fill="x", padx=18, pady=(0, 14))

        tk.Button(
            footer,
            text="Abrir os 3 modulos",
            command=self._open_all,
            bg="#1F4E79",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=6,
        ).pack(side="left")

        tk.Button(
            footer,
            text="Sair",
            command=self.destroy,
            font=("Segoe UI", 10),
            padx=14,
            pady=6,
        ).pack(side="right")

    def _module_card(self, parent: tk.Widget, col: int, title: str, desc: str, bg: str, tool: str) -> None:
        box = tk.Frame(parent, bg=bg, bd=1, relief="solid")
        box.grid(row=0, column=col, sticky="nsew", padx=6, pady=6)

        tk.Label(
            box,
            text=title,
            bg=bg,
            fg="#1F4E79",
            font=("Segoe UI", 12, "bold"),
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=10, pady=(10, 6))

        tk.Label(
            box,
            text=desc,
            bg=bg,
            fg="#1F1F1F",
            font=("Segoe UI", 10),
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=10, pady=(0, 12))

        tk.Button(
            box,
            text="Abrir",
            command=lambda t=tool: _spawn_tool(t),
            bg="#FFFFFF",
            fg="#1F4E79",
            font=("Segoe UI", 10, "bold"),
            padx=16,
            pady=6,
        ).pack(anchor="w", padx=10, pady=(0, 10))

    def _open_all(self) -> None:
        try:
            _spawn_tool("dashboard")
            _spawn_tool("editor")
            _spawn_tool("aprazador")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc), parent=self)


def main() -> int:
    argv = list(sys.argv[1:])
    if len(argv) >= 2 and argv[0] == "--tool":
        return _dispatch_tool(argv[1])

    app = APSLiteHub()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

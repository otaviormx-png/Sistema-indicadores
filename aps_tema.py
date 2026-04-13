"""
aps_tema.py — Adaptador de tema moderno (opcional).

Se CustomTkinter estiver instalado, aplica tema flat moderno com dark mode.
Se não estiver, cai silenciosamente de volta ao Tkinter padrão.

Instalação (opcional):
    pip install customtkinter

Como ativar:
    Em config.toml, adicione:
      [interface]
      tema_moderno = true

Ou defina a variável de ambiente:
    APS_TEMA_MODERNO=1
"""
from __future__ import annotations

import os

import aps_config

_ATIVO = False
_CTK = None


def _deve_usar() -> bool:
    env = os.environ.get("APS_TEMA_MODERNO", "").strip()
    if env in {"1", "true", "yes"}:
        return True
    return aps_config.get("interface", "tema_moderno", False) is True


def init() -> bool:
    """
    Tenta inicializar o CustomTkinter.
    Retorna True se ativo, False se Tkinter padrão será usado.
    """
    global _ATIVO, _CTK
    if not _deve_usar():
        return False
    try:
        import customtkinter as ctk  # type: ignore
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        _CTK = ctk
        _ATIVO = True
        return True
    except ImportError:
        return False


def ativo() -> bool:
    return _ATIVO


def patch_style(style_obj) -> None:
    """
    Aplica ajustes de estilo ao ttk.Style quando CustomTkinter não está disponível,
    tornando a interface padrão um pouco mais limpa.
    """
    if _ATIVO:
        return
    try:
        style_obj.theme_use("clam")
        style_obj.configure(".", font=("Segoe UI", 10))
        style_obj.configure("TButton",
                            padding=6,
                            relief="flat",
                            background="#E8F0FE",
                            foreground="#1F4E79")
        style_obj.map("TButton",
                      background=[("active", "#C9DAF8"), ("pressed", "#A8C7FA")])
        style_obj.configure("TLabelframe", borderwidth=1)
        style_obj.configure("TEntry", padding=4)
        style_obj.configure("Treeview", rowheight=26)
    except Exception:
        pass


def get_module():
    """Retorna o módulo customtkinter se ativo, None caso contrário."""
    return _CTK

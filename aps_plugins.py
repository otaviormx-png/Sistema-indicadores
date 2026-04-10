"""
aps_plugins.py — Descoberta automática de indicadores na pasta plugins/.

Como criar um novo indicador (ex: C8):
  1. Crie o arquivo  plugins/c8_meu_indicador.py
  2. Defina nele:
       CFG = IndicatorConfig(code="C8", ...)
       def processar(entrada, saida): ...
  3. Reinicie o sistema — C8 aparecerá automaticamente na interface.

Não é necessário alterar sistema_aps.py nem aps_interface.py.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


_PLUGINS_DIR = Path(__file__).parent / "plugins"


def _load_plugin(path: Path):
    """Importa um .py como módulo e retorna o objeto módulo."""
    spec = importlib.util.spec_from_file_location(path.stem, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[path.stem] = mod
    spec.loader.exec_module(mod)
    return mod


def load_all() -> list[tuple]:
    """
    Varre plugins/ em ordem alfabética.
    Retorna lista de (CFG, processar) para cada plugin válido.
    """
    if not _PLUGINS_DIR.exists():
        return []

    result = []
    for path in sorted(_PLUGINS_DIR.glob("*.py")):
        if path.name.startswith("_"):
            continue
        try:
            mod = _load_plugin(path)
            cfg = getattr(mod, "CFG", None)
            processar = getattr(mod, "processar", None)
            if cfg is not None and callable(processar):
                result.append((cfg, processar))
        except Exception as exc:
            print(f"[aps_plugins] Erro ao carregar {path.name}: {exc}")
    return result

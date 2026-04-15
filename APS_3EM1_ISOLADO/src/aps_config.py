"""
aps_config.py — Leitor centralizado de configuração.

Carrega config.toml na mesma pasta do executável/script.
Se o arquivo não existir, usa valores padrão embutidos.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Python 3.11+ tem tomllib nativo; antes usa tomli (instalado como dependência).
try:
    import tomllib  # type: ignore
except ModuleNotFoundError:
    try:
        import tomli as tomllib  # type: ignore
    except ModuleNotFoundError:
        tomllib = None  # type: ignore


# ------------------------------------------------------------------
# Valores padrão (usados quando config.toml não existe ou está incompleto)
# ------------------------------------------------------------------
_DEFAULTS: dict = {
    "caminhos": {
        "entrada_padrao": "",
        "saida_padrao": "",
    },
    "processamento": {
        "ignorar_marcadores": [
            "resultado",
            "aps_resultados",
            "busca ativa",
            "estatisticas",
            "estatísticas",
            "resumo",
            "_interativa",
            "comparativo",
            "dashboard",
        ],
    },
    "cores": {
        "azul_escuro": "1F4E79",
        "azul_medio": "2E75B6",
        "azul_claro": "D6E4F0",
        "azul_header": "BDD7EE",
        "verde_ok": "C6EFCE",
        "verde_texto": "276221",
        "verde_escuro": "375623",
        "amarelo": "FFEB9C",
        "amarelo_txt": "9C5700",
        "vermelho": "FFC7CE",
        "vermelho_txt": "9C0006",
        "laranja": "F4B942",
        "laranja_txt": "843C0C",
        "cinza_claro": "F2F2F2",
        "cinza_medio": "D9D9D9",
        "cinza_escuro": "595959",
        "branco": "FFFFFF",
        "preto": "000000",
        "roxo": "7030A0",
        "roxo_claro": "E8D5F5",
        "amarelo_header": "FFE599",
        "verde_header": "E2EFDA",
        "azul_clinico": "C9DAF8",
    },
}


def _config_path() -> Path:
    """Localiza config.toml ao lado do executável ou do script."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    return base / "config.toml"


def _deep_merge(base: dict, override: dict) -> dict:
    """Mescla override sobre base, respeitando subchaves."""
    result = dict(base)
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(result.get(k), dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def _load() -> dict:
    path = _config_path()
    if not path.exists() or tomllib is None:
        return _DEFAULTS

    try:
        with open(path, "rb") as f:
            raw = tomllib.load(f)
        return _deep_merge(_DEFAULTS, raw)
    except Exception:
        return _DEFAULTS


# Carrega uma vez na importação.
_CFG = _load()


# ------------------------------------------------------------------
# Acesso público
# ------------------------------------------------------------------

def get(section: str, key: str, default=None):
    """Retorna _CFG[section][key], ou default se não existir."""
    return _CFG.get(section, {}).get(key, default)


def cores() -> dict:
    return dict(_CFG.get("cores", _DEFAULTS["cores"]))


def ignorar_marcadores() -> list[str]:
    return list(_CFG.get("processamento", {}).get(
        "ignorar_marcadores",
        _DEFAULTS["processamento"]["ignorar_marcadores"],
    ))


def entrada_padrao() -> str:
    return _CFG.get("caminhos", {}).get("entrada_padrao", "")


def saida_padrao() -> str:
    return _CFG.get("caminhos", {}).get("saida_padrao", "")

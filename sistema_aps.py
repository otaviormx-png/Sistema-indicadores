
from __future__ import annotations

import hashlib
import json
import traceback
from pathlib import Path
from datetime import datetime
from typing import Callable

import aps_config
import aps_plugins
from aps_utils import candidate_file_for_indicator, normalize_text

from c1_mais_acesso import CFG as C1_CFG, processar as processar_c1
from c2_infantil import CFG as C2_CFG, processar as processar_c2
from c3_gestacao import CFG as C3_CFG, processar as processar_c3
from c4_diabetes import CFG as C4_CFG, processar as processar_c4
from c5_hipertensao import CFG as C5_CFG, processar as processar_c5
from c6_idoso import CFG as C6_CFG, processar as processar_c6
from c7_mulher import CFG as C7_CFG, processar as processar_c7


def _md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _cache_path(out_dir: Path) -> Path:
    return out_dir / ".aps_cache.json"


def _load_cache(out_dir: Path) -> dict:
    p = _cache_path(out_dir)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_cache(out_dir: Path, cache: dict) -> None:
    try:
        _cache_path(out_dir).write_text(
            json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _resolve_default_path(config_val: str, fallback: Path) -> Path:
    """Usa config_val se não estiver vazio, senão usa fallback."""
    if config_val and config_val.strip():
        return Path(config_val.strip())
    return fallback


ROOT_DESKTOP = Path.home() / "Desktop"
OUT_DIR = _resolve_default_path(
    aps_config.saida_padrao(),
    ROOT_DESKTOP / "APS_RESULTADOS",
)
OUT_DIR.mkdir(parents=True, exist_ok=True)

_entrada_config = aps_config.entrada_padrao()
ROOT_INPUT = _resolve_default_path(_entrada_config, ROOT_DESKTOP)

INDICADORES = [
    (C1_CFG, processar_c1),
    (C2_CFG, processar_c2),
    (C3_CFG, processar_c3),
    (C4_CFG, processar_c4),
    (C5_CFG, processar_c5),
    (C6_CFG, processar_c6),
    (C7_CFG, processar_c7),
] + aps_plugins.load_all()  # carrega plugins da pasta plugins/ automaticamente

# Carrega marcadores de ignorar do config
IGNORE_MARKERS = aps_config.ignorar_marcadores()


def desktop_files(root: Path | None = None) -> list[Path]:
    root = Path(root) if root else ROOT_INPUT
    files = []
    for f in root.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
            continue
        name = normalize_text(f.name)
        full = normalize_text(str(f))
        if any(marker in name or marker in full for marker in IGNORE_MARKERS):
            continue
        files.append(f)
    return files


def get_indicators():
    return INDICADORES


def processar_indicador(
    cfg,
    func: Callable,
    files: list[Path],
    out_dir: Path,
    stamp: str | None = None,
) -> Path | None:
    entrada = candidate_file_for_indicator(files, cfg)
    if not entrada:
        return None
    stamp = stamp or datetime.now().strftime("%Y%m%d_%H%M%S")
    saida = out_dir / f"{cfg.code}_{stamp}.xlsx"
    func(entrada, saida)
    return saida


def processar_todos(
    root: Path | None = None,
    out_dir: Path | None = None,
    logger: Callable[[str], None] | None = None,
) -> list[dict]:
    """
    Processa todos os indicadores.

    CORREÇÃO: cada indicador é isolado em try/except.
    Um erro em C3 não cancela C4–C7.
    """
    root = Path(root) if root else ROOT_INPUT
    out_dir = Path(out_dir) if out_dir else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    log = logger or (lambda msg: print(msg))
    files = desktop_files(root)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    resultados = []
    log(f"Arquivos elegíveis: {len(files)}")

    for cfg, func in INDICADORES:
        entrada = candidate_file_for_indicator(files, cfg)
        if not entrada:
            log(f"{cfg.code}: bruto não encontrado.")
            resultados.append({
                "code": cfg.code,
                "status": "não encontrado",
                "entrada": None,
                "saida": None,
                "erro": None,
            })
            continue

        saida = out_dir / f"{cfg.code}_{stamp}.xlsx"
        log(f"Início do processamento: {cfg.code}")
        log(f"  Entrada: {entrada}")
        log(f"  Saída:   {saida}")

        try:
            func(entrada, saida)
            log(f"  ✔ Concluído: {cfg.code}")
            resultados.append({
                "code": cfg.code,
                "status": "ok",
                "entrada": entrada,
                "saida": saida,
                "erro": None,
            })
        except Exception:
            tb = traceback.format_exc()
            log(f"  ✘ Erro em {cfg.code}:\n{tb}")
            resultados.append({
                "code": cfg.code,
                "status": "erro",
                "entrada": entrada,
                "saida": None,
                "erro": tb,
            })

    return resultados


def process_selected(
    selected_codes: list[str],
    in_dir: Path | None = None,
    out_dir: Path | None = None,
    log: Callable[[str], None] | None = None,
    use_cache: bool = True,
) -> list[dict]:
    """
    Processa apenas os indicadores selecionados.
    Se use_cache=True e o arquivo de entrada não mudou desde o último
    processamento (verificação por MD5), reutiliza o xlsx existente.
    """
    in_dir = Path(in_dir) if in_dir else ROOT_INPUT
    out_dir = Path(out_dir) if out_dir else OUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    _log = log or (lambda msg: print(msg))
    files = desktop_files(in_dir)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    resultados = []
    cache = _load_cache(out_dir) if use_cache else {}
    cache_updated = False
    _log(f"Arquivos elegíveis: {len(files)}")

    for cfg, func in INDICADORES:
        if cfg.code not in selected_codes:
            continue

        entrada = candidate_file_for_indicator(files, cfg)
        if not entrada:
            _log(f"{cfg.code}: bruto não encontrado.")
            resultados.append({
                "code": cfg.code, "status": "não encontrado",
                "entrada": None, "saida": None, "erro": None,
            })
            continue

        # Verifica cache
        if use_cache:
            try:
                file_hash = _md5(entrada)
                cache_key = cfg.code
                cached = cache.get(cache_key, {})
                cached_saida = Path(cached.get("saida", "")) if cached.get("saida") else None
                if (cached.get("hash") == file_hash
                        and cached_saida and cached_saida.exists()):
                    _log(f"{cfg.code}: arquivo inalterado — reutilizando {cached_saida.name}")
                    resultados.append({
                        "code": cfg.code, "status": "ok",
                        "entrada": entrada, "saida": cached_saida, "erro": None,
                        "cache_hit": True,
                    })
                    continue
            except Exception:
                pass  # falha no cache: processa normalmente

        saida = out_dir / f"{cfg.code}_{stamp}.xlsx"
        _log(f"Início do processamento: {cfg.code}")
        _log(f"  Entrada: {entrada}")
        _log(f"  Saída:   {saida}")

        try:
            func(entrada, saida)
            _log(f"  ✔ Concluído: {cfg.code}")
            if use_cache:
                try:
                    cache[cfg.code] = {"hash": _md5(entrada), "saida": str(saida)}
                    cache_updated = True
                except Exception:
                    pass
            resultados.append({
                "code": cfg.code, "status": "ok",
                "entrada": entrada, "saida": saida, "erro": None,
            })
        except Exception:
            tb = traceback.format_exc()
            _log(f"  ✘ Erro em {cfg.code}:\n{tb}")
            resultados.append({
                "code": cfg.code, "status": "erro",
                "entrada": entrada, "saida": None, "erro": tb,
            })

    if cache_updated:
        _save_cache(out_dir, cache)

    return resultados


def main():
    processar_todos()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        erro = traceback.format_exc()
        try:
            (ROOT_DESKTOP / "APS_erro_log.txt").write_text(erro, encoding="utf-8")
        except Exception:
            pass
        raise

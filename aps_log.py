"""
aps_log.py — Log persistente rotativo.

Salva cada execução em APS_log.txt ao lado da pasta de saída.
Máximo de MAX_LINES linhas; ao ultrapassar, as mais antigas são removidas.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

MAX_LINES = 500


def _log_path(out_dir: Path) -> Path:
    return out_dir / "APS_log.txt"


def _read_lines(path: Path) -> list[str]:
    if not path.exists():
        return []
    try:
        return path.read_text(encoding="utf-8").splitlines()
    except Exception:
        return []


def _write(path: Path, lines: list[str]) -> None:
    try:
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except Exception:
        pass


def append(out_dir: Path, msg: str) -> None:
    """Acrescenta uma linha ao log, rotacionando se necessário."""
    path = _log_path(out_dir)
    lines = _read_lines(path)
    lines.append(msg)
    if len(lines) > MAX_LINES:
        lines = lines[-MAX_LINES:]
    _write(path, lines)


def log_session_start(out_dir: Path, selected_codes: list[str]) -> None:
    ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    append(out_dir, "")
    append(out_dir, f"=== SESSÃO {ts} — indicadores: {', '.join(selected_codes)} ===")


def log_result(out_dir: Path, results: list[dict]) -> None:
    for r in results:
        ts = datetime.now().strftime("%H:%M:%S")
        status = r.get("status", "?")
        code = r.get("code", "?")
        if status == "ok":
            saida = Path(r["saida"]).name if r.get("saida") else "-"
            append(out_dir, f"  [{ts}] {code}: OK → {saida}")
        elif status == "erro":
            first_line = (r.get("erro") or "").splitlines()[-1] if r.get("erro") else "erro desconhecido"
            append(out_dir, f"  [{ts}] {code}: ERRO — {first_line}")
        else:
            append(out_dir, f"  [{ts}] {code}: não encontrado")

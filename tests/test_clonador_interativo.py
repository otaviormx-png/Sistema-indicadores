from __future__ import annotations

from pathlib import Path

import openpyxl

from aps_clonador_interativo import clone_interactive, refresh_interactive_workbook


def _make_base_workbook(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "📋 Dados C1"

    headers = [
        "Nome",
        "Microárea",
        "Telefone celular",
        "Telefone residencial",
        "A - Consulta em dia",
        "B - Visita em dia",
    ]
    for c, value in enumerate(headers, 1):
        ws.cell(3, c, value)

    rows = [
        ["Carlos", "003", "333", "", "SIM", "SIM"],
        ["Bruno", "002", "222", "", "", ""],
        ["Ana", "001", "111", "", "SIM", ""],
    ]
    for r, row in enumerate(rows, 4):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)

    wb.save(path)
    return path


def test_clone_interactive_cria_busca_e_resumo_estaticos(tmp_path):
    entrada = _make_base_workbook(tmp_path / "entrada.xlsx")
    saida = clone_interactive(entrada)

    wb = openpyxl.load_workbook(saida)
    assert "🔍 Busca Ativa" in wb.sheetnames
    assert "📊 Resumo" in wb.sheetnames

    ws_busca = wb["🔍 Busca Ativa"]
    assert ws_busca["B5"].value == "Bruno"
    assert ws_busca["A5"].value == "🔴 URGENTE"
    assert ws_busca["B6"].value == "Ana"
    assert ws_busca["A6"].value == "🟠 ALTA"
    assert ws_busca["B7"].value == "Carlos"
    assert ws_busca["A7"].value == "🟢 CONCLUÍDO"

    ws_resumo = wb["📊 Resumo"]
    assert "RESUMO" in str(ws_resumo["A1"].value)
    assert ws_resumo["A10"].value == "Ótimo"


def test_refresh_interactive_workbook_recalcula_pontuacao(tmp_path):
    path = _make_base_workbook(tmp_path / "base.xlsx")
    clone = clone_interactive(path)

    wb = openpyxl.load_workbook(clone)
    ws = wb["📋 Dados C1"]
    ws["F5"] = "SIM"  # Bruno passa a ter um critério atendido
    wb.save(clone)
    wb.close()

    refresh_interactive_workbook(clone)

    wb2 = openpyxl.load_workbook(clone)
    ws2 = wb2["📋 Dados C1"]
    assert ws2[5][6].value == 50  # Pontuação na coluna G, índice 6 zero-based na tupla row? adjust below
    assert ws2[5][7].value == "Suficiente"
    busca = wb2["🔍 Busca Ativa"]
    nomes = [busca[f"B{r}"].value for r in (5, 6, 7)]
    assert "Bruno" in nomes

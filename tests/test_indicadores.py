"""
tests/test_indicadores.py

Testes unitários para os indicadores C1–C7.
Usa CSVs sintéticos em memória — não requer arquivos reais do e-SUS.

Execute com:
    pytest                    # todos os testes
    pytest -v                 # verbose
    pytest tests/ --tb=short  # traceback curto
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import openpyxl
import pandas as pd
import pytest


# ------------------------------------------------------------------
# Dados base e helpers
# ------------------------------------------------------------------

BASE = {
    "Nome": "Paciente Teste",
    "Data de nascimento": "01/01/1980",
    "Idade": "44 anos",
    "Sexo": "Feminino",
    "Raça/cor": "Branca",
    "Microárea": "001",
    "Rua": "Rua das Flores",
    "Número": "100",
    "Complemento": "",
    "Bairro": "Centro",
    "Telefone celular": "19999990000",
    "Telefone residencial": "",
    "Telefone de contato": "",
    "CPF": "000.000.000-00",
    "CNS": "123456789012345",
    "Meses desde o último atendimento médico": "3",
    "Meses desde o último atendimento de enfermagem": "3",
    "Meses desde o último atendimento odontológico": "10",
    "Meses desde a última visita domiciliar": "6",
    "Última medição de peso": "65",
    "Última medição de altura": "165",
    "Data da ultima medição de peso e altura": "01/01/2024",
    "Última medição de pressão arterial": "120/80",
    "Data da última medição de pressão arterial": "01/01/2024",
    "Últimas visitas domiciliares": "01/01/2024",
    "Quantidade de visitas domiciliares": "3",
}


def make_csv(rows) -> Path:
    if isinstance(rows, dict):
        rows = [rows]
    tmp = tempfile.NamedTemporaryFile(
        suffix=".csv", delete=False, mode="w", encoding="utf-8-sig", newline=""
    )
    pd.DataFrame(rows).to_csv(tmp, index=False, sep=";")
    tmp.close()
    return Path(tmp.name)


def read_result(saida: Path) -> pd.DataFrame:
    """Lê primeira aba do xlsx: linha 1-2 = títulos, linha 3 = cabeçalho, linha 4+ = dados."""
    wb = openpyxl.load_workbook(saida)
    ws = wb.worksheets[0]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    data = [dict(zip(headers, row)) for row in ws.iter_rows(min_row=4, values_only=True)]
    return pd.DataFrame(data)


def run(processar_fn, rows) -> pd.DataFrame:
    saida = Path(tempfile.mktemp(suffix=".xlsx"))
    processar_fn(make_csv(rows), saida)
    return read_result(saida)


# ------------------------------------------------------------------
# C1 — Mais Acesso
# ------------------------------------------------------------------

class TestC1:
    def setup_method(self):
        from c1_mais_acesso import processar
        self.p = processar

    def test_pontuacao_maxima(self):
        assert int(run(self.p, BASE)["Pontuação"].iloc[0]) == 100

    def test_pontuacao_zero(self):
        row = {**BASE,
               "Meses desde o último atendimento médico": "99",
               "Meses desde o último atendimento de enfermagem": "99",
               "Meses desde o último atendimento odontológico": "99",
               "Meses desde a última visita domiciliar": "99",
               "Microárea": ""}
        assert int(run(self.p, row)["Pontuação"].iloc[0]) == 0

    def test_linha_sem_nome_ignorada(self):
        assert len(run(self.p, {**BASE, "Nome": ""})) == 0

    def test_classificacao_otimo(self):
        assert run(self.p, BASE)["Classificação"].iloc[0] == "Ótimo"

    def test_multiplos_pacientes(self):
        rows = [{**BASE, "Nome": f"Paciente {i}"} for i in range(5)]
        assert len(run(self.p, rows)) == 5


# ------------------------------------------------------------------
# C4 — Diabetes
# ------------------------------------------------------------------

class TestC4:
    def setup_method(self):
        from c4_diabetes import processar
        self.p = processar

    def _completo(self):
        return {**BASE,
                "Meses desde o último atendimento médico": "3",
                "Última medição de pressão arterial": "120/80",
                "Última medição de peso": "70",
                "Última medição de altura": "170",
                "Quantidade de visitas domiciliares": "3",
                "Hemoglobina glicada": "5.8",
                "Data da avaliação dos pés": "01/01/2024"}

    def test_pontuacao_maxima(self):
        assert int(run(self.p, self._completo())["Pontuação"].iloc[0]) == 100

    def test_sem_hb_glicada(self):
        """Sem hemoglobina glicada perde 15 pts → 85."""
        assert int(run(self.p, {**self._completo(), "Hemoglobina glicada": ""})["Pontuação"].iloc[0]) == 85

    def test_sem_pes(self):
        """Sem avaliação dos pés perde 15 pts → 85."""
        assert int(run(self.p, {**self._completo(), "Data da avaliação dos pés": ""})["Pontuação"].iloc[0]) == 85


# ------------------------------------------------------------------
# C5 — Hipertensão
# ------------------------------------------------------------------

class TestC5:
    def setup_method(self):
        from c5_hipertensao import processar
        self.p = processar

    def _completo(self):
        return {**BASE,
                "Meses desde o último atendimento médico": "4",
                "Última medição de pressão arterial": "130/85",
                "Última medição de peso": "72",
                "Última medição de altura": "168",
                "Quantidade de visitas domiciliares": "2"}

    def test_pontuacao_maxima(self):
        assert int(run(self.p, self._completo())["Pontuação"].iloc[0]) == 100

    def test_sem_pa(self):
        """Sem PA aferida perde 25 pts → 75."""
        assert int(run(self.p, {**self._completo(), "Última medição de pressão arterial": ""})["Pontuação"].iloc[0]) == 75


# ------------------------------------------------------------------
# Resiliência do sistema
# ------------------------------------------------------------------

class TestSistema:
    def test_csv_encoding_latin1(self, tmp_path):
        from c1_mais_acesso import processar
        entrada = tmp_path / "c1.csv"
        pd.DataFrame([BASE]).to_csv(entrada, index=False, sep=";", encoding="latin1")
        saida = tmp_path / "saida.xlsx"
        processar(entrada, saida)
        assert saida.exists()

    def test_linhas_vazias_ignoradas(self, tmp_path):
        from c1_mais_acesso import processar
        entrada = tmp_path / "c1.csv"
        pd.DataFrame([BASE, {k: "" for k in BASE}]).to_csv(
            entrada, index=False, sep=";", encoding="utf-8-sig"
        )
        saida = tmp_path / "saida.xlsx"
        processar(entrada, saida)
        assert len(read_result(saida)) == 1

    def test_erro_isolado_entre_indicadores(self, tmp_path):
        """Erro em C1 não deve cancelar C5. Valida o try/except por indicador."""
        from sistema_aps import process_selected

        row = {**BASE,
               "Meses desde o último atendimento médico": "4",
               "Última medição de pressão arterial": "130/85",
               "Última medição de peso": "72",
               "Última medição de altura": "168",
               "Quantidade de visitas domiciliares": "2"}
        pd.DataFrame([row]).to_csv(
            tmp_path / "c5.csv", index=False, sep=";", encoding="utf-8-sig"
        )
        (tmp_path / "c1.csv").write_bytes(b"\xff\xfe" + b"\x00" * 100)  # corrompido

        res = process_selected(["C1", "C5"], in_dir=tmp_path, out_dir=tmp_path / "out")
        status = {r["code"]: r["status"] for r in res}

        assert status.get("C5") == "ok"
        assert status.get("C1") in ("erro", "não encontrado")

    def test_build_base_row_sem_dataframe_fantasma(self):
        import inspect
        from aps_utils import build_base_row
        assert "pd.DataFrame(columns=row.index)" not in inspect.getsource(build_base_row)

    def test_config_carrega(self):
        import aps_config
        assert aps_config.cores()["azul_escuro"] == "1F4E79"
        assert "resultado" in aps_config.ignorar_marcadores()

from __future__ import annotations

import shutil
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from uuid import uuid4

import pandas as pd

from controle_quadrimestral import build_control_dataframe


TMP_ROOT = Path(__file__).resolve().parent / "_tmp_runtime"
TMP_ROOT.mkdir(parents=True, exist_ok=True)


@contextmanager
def local_tmpdir():
    path = TMP_ROOT / uuid4().hex
    path.mkdir(parents=True, exist_ok=False)
    try:
        yield path
    finally:
        shutil.rmtree(path, ignore_errors=True)


def test_controle_quadrimestral_regras_basicas():
    with local_tmpdir() as td:
        entrada = td / "entrada.csv"
        rows = [
            {
                "Nome": "Paciente Coberto",
                "CNS": "111",
                "CPF": "111",
                "Telefone celular": "21999990001",
                "Dias desde o último atendimento médico": "10",
                "Dias desde o último atendimento de enfermagem": "40",
            },
            {
                "Nome": "Paciente Pendente",
                "CNS": "222",
                "CPF": "222",
                "Telefone celular": "21999990002",
                "Dias desde o último atendimento médico": "150",
                "Dias desde o último atendimento de enfermagem": "-",
            },
            {
                "Nome": "Paciente Sem Dado",
                "CNS": "333",
                "CPF": "333",
                "Telefone celular": "21999990003",
                "Dias desde o último atendimento médico": "-",
                "Dias desde o último atendimento de enfermagem": "-",
            },
        ]
        pd.DataFrame(rows).to_csv(entrada, index=False, sep=";", encoding="utf-8-sig")

        df = build_control_dataframe(entrada, ref_date=date(2026, 4, 13))
        by_name = {r["Nome"]: r for _, r in df.iterrows()}

        a = by_name["Paciente Coberto"]
        assert a["Situacao"] == "COBERTO NO QUADRIMESTRE ATUAL"
        assert a["Proximo quadrimestre obrigatorio"] == "2026-Q2"
        assert a["Data limite"] == "31/08/2026"
        assert a["Semaforo"] == "VERDE"

        b = by_name["Paciente Pendente"]
        assert b["Situacao"] == "PRECISA ATENDER NO QUADRIMESTRE ATUAL"
        assert b["Proximo quadrimestre obrigatorio"] == "2026-Q1"
        assert b["Data limite"] == "30/04/2026"
        assert b["Semaforo"] == "AMARELO"

        c = by_name["Paciente Sem Dado"]
        assert c["Situacao"] == "SEM DADO DE ATENDIMENTO"
        assert c["Proximo quadrimestre obrigatorio"] == "2026-Q1"
        assert c["Data limite"] == "30/04/2026"

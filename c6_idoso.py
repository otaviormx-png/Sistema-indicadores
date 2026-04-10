
from __future__ import annotations

import sys
from pathlib import Path
import pandas as pd

from aps_utils import (
    BASE_CLINICAL_COLUMNS,
    BASE_PERSON_COLUMNS,
    IndicatorConfig,
    build_base_row,
    classify_score,
    count_ge,
    has_any_text,
    months_leq,
    process_indicator,
    to_numeric,
    value,
)

FILTER_FUNC = lambda b,r: to_numeric(str(b.get("Idade","")).split(" ")[0], 0) >= 60

CRITERIA = [
    {"letter":"A", "label":"Consulta médico/enf (12m)", "weight":25, "func": lambda b,r: months_leq(b.get("Meses desde o último atendimento médico"), 12) or months_leq(b.get("Meses desde o último atendimento de enfermagem"), 12)},
    {"letter":"B", "label":"Antropometria (12m)", "weight":25, "func": lambda b,r: count_ge(value(r, "Registros de peso e altura simultâneos nos últimos 12 meses"), 1) or (has_any_text(b.get("Última medição de peso")) and has_any_text(b.get("Última medição de altura")))},
    {"letter":"C", "label":"Visitas ACS ≥2", "weight":25, "func": lambda b,r: count_ge(b.get("Quantidade de visitas domiciliares"), 2)},
    {"letter":"D", "label":"Vacina influenza (12m)", "weight":25, "func": lambda b,r: has_any_text(value(r, "Influenza (últimos 12 meses)"))},
]
EXTRA_COLUMNS = ['Registros de peso e altura simultâneos nos últimos 12 meses', 'Influenza (últimos 12 meses)', 'IVCF-20 Índice', 'IVCF-20 Pontuação']
CODE = 'C6'
TITULO = 'PLANILHA DE CUIDADO DA PESSOA IDOSA  |  Indicador C6 – Atenção Primária à Saúde'
CRITERIO_BLOCO = '◀ CRITÉRIOS C6 – NOTA METODOLÓGICA ▶'
SUBTITULO = 'Cada critério vale 25 pts  |  A=Consulta médico/enf (12m)  |  B=Antropometria (12m)  |  C=≥2 visitas ACS com ≥30d  |  D=Vacina influenza (12m)'
THEME_KEYWORDS = ['Pessoa idosa', 'idosa', 'idoso']
OFFICIAL_LIKE = True

def build_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_raw.iterrows():
        nome = value(row, "Nome", "Nome do cidadão", "Paciente")
        if not str(nome).strip() or str(nome).strip() == "-":
            continue

        base = build_base_row(row, EXTRA_COLUMNS)
        if FILTER_FUNC is not None and not FILTER_FUNC(base, row):
            continue

        criterio_vals = []
        pendencias = []
        score = 0

        for item in CRITERIA:
            ok = bool(item["func"](base, row))
            crit_name = f"{item['letter']} - {item['label']}"
            criterio_vals.append((crit_name, "SIM" if ok else "NÃO"))
            if ok:
                score += item["weight"]
            else:
                pendencias.append(item["label"])

        classificacao, prioridade = classify_score(score)
        out = {}
        out.update(base)
        for crit_name, crit_value in criterio_vals:
            out[crit_name] = crit_value
        out["Pontuação"] = score
        out["Classificação"] = classificacao
        out["Prioridade"] = prioridade
        out["Pendências"] = " | ".join(pendencias)
        rows.append(out)

    ordered = BASE_PERSON_COLUMNS + [c for c in BASE_CLINICAL_COLUMNS + EXTRA_COLUMNS if c not in BASE_PERSON_COLUMNS]
    ordered += [f"{c['letter']} - {c['label']}" for c in CRITERIA] + ["Pontuação", "Classificação", "Prioridade", "Pendências"]
    if not rows:
        return pd.DataFrame(columns=ordered)
    df = pd.DataFrame(rows)
    for col in ordered:
        if col not in df.columns:
            df[col] = ""
    return df[ordered]

CFG = IndicatorConfig(
    code=CODE,
    titulo=TITULO,
    criterio_bloco=CRITERIO_BLOCO,
    subtitulo=SUBTITULO,
    theme_keywords=THEME_KEYWORDS,
    criteria=[{"letter": c["letter"], "label": c["label"], "weight": c["weight"]} for c in CRITERIA],
    extra_columns=EXTRA_COLUMNS,
    builder=build_dataframe,
    official_like=OFFICIAL_LIKE,
)

def processar(entrada: str | Path, saida: str | Path):
    process_indicator(CFG, entrada, saida)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(f"Uso: python {Path(__file__).name} <entrada.csv/xlsx> <saida.xlsx>")
    processar(sys.argv[1], sys.argv[2])

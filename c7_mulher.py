
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

FILTER_FUNC = None

CRITERIA = [
    {"letter":"A", "label":"Colo uterino (25-64a / 36m)", "weight":20, "func": lambda b,r: (
        25 <= to_numeric(str(b.get("Idade", "")).split(" ")[0], 0) <= 64 and (
            has_any_text(value(r, "Exame de rastreamento de câncer de colo de útero última solicitação")) or
            has_any_text(value(r, "Exame de rastreamento de câncer de colo de útero data última solicitação")) or
            has_any_text(value(r, "Exame de rastreamento de câncer de colo de útero última avaliação")) or
            has_any_text(value(r, "Exame de rastreamento de câncer de colo de útero data última avaliação"))
        )
    ) or not (25 <= to_numeric(str(b.get("Idade", "")).split(" ")[0], 0) <= 64)},
    {"letter":"B", "label":"HPV (9-14a)", "weight":30, "func": lambda b,r: (
        9 <= to_numeric(str(b.get("Idade", "")).split(" ")[0], 0) <= 14 and has_any_text(value(r, "HPV"))
    ) or not (9 <= to_numeric(str(b.get("Idade", "")).split(" ")[0], 0) <= 14)},
    {"letter":"C", "label":"Saúde sexual/reprodutiva (12m)", "weight":30, "func": lambda b,r: has_any_text(value(r, "Data da última consulta de saúde sexual e reprodutiva"))},
    {"letter":"D", "label":"Mamografia (50-69a / 24m)", "weight":20, "func": lambda b,r: (
        50 <= to_numeric(str(b.get("Idade", "")).split(" ")[0], 0) <= 69 and (
            has_any_text(value(r, "Exame de rastreamento de câncer de mama data Última solicitação")) or
            has_any_text(value(r, "Exame de rastreamento de câncer de mama data Última realização")) or
            has_any_text(value(r, "Exame de rastreamento de câncer de mama data Última avaliação"))
        )
    ) or not (50 <= to_numeric(str(b.get("Idade", "")).split(" ")[0], 0) <= 69)},
]
EXTRA_COLUMNS = ['Data da última consulta de saúde sexual e reprodutiva', 'Rastreamento e acompanhamento de HIV data ultima avaliação', 'Rastreamento e acompanhamento de Sífilis data ultima avaliação', 'Rastreamento e acompanhamento de Hepatite B data ultima avaliação', 'Rastreamento e acompanhamento de Hepatite C data ultima avaliação', 'Exame de rastreamento de câncer de colo de útero última solicitação', 'Exame de rastreamento de câncer de colo de útero data última solicitação', 'Exame de rastreamento de câncer de colo de útero última avaliação', 'Exame de rastreamento de câncer de colo de útero data última avaliação', 'Exame de rastreamento de câncer de mama data Última solicitação', 'Exame de rastreamento de câncer de mama data Última realização', 'Exame de rastreamento de câncer de mama data Última avaliação', 'HPV']
CODE = 'C7'
TITULO = 'PLANILHA DE CUIDADO DA MULHER NA PREVENÇÃO DO CÂNCER  |  Indicador C7 – Atenção Primária à Saúde'
CRITERIO_BLOCO = '◀ CRITÉRIOS C7 – NOTA METODOLÓGICA ▶'
SUBTITULO = 'A=colo do útero (20)  |  B=HPV (30)  |  C=saúde sexual/reprodutiva (30)  |  D=mamografia (20)'
THEME_KEYWORDS = ['Saúde da mulher', 'mulher', 'cancer', 'saude da mulher']
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

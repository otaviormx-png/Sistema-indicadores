from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

from aps_utils import (
    BASE_CLINICAL_COLUMNS,
    BASE_PERSON_COLUMNS,
    IndicatorConfig,
    classify_score,
    normalize_text,
    process_indicator,
    to_numeric,
)

FILTER_FUNC = None

CRITERIA = [
    {
        "letter": "A",
        "label": "Formula oficial C1 calculada",
        "weight": 100,
        "func": None,
    },
]

EXTRA_COLUMNS = [
    "Numerador C1 (demanda programada)",
    "Denominador C1 (todas as demandas)",
    "Percentual C1 oficial (%)",
    "Fonte do calculo C1",
]

CODE = "C1O"
TITULO = "PLANILHA C1 OFICIAL  |  Mais acesso (metodologia oficial)"
CRITERIO_BLOCO = "CRITERIO C1 OFICIAL"
SUBTITULO = "Numerador: atendimentos por demanda programada | Denominador: total de demandas"
THEME_KEYWORDS = ["mais acesso", "geral", "demanda programada", "c1"]
OFFICIAL_LIKE = True


def _sum_columns(df: pd.DataFrame, cols: list[str]) -> float:
    total = 0.0
    for col in cols:
        total += df[col].map(lambda v: to_numeric(v, 0)).sum()
    return float(total)


def _cols_by_tokens(df: pd.DataFrame, required: list[str], forbidden: list[str] | None = None) -> list[str]:
    forbidden = forbidden or []
    out: list[str] = []
    for col in df.columns:
        norm = normalize_text(col)
        if all(tok in norm for tok in required) and all(tok not in norm for tok in forbidden):
            out.append(col)
    return out


def _from_tipo_demanda(df: pd.DataFrame) -> tuple[float, float, str]:
    tipo_cols = _cols_by_tokens(df, ["tipo", "demanda"])
    if not tipo_cols:
        return 0.0, 0.0, ""
    col = tipo_cols[0]
    values = df[col].fillna("").map(normalize_text)

    def is_programada(txt: str) -> bool:
        keys = [
            "consulta agendada programada",
            "cuidado continuado",
            "consulta agendada",
        ]
        return any(k in txt for k in keys) or ("programad" in txt)

    def is_denominador(txt: str) -> bool:
        espontanea_keys = [
            "escuta inicial",
            "consulta no dia",
            "urgencia",
            "espont",
        ]
        return is_programada(txt) or any(k in txt for k in espontanea_keys)

    numerador = float(values.map(lambda x: 1 if is_programada(x) else 0).sum())
    denominador = float(values.map(lambda x: 1 if is_denominador(x) else 0).sum())
    return numerador, denominador, f"tipo demanda: {col}"


def _compute_c1_official(df_raw: pd.DataFrame) -> tuple[float, float, str]:
    numerador_cols = _cols_by_tokens(
        df_raw,
        required=["demanda", "programad"],
        forbidden=["espont"],
    )
    denominador_cols = _cols_by_tokens(
        df_raw,
        required=["demanda"],
    )
    denominador_cols = [
        c for c in denominador_cols if any(x in normalize_text(c) for x in ["total", "todos", "tipos"])
    ]
    espont_cols = _cols_by_tokens(
        df_raw,
        required=["demanda", "espont"],
    )

    numerador = _sum_columns(df_raw, numerador_cols) if numerador_cols else 0.0
    denominador = _sum_columns(df_raw, denominador_cols) if denominador_cols else 0.0

    if denominador <= 0 and numerador_cols and espont_cols:
        denominador = numerador + _sum_columns(df_raw, espont_cols)

    source_parts = []
    if numerador_cols:
        source_parts.append("num=" + ", ".join(numerador_cols))
    if denominador_cols:
        source_parts.append("den=" + ", ".join(denominador_cols))
    elif espont_cols and numerador_cols:
        source_parts.append("den=num+espont (" + ", ".join(espont_cols) + ")")

    if numerador <= 0 and denominador <= 0:
        n2, d2, src2 = _from_tipo_demanda(df_raw)
        if d2 > 0:
            numerador, denominador = n2, d2
            source_parts = [src2]

    return numerador, denominador, " | ".join(source_parts)


def build_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    numerador, denominador, fonte = _compute_c1_official(df_raw)

    if denominador > 0:
        percentual = round((numerador / denominador) * 100, 2)
        percentual = min(percentual, 100.0)
        criterio = "SIM"
        pendencias = ""
    else:
        percentual = 0.0
        criterio = "NÃO"
        pendencias = "Não foi possível calcular denominador do C1 oficial"

    classificacao, prioridade = classify_score(percentual)

    row = {c: "" for c in (BASE_PERSON_COLUMNS + BASE_CLINICAL_COLUMNS + EXTRA_COLUMNS)}
    row["Nome"] = "AGREGADO C1 OFICIAL"
    row["Numerador C1 (demanda programada)"] = numerador
    row["Denominador C1 (todas as demandas)"] = denominador
    row["Percentual C1 oficial (%)"] = percentual
    row["Fonte do calculo C1"] = fonte
    row["A - Formula oficial C1 calculada"] = criterio
    row["Pontuação"] = percentual
    row["Classificação"] = classificacao
    row["Prioridade"] = prioridade
    row["Pendências"] = pendencias

    ordered = BASE_PERSON_COLUMNS + [c for c in BASE_CLINICAL_COLUMNS + EXTRA_COLUMNS if c not in BASE_PERSON_COLUMNS]
    ordered += [f"{c['letter']} - {c['label']}" for c in CRITERIA] + ["Pontuação", "Classificação", "Prioridade", "Pendências"]

    df = pd.DataFrame([row])
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
    official_like=True,
)


def processar(entrada: str | Path, saida: str | Path):
    process_indicator(CFG, entrada, saida)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(f"Uso: python {Path(__file__).name} <entrada.csv/xlsx> <saida.xlsx>")
    processar(sys.argv[1], sys.argv[2])


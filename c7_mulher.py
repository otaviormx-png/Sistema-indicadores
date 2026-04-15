from __future__ import annotations

import sys
from pathlib import Path
import pandas as pd

from aps_utils import (
    BASE_CLINICAL_COLUMNS,
    BASE_PERSON_COLUMNS,
    IndicatorConfig,
    age_years,
    build_base_row,
    classify_score,
    has_any_text,
    has_recent_date_or_text,
    process_indicator,
    value,
    value_norm,
)


def _idade(base: dict) -> int:
    return age_years(base.get("Idade"), -1)


def _faixa(base: dict, minimo: int, maximo: int) -> bool:
    idade = _idade(base)
    return minimo <= idade <= maximo


FILTER_FUNC = lambda b, r: _faixa(b, 9, 69)

CRITERIA = [
    {
        "letter": "A",
        "label": "Colo uterino (25-64a / 36m)",
        "weight": 20,
        "applies": lambda b, r: _faixa(b, 25, 64),
        "func": lambda b, r: any(
            has_recent_date_or_text(
                value_norm(r, field),
                36,
            )
            for field in [
                "Exame de rastreamento de câncer de colo de útero última solicitação",
                "Exame de rastreamento de câncer de colo de útero data última solicitação",
                "Exame de rastreamento de câncer de colo de útero última avaliação",
                "Exame de rastreamento de câncer de colo de útero data última avaliação",
            ]
        ),
    },
    {
        "letter": "B",
        "label": "HPV (9-14a)",
        "weight": 30,
        "applies": lambda b, r: _faixa(b, 9, 14),
        "func": lambda b, r: has_any_text(value_norm(r, "HPV")),
    },
    {
        "letter": "C",
        "label": "Saúde sexual/reprodutiva (12m)",
        "weight": 30,
        "applies": lambda b, r: _faixa(b, 14, 69),
        "func": lambda b, r: has_recent_date_or_text(
            value_norm(r, "Data da última consulta de saúde sexual e reprodutiva"),
            12,
        ),
    },
    {
        "letter": "D",
        "label": "Mamografia (50-69a / 24m)",
        "weight": 20,
        "applies": lambda b, r: _faixa(b, 50, 69),
        "func": lambda b, r: any(
            has_recent_date_or_text(
                value_norm(r, field),
                24,
            )
            for field in [
                "Exame de rastreamento de câncer de mama data última solicitação",
                "Exame de rastreamento de câncer de mama data última realização",
                "Exame de rastreamento de câncer de mama data última avaliação",
            ]
        ),
    },
]

EXTRA_COLUMNS = [
    "Data da última consulta de saúde sexual e reprodutiva",
    "Rastreamento e acompanhamento de HIV data ultima avaliação",
    "Rastreamento e acompanhamento de Sífilis data ultima avaliação",
    "Rastreamento e acompanhamento de Hepatite B data ultima avaliação",
    "Rastreamento e acompanhamento de Hepatite C data ultima avaliação",
    "Exame de rastreamento de câncer de colo de útero última solicitação",
    "Exame de rastreamento de câncer de colo de útero data última solicitação",
    "Exame de rastreamento de câncer de colo de útero última avaliação",
    "Exame de rastreamento de câncer de colo de útero data última avaliação",
    "Exame de rastreamento de câncer de mama data última solicitação",
    "Exame de rastreamento de câncer de mama data última realização",
    "Exame de rastreamento de câncer de mama data última avaliação",
    "HPV",
]
CODE = "C7"
TITULO = "PLANILHA DE CUIDADO DA MULHER NA PREVENÇÃO DO CÂNCER  |  Indicador C7 – Atenção Primária à Saúde"
CRITERIO_BLOCO = "◀ CRITÉRIOS C7 – NOTA METODOLÓGICA ▶"
SUBTITULO = "A=colo do útero (20)  |  B=HPV (30)  |  C=saúde sexual/reprodutiva (30)  |  D=mamografia (20)"
THEME_KEYWORDS = ["Saúde da mulher", "mulher", "cancer", "saude da mulher"]
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
        score_raw = 0
        score_max = 0

        for item in CRITERIA:
            crit_name = f"{item['letter']} - {item['label']}"
            aplica = bool(item["applies"](base, row))
            if not aplica:
                criterio_vals.append((crit_name, "N/A"))
                continue

            score_max += item["weight"]
            ok = bool(item["func"](base, row))
            criterio_vals.append((crit_name, "SIM" if ok else "NÃO"))
            if ok:
                score_raw += item["weight"]
            else:
                pendencias.append(item["label"])

        score = round((score_raw / score_max) * 100) if score_max else 0
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

"""
plugins/c8_exemplo.py — Template de indicador personalizado.

Renomeie este arquivo para c8_meu_indicador.py (ou c9_, c10_...)
e ajuste os critérios para o seu caso de uso.

ATENÇÃO: este arquivo é apenas um exemplo comentado.
         Renomeie-o antes de usar — arquivos começando com
         'exemplo' não são carregados automaticamente.
"""
from __future__ import annotations

from pathlib import Path
import pandas as pd

from aps_utils import (
    BASE_CLINICAL_COLUMNS,
    BASE_PERSON_COLUMNS,
    IndicatorConfig,
    build_base_row,
    classify_score,
    has_any_text,
    months_leq,
    process_indicator,
    value,
)

# ------------------------------------------------------------------
# 1. Defina os critérios do seu indicador
# ------------------------------------------------------------------
CRITERIA = [
    {
        "letter": "A",
        "label": "Consulta nos últimos 6 meses",
        "weight": 50,
        "func": lambda b, r: months_leq(
            b.get("Meses desde o último atendimento médico"), 6),
    },
    {
        "letter": "B",
        "label": "Visita domiciliar registrada",
        "weight": 50,
        "func": lambda b, r: has_any_text(
            b.get("Meses desde a última visita domiciliar")),
    },
]

EXTRA_COLUMNS: list[str] = []  # colunas extras além das padrão

# ------------------------------------------------------------------
# 2. Metadados do indicador
# ------------------------------------------------------------------
CODE = "C8"
TITULO = "PLANILHA PERSONALIZADA  |  Indicador C8 – Exemplo"
CRITERIO_BLOCO = "◀ CRITÉRIOS C8 ▶"
SUBTITULO = "A=Consulta (50)  |  B=Visita domiciliar (50)"
THEME_KEYWORDS = ["meu_indicador"]
OFFICIAL_LIKE = False
FILTER_FUNC = None  # ou: lambda base, row: <condição de filtro>


# ------------------------------------------------------------------
# 3. Builder — não precisa alterar se os critérios acima estiverem ok
# ------------------------------------------------------------------
def build_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df_raw.iterrows():
        nome = value(row, "Nome", "Nome do cidadão", "Paciente")
        if not str(nome).strip() or str(nome).strip() == "-":
            continue
        base = build_base_row(row, EXTRA_COLUMNS)
        if FILTER_FUNC is not None and not FILTER_FUNC(base, row):
            continue
        criterio_vals, pendencias, score = [], [], 0
        for item in CRITERIA:
            ok = bool(item["func"](base, row))
            crit_name = f"{item['letter']} - {item['label']}"
            criterio_vals.append((crit_name, "SIM" if ok else "NÃO"))
            if ok:
                score += item["weight"]
            else:
                pendencias.append(item["label"])
        classificacao, prioridade = classify_score(score)
        out = {**base}
        for cn, cv in criterio_vals:
            out[cn] = cv
        out["Pontuação"] = score
        out["Classificação"] = classificacao
        out["Prioridade"] = prioridade
        out["Pendências"] = " | ".join(pendencias)
        rows.append(out)

    ordered = (BASE_PERSON_COLUMNS
               + [c for c in BASE_CLINICAL_COLUMNS + EXTRA_COLUMNS
                  if c not in BASE_PERSON_COLUMNS]
               + [f"{c['letter']} - {c['label']}" for c in CRITERIA]
               + ["Pontuação", "Classificação", "Prioridade", "Pendências"])
    if not rows:
        return pd.DataFrame(columns=ordered)
    df = pd.DataFrame(rows)
    for col in ordered:
        if col not in df.columns:
            df[col] = ""
    return df[ordered]


# ------------------------------------------------------------------
# 4. Objeto de configuração e função processar — obrigatórios
# ------------------------------------------------------------------
CFG = IndicatorConfig(
    code=CODE,
    titulo=TITULO,
    criterio_bloco=CRITERIO_BLOCO,
    subtitulo=SUBTITULO,
    theme_keywords=THEME_KEYWORDS,
    criteria=[{"letter": c["letter"], "label": c["label"],
               "weight": c["weight"]} for c in CRITERIA],
    extra_columns=EXTRA_COLUMNS,
    builder=build_dataframe,
    official_like=OFFICIAL_LIKE,
)


def processar(entrada: str | Path, saida: str | Path) -> None:
    process_indicator(CFG, entrada, saida)

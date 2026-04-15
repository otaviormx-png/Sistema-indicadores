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
    is_team_type_76,
    months_leq,
    process_indicator,
    value,
    value_norm,
)

FILTER_FUNC = None

CRITERIA = [
    {
        "letter": "A",
        "label": "Consulta medico/enf (6m)",
        "weight": 20,
        "func": lambda b, r: months_leq(b.get("Meses desde o último atendimento médico"), 6)
        or months_leq(b.get("Meses desde o último atendimento de enfermagem"), 6)
        or has_any_text(value_norm(r, "Data da última consulta")),
    },
    {
        "letter": "B",
        "label": "PA aferida (6m)",
        "weight": 15,
        "func": lambda b, r: has_any_text(b.get("Última medição de pressão arterial")),
    },
    {
        "letter": "C",
        "label": "Peso+altura (12m)",
        "weight": 15,
        "func": lambda b, r: has_any_text(b.get("Última medição de peso"))
        and has_any_text(b.get("Última medição de altura")),
    },
    {
        "letter": "D",
        "label": "Visitas ACS >=2",
        "weight": 20,
        "func": lambda b, r: is_team_type_76(r) or count_ge(b.get("Quantidade de visitas domiciliares"), 2),
    },
    {
        "letter": "E",
        "label": "Hb glicada (12m)",
        "weight": 15,
        "func": lambda b, r: has_any_text(
            value_norm(
                r,
                "Hemoglobina glicada",
                "Data da última avaliação de hemoglobina glicada",
                "Data da última solicitação de hemoglobina glicada",
            )
        ),
    },
    {
        "letter": "F",
        "label": "Avaliacao dos pes (12m)",
        "weight": 15,
        "func": lambda b, r: has_any_text(value_norm(r, "Data da avaliação dos pés")),
    },
]
EXTRA_COLUMNS = [
    "Hemoglobina glicada",
    "Data da última avaliação de hemoglobina glicada",
    "Data da última solicitação de hemoglobina glicada",
    "Data da avaliação dos pés",
    "Data da última consulta",
    "Consultas (últimos 36 meses)",
    "Incluído na lista de problemas e condições",
]
CODE = "C4"
TITULO = "PLANILHA DE CUIDADO DA PESSOA COM DIABETES  |  Indicador C4 – Atenção Primária à Saúde"
CRITERIO_BLOCO = "◀ CRITÉRIOS C4 – NOTA METODOLÓGICA ▶"
SUBTITULO = "A=Consulta (20)  |  B=PA (15)  |  C=Peso+altura (15)  |  D=Visitas ACS (20)  |  E=Hemoglobina glicada (15)  |  F=Pes (15)"
THEME_KEYWORDS = ["Diabetes", "diabetes"]
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


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
    to_numeric,
    value,
)

FILTER_FUNC = None

CRITERIA = [
    {"letter":"A", "label":"1ª consulta até 30 dias", "weight":20, "func": lambda b,r: has_any_text(value(r, "Data da primeira consulta")) and to_numeric(value(r, "Idade na primeira consulta"), 999) <= 30},
    {"letter":"B", "label":"≥9 consultas até 24m", "weight":20, "func": lambda b,r: count_ge(value(r, "Quantidade de consultas até 24 meses"), 9)},
    {"letter":"C", "label":"≥9 peso/altura até 24m", "weight":20, "func": lambda b,r: count_ge(value(r, "Quantidade de medições de peso/altura simultâneas até 24 meses"), 9)},
    {"letter":"D", "label":"2 visitas nos marcos", "weight":20, "func": lambda b,r: is_team_type_76(r) or (has_any_text(value(r, "Data da primeira visita domiciliar")) and has_any_text(value(r, "Data da segunda visita domiciliar")))},
    {"letter":"E", "label":"Vacinas completas", "weight":20, "func": lambda b,r: (
        has_any_text(value(r, "Difteria, Tétano, Pertusis, Hepatite B, Haemophilus Influenza B")) and
        has_any_text(value(r, "Poliomielite")) and
        has_any_text(value(r, "Sarampo, Caxumba, Rubéola")) and
        has_any_text(value(r, "Pneumocócica"))
    )},
]
EXTRA_COLUMNS = ['Data da primeira consulta', 'Idade na primeira consulta', 'Quantidade de consultas até 24 meses', 'Quantidade de medições de peso/altura simultâneas até 24 meses', 'Data da primeira visita domiciliar', 'Data da segunda visita domiciliar', 'Quantidade de visitas domiciliares até os 24 meses de idade', 'Difteria, Tétano, Pertusis, Hepatite B, Haemophilus Influenza B', 'Poliomielite', 'Sarampo, Caxumba, Rubéola', 'Pneumocócica', 'Última suplementação', 'Idade na última suplementação', 'Doses de 6 a 11 meses', 'Doses de 12 a 23 meses', 'Doses de 24 a 59 meses']
CODE = 'C2'
TITULO = 'PLANILHA DE CUIDADO NO DESENVOLVIMENTO INFANTIL  |  Indicador C2 – Atenção Primária à Saúde'
CRITERIO_BLOCO = '◀ CRITÉRIOS C2 – NOTA METODOLÓGICA ▶'
SUBTITULO = 'Cada critério vale 20 pts  |  A=1ª consulta até 30 dias  |  B=≥9 consultas  |  C=≥9 peso/altura  |  D=2 visitas nos marcos  |  E=Vacinas completas'
THEME_KEYWORDS = ['Desenvolvimento infantil', 'infantil']
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

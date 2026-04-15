
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
    {"letter":"A", "label":"1ª consulta até 12 sem", "weight":10, "func": lambda b,r: count_ge(value(r, "Quantidade de atendimentos até 12 semanas no pré-natal"), 1)},
    {"letter":"B", "label":"≥7 consultas pré-natal", "weight":9, "func": lambda b,r: count_ge(value(r, "Quantidade de atendimentos no pré-natal"), 7)},
    {"letter":"C", "label":"≥7 PA na gestação", "weight":9, "func": lambda b,r: count_ge(value(r, "Quantidade de medições de pressão arterial"), 7)},
    {"letter":"D", "label":"≥7 peso/altura", "weight":9, "func": lambda b,r: count_ge(value(r, "Quantidade de medições simultâneas de peso e altura"), 7)},
    {"letter":"E", "label":"≥3 visitas ACS pré-natal", "weight":9, "func": lambda b,r: is_team_type_76(r) or count_ge(value(r, "Quantidade de visitas domiciliares no pré-natal"), 3)},
    {"letter":"F", "label":"dTpa registrada", "weight":9, "func": lambda b,r: has_any_text(value(r, "dTpa"))},
    {"letter":"G", "label":"Exames 1º trimestre", "weight":9, "func": lambda b,r: (
        has_any_text(value(r, "Exame de HIV no primeiro trimestre")) and
        has_any_text(value(r, "Exame de Sífilis no primeiro trimestre)")) and
        has_any_text(value(r, "Exame de Hepatite B no primeiro trimestre")) and
        has_any_text(value(r, "Exame de Hepatite C no primeiro trimestre"))
    )},
    {"letter":"H", "label":"Exames 3º trimestre", "weight":9, "func": lambda b,r: (
        has_any_text(value(r, "Exame de HIV no terceiro trimestre")) and
        has_any_text(value(r, "Exame de Sifilis no terceiro trimestre"))
    )},
    {"letter":"I", "label":"Consulta no puerpério", "weight":9, "func": lambda b,r: count_ge(value(r, "Quantidade de atendimentos no puerpério"), 1) or has_any_text(value(r, "Última consulta de puerpério"))},
    {"letter":"J", "label":"Visita no puerpério", "weight":9, "func": lambda b,r: is_team_type_76(r) or count_ge(value(r, "Quantidade de visitas domiciliares no puerpério"), 1)},
    {"letter":"K", "label":"Saúde bucal gestação", "weight":9, "func": lambda b,r: count_ge(value(r, "Quantidade de atendimentos odontológicos no pré-natal"), 1)},
]
EXTRA_COLUMNS = ['Risco gestacional', 'DUM', 'IG (DUM) (semanas)', 'IG (DUM) (dias)', 'DPP (DUM)', 'IG (ecografia obstétrica) (semanas)', 'IG (ecografia obstétrica) (dias)', 'DPP (ecografia obstétrica)', 'Quantidade de atendimentos no pré-natal', 'Quantidade de atendimentos até 12 semanas no pré-natal', 'Última consulta de pré-natal', 'Quantidade de atendimentos odontológicos no pré-natal', 'dTpa', 'Quantidade de medições de altura uterina', 'Quantidade de medições de pressão arterial', 'Quantidade de medições simultâneas de peso e altura', 'Exame de HIV no primeiro trimestre', 'Exame de Sífilis no primeiro trimestre)', 'Exame de Hepatite B no primeiro trimestre', 'Exame de Hepatite C no primeiro trimestre', 'Exame de HIV no terceiro trimestre', 'Exame de Sifilis no terceiro trimestre', 'Quantidade de visitas domiciliares no pré-natal', 'Quantidade de visitas domiciliares no puerpério', 'Quantidade de atendimentos no puerpério', 'Última consulta de puerpério']
CODE = 'C3'
TITULO = 'PLANILHA DE CUIDADO NA GESTAÇÃO E PUERPÉRIO  |  Indicador C3 – Atenção Primária à Saúde'
CRITERIO_BLOCO = '◀ CRITÉRIOS C3 – NOTA METODOLÓGICA ▶'
SUBTITULO = 'A=1ª consulta até 12 sem (10)  |  B..K=boas práticas gestação/puerpério conforme nota'
THEME_KEYWORDS = ['Gestação e puerpério', 'gestacao', 'puerperio']
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

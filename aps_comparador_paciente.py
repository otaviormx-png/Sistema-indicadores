"""
aps_comparador_paciente.py
Cruza planilhas C1-C7 por paciente.
Exporta planilha visualmente rica com 3 abas.
"""
from __future__ import annotations

import os, re, unicodedata
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Paleta ───────────────────────────────────────────────────────────────
C = {
    "azul":    "1F4E79", "azul_m":  "2E75B6", "azul_h":  "BDD7EE", "azul_c": "EBF3FB",
    "verde":   "C6EFCE", "verde_e": "375623", "verde_c": "E8F5E9",
    "amarelo": "FFEB9C", "amar_t":  "9C5700", "amar_m":  "FFE699",
    "verm":    "FFC7CE", "verm_t":  "9C0006", "verm_e":  "C00000", "verm_c": "FFEBEE",
    "laranja": "FFF2CC", "lar_t":   "843C0C",
    "roxo":    "7030A0", "roxo_c":  "F3E5F5",
    "cinza":   "F5F5F5", "cinza_m": "E0E0E0", "branco":  "FFFFFF",
}
# Mapeia AMBOS os formatos: com emoji (planilha interativa) e sem (sistema original)
def _norm_prio(p: str) -> str:
    """
    Normaliza prioridade para formato padronizado com emoji.
    Cobre dois formatos:
      Sistema APS original: "Alta", "Média", "Baixa", "Concluído"
      Clonador interativo:  "🔴 URGENTE", "🟠 ALTA", "🟡 MONITORAR", "🟢 CONCLUÍDO"
    """
    p = str(p).strip()
    # 1. Verifica formato com emoji direto (do clonador) — PRIMEIRO para não confundir
    if "URGENTE" in p: return "🔴 URGENTE"
    if "MONITORAR" in p or "MONITOR" in p: return "🟡 MONITORAR"
    if "CONCLUÍDO" in p or "CONCLUIDO" in p: return "🟢 CONCLUÍDO"
    if "🟠" in p: return "🟠 ALTA"
    if "🟡" in p: return "🟡 MONITORAR"
    if "🟢" in p: return "🟢 CONCLUÍDO"
    if "🔴" in p: return "🔴 URGENTE"
    # 2. Formato texto sem emoji (sistema APS original)
    pu = p.upper().strip()
    if pu in ("ALTA",):                    return "🔴 URGENTE"   # Alta = pior no sistema original
    if pu in ("MÉDIA", "MEDIA"):           return "🟠 ALTA"
    if pu in ("BAIXA",):                   return "🟡 MONITORAR"
    if pu in ("CONCLUÍDO","CONCLUIDO","ÓTIMO","OTIMO"): return "🟢 CONCLUÍDO"
    # Fallback por substring
    if "URGENT" in pu or "CRITI" in pu:   return "🔴 URGENTE"
    if "BAIXA" in pu or "BOM" in pu:      return "🟡 MONITORAR"
    if "CONCLU" in pu or "ÓTIMO" in pu:   return "🟢 CONCLUÍDO"
    return p

PRIO_ORDER = {"🔴 URGENTE":0,"🟠 ALTA":1,"🟡 MONITORAR":2,"🟢 CONCLUÍDO":3}
PRIO_THEME = {
    "🔴 URGENTE":   (C["verm_c"],   C["verm_e"]),
    "🟠 ALTA":      (C["laranja"],  C["lar_t"]),
    "🟡 MONITORAR": (C["amarelo"],  C["amar_t"]),
    "🟢 CONCLUÍDO": (C["verde_c"],  C["verde_e"]),
}

def _f(h):  return PatternFill("solid", fgColor=h)
def _fn(bold=False,color="000000",size=9,italic=False):
    return Font(name="Calibri",bold=bold,color=color,size=size,italic=italic)
def _al(h="left",v="center",wrap=False):
    return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
def _bd(style="thin",color="CCCCCC"):
    s=Side(style=style,color=color); return Border(left=s,right=s,top=s,bottom=s)

def _norm(t):
    if t is None: return ""
    t = unicodedata.normalize("NFKD",str(t).strip()).encode("ascii","ignore").decode("ascii")
    return re.sub(r"\s+"," ",t).lower().strip()

def _detect_code(p:Path)->str:
    m=re.search(r"\b(C\d)\b",p.stem,re.I)
    return m.group(1).upper() if m else p.stem[:4]

def _read(path:Path):
    try:
        xls=pd.ExcelFile(path)
        sh=next((s for s in xls.sheet_names if str(s).startswith("📋 Dados")),xls.sheet_names[0])
        for h in (2,1,0):
            try:
                df=pd.read_excel(path,sheet_name=sh,header=h,dtype=str).dropna(how="all")
                if "Nome" in df.columns and len(df):
                    return df[df["Nome"].str.strip().ne("")].reset_index(drop=True)
            except: continue
    except: pass
    return None

# ── Motor ─────────────────────────────────────────────────────────────────
def build_unified(paths:list[Path])->pd.DataFrame:
    data={}
    for p in paths:
        df=_read(p)
        if df is None or "Nome" not in df.columns: continue
        df["_nn"]=df["Nome"].apply(_norm)
        data[_detect_code(p)]=df
    if not data: return pd.DataFrame()

    patients={}
    for code,df in data.items():
        for _,row in df.iterrows():
            nn=row["_nn"]
            if not nn: continue
            if nn not in patients:
                tel=(row.get("Telefone celular","") or row.get("Telefone residencial","")
                     or row.get("Telefone de contato","") or "")
                patients[nn]={"Nome":row.get("Nome",""),
                              "Microárea":row.get("Microárea","") or row.get("Microarea",""),
                              "Telefone":str(tel).strip() if str(tel).strip() not in {"","nan"} else ""}

    codes=sorted(data.keys())
    records=[]
    for nn,ident in patients.items():
        rec={"Nome":ident["Nome"],"Microárea":ident["Microárea"],"Telefone":ident["Telefone"]}
        present=[]; pend_parts=[]; prios=[]; sum_pts=0; cnt=0
        for code in codes:
            match=data[code][data[code]["_nn"]==nn]
            if match.empty: rec[code]="—"; continue
            r=match.iloc[0]
            pts=str(r.get("Pontuação","")).strip()
            pend=str(r.get("Pendências","")).strip()
            rec[code]=f"{pts} pts"
            present.append(code)
            try: sum_pts+=float(pts); cnt+=1
            except: pass
            if pend and pend.lower() not in {"","nan","none","-"}:
                pend_parts.append(f"[{code}] {pend}")
        media=int(round(sum_pts/cnt,0)) if cnt else 0
        # Prioridade calculada SEMPRE pela pontuação média — consistente independente do formato
        if   media >= 100: best = "🟢 CONCLUÍDO"
        elif media >= 75:  best = "🟡 MONITORAR"
        elif media >= 50:  best = "🟠 ALTA"
        else:              best = "🔴 URGENTE"
        rec.update({
            "Indicadores": " · ".join(present) if present else "—",
            "Qtd":len(present), "Pendências":len(pend_parts),
            "Média":media, "Prioridade":best,
            "O que fazer":"\n".join(pend_parts) if pend_parts else "✔ Em dia",
        })
        records.append(rec)

    df_out=pd.DataFrame(records)
    df_out["_o"]=df_out["Prioridade"].map(lambda p:PRIO_ORDER.get(p,4))
    # Ordena por: prioridade (urgente primeiro) → pontuação média (menor primeiro) → nome
    return df_out.sort_values(["_o","Média","Nome"],ascending=[True,True,True]).drop(columns=["_o"]).reset_index(drop=True)

# ── Excel premium ──────────────────────────────────────────────────────────
def _title(ws,row,text,bg,fg,size,ncols,h=28):
    ecl=get_column_letter(ncols)
    try: ws.merge_cells(f"A{row}:{ecl}{row}")
    except: pass
    c=ws.cell(row,1,text)
    c.font=_fn(bold=True,color=fg,size=size); c.fill=_f(bg)
    c.alignment=_al("center"); c.border=_bd("medium","595959")
    ws.row_dimensions[row].height=h

def _mw(ws,row,col,value,ec=None,er=None):
    """Escreve com merge opcional (como no clonador)."""
    if ec or er:
        r2=er or row; c2=ec or col
        try: ws.merge_cells(f"{get_column_letter(col)}{row}:{get_column_letter(c2)}{r2}")
        except: pass
    cell=ws.cell(row,col)
    cell.value=value
    return cell

def _section(ws,row,text,bg,fg,ncols):
    ecl=get_column_letter(ncols)
    try: ws.merge_cells(f"A{row}:{ecl}{row}")
    except: pass
    c=ws.cell(row,1,text)
    c.font=_fn(bold=True,color=fg,size=9); c.fill=_f(bg)
    c.alignment=_al("left"); c.border=_bd()
    ws.row_dimensions[row].height=18

def export_excel(df:pd.DataFrame, out_path:Path):
    """
    Gera planilha de cruzamento com 4 abas e visual profissional:
      📋 Pacientes  — lista completa com separadores por prioridade
      🔴 Urgentes   — top críticos
      📊 Resumo     — mini-dashboard com barras e KPIs
      🟢 Em Dia     — concluídos
    """
    from openpyxl.styles import Font as XFont
    wb=Workbook()
    ind_cols=sorted([c for c in df.columns if re.match(r"^C\d$",c)])
    total_pac=len(df)
    _prio_col = df["Prioridade"].astype(str) if "Prioridade" in df.columns else pd.Series(dtype=str)
    _pend_col = pd.to_numeric(df["Pendências"], errors="coerce").fillna(0) if "Pendências" in df.columns else pd.Series(dtype=float)
    urgentes = int(_prio_col.str.contains("URGENTE", na=False).sum())
    em_dia   = int((_pend_col == 0).sum())
    multi    = int((_pend_col > 1).sum())

    # ═══════════════════════════════════════════════════════════════════
    # ABA 1: Pacientes — lista completa
    # ═══════════════════════════════════════════════════════════════════
    ws=wb.active; ws.title="📋 Pacientes"
    cols=(["Nome","Microárea","Telefone"]+ind_cols
          +["Indicadores","Qtd","Pendências","Média","Prioridade","O que fazer"])
    cols=[c for c in cols if c in df.columns]; n=len(cols)

    # Linha 1 — Título principal
    _title(ws,1,"APS  ·  CRUZAMENTO UNIFICADO POR PACIENTE",C["azul"],C["branco"],15,n,36)

    # Linha 2 — Barra de KPIs inline
    kpi=(f"  👥 {total_pac} pacientes   |   🔴 {urgentes} urgentes   |   "
         f"✔ {em_dia} em dia   |   ⚠ {multi} com 2+ pendências   |   "
         f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    _title(ws,2,kpi,C["azul_c"],C["azul"],9,n,18)

    # Linha 3 — Legenda semântica
    _title(ws,3,
           "  🟢 Concluído = 100 pts     🟡 Monitorar = 75-99     🟠 Alta = 50-74     "
           "🔴 Urgente < 50 pts     ⬜ Cinza = não está nesta lista",
           C["roxo_c"],C["roxo"],8,n,13)

    # Linha 4 — Banner separador
    c4=_mw(ws,4,1,"  ▼  LISTA COMPLETA  —  ordenada por prioridade e número de pendências",ec=n)
    c4.font=_fn(bold=True,color=C["branco"],size=9); c4.fill=_f(C["azul_m"])
    c4.alignment=_al("left"); c4.border=_bd(); ws.row_dimensions[4].height=16

    # Linha 5 — Cabeçalhos (duas faixas de cor: azul p/ identificação, amarelo p/ indicadores, roxo p/ síntese)
    for ci,col in enumerate(cols,1):
        if   col in ("Nome","Microárea","Telefone"): bg=C["azul"];  fg=C["branco"]
        elif col in ind_cols:                        bg=C["amar_m"];fg="1F3864"
        else:                                        bg=C["roxo"];  fg=C["branco"]
        cell=ws.cell(5,ci,col)
        cell.fill=_f(bg); cell.font=_fn(bold=True,color=fg,size=9)
        cell.alignment=_al("center",wrap=True); cell.border=_bd("medium","595959")
    ws.row_dimensions[5].height=42

    # Dados com separadores visuais por grupo de prioridade
    prev_prio=None; dr=6
    PRIO_LABEL={"🔴 URGENTE":"🔴  URGENTE — ação imediata",
                "🟠 ALTA":   "🟠  ALTA — acompanhamento próximo",
                "🟡 MONITORAR":"🟡  MONITORAR — próximo de concluir",
                "🟢 CONCLUÍDO":"🟢  CONCLUÍDO — todas as listas completas"}

    for _,row in df[cols].iterrows():
        prio=str(row.get("Prioridade",""))
        npend=int(row.get("Pendências",0) or 0)

        if prio!=prev_prio and prio in PRIO_THEME:
            bg_s,fg_s=PRIO_THEME[prio]
            try: ws.merge_cells(f"A{dr}:{get_column_letter(n)}{dr}")
            except: pass
            sep=ws.cell(dr,1,f"  {PRIO_LABEL.get(prio,prio)}")
            sep.fill=_f(bg_s); sep.font=_fn(bold=True,color=fg_s,size=9)
            sep.alignment=_al("left"); sep.border=_bd("medium","888888")
            ws.row_dimensions[dr].height=15; dr+=1; prev_prio=prio

        row_bg,row_fg=PRIO_THEME.get(prio,(C["cinza"],C["azul"]))
        if npend==0: row_bg=C["verde_c"]; row_fg=C["verde_e"]
        # Zebra suave dentro do grupo
        is_even=(dr%2==0)

        for ci,col in enumerate(cols,1):
            val=row.get(col,""); vs="" if str(val) in {"nan","None"} else str(val)
            cell=ws.cell(dr,ci,vs); cell.border=_bd()

            if col in ("Nome","Microárea","Telefone"):
                cell.fill=_f("F8FBFF" if is_even else C["branco"])
                cell.font=_fn(bold=(col=="Nome"),size=9,color="1A1A2E")
                cell.alignment=_al("left")

            elif col in ind_cols:
                cell.alignment=_al("center")
                if vs=="—":
                    cell.fill=_f("EEEEEE"); cell.font=_fn(color="BBBBBB",size=8)
                    cell.value="—"
                else:
                    try:
                        pts=float(vs.replace("pts","").strip())
                        if pts>=100:
                            cell.fill=_f(C["verde"]); cell.font=_fn(bold=True,color=C["verde_e"],size=9)
                            cell.value="✔ 100"
                        elif pts>=75:
                            cell.fill=_f(C["amar_m"]); cell.font=_fn(bold=True,color=C["amar_t"],size=9)
                            cell.value=f"⚡{int(pts)}"
                        elif pts>=50:
                            cell.fill=_f(C["amarelo"]); cell.font=_fn(bold=True,color=C["amar_t"],size=9)
                            cell.value=f"⚠{int(pts)}"
                        else:
                            cell.fill=_f(C["verm"]); cell.font=_fn(bold=True,color=C["verm_t"],size=9)
                            cell.value=f"🔴{int(pts)}"
                    except:
                        cell.fill=_f(C["cinza"]); cell.font=_fn(size=9)

            elif col=="Prioridade":
                bg2,fg2=PRIO_THEME.get(vs,(C["cinza"],C["azul"]))
                cell.fill=_f(bg2); cell.font=_fn(bold=True,color=fg2,size=9)
                cell.alignment=_al("center")

            elif col=="O que fazer":
                cell.fill=_f(row_bg)
                cell.font=_fn(size=9,color=C["verm_t"] if npend>0 else C["verde_e"],italic=(npend==0))
                cell.alignment=_al("left",wrap=True)

            elif col in ("Qtd","Pendências","Média"):
                cell.fill=_f(row_bg); cell.font=_fn(bold=True,color=row_fg,size=10)
                cell.alignment=_al("center")

            elif col=="Indicadores":
                cell.fill=_f(row_bg); cell.font=_fn(size=8,color=row_fg)
                cell.alignment=_al("left")

            else:
                cell.fill=_f(row_bg); cell.font=_fn(size=9,color=row_fg); cell.alignment=_al("center")

        nlines=str(row.get("O que fazer","")).count("\n")+1
        ws.row_dimensions[dr].height=max(18,14*nlines); dr+=1

    WD={"Nome":32,"Microárea":11,"Telefone":15,"Indicadores":26,
        "Qtd":6,"Pendências":10,"Média":8,"Prioridade":18,"O que fazer":55}
    for ci,col in enumerate(cols,1):
        ws.column_dimensions[get_column_letter(ci)].width=WD.get(col,13)
    ws.freeze_panes="D6"
    ws.auto_filter.ref=f"A5:{get_column_letter(n)}5"

    # ═══════════════════════════════════════════════════════════════════
    # ABA 2: Urgentes
    # ═══════════════════════════════════════════════════════════════════
    ws2=wb.create_sheet("🔴 Urgentes")
    df2=(df[df["Pendências"]>0].sort_values(["Pendências","Média"],ascending=[False,True]).head(60))
    h2=["Nome","Microárea","Telefone","Indicadores","Qtd","Pendências","Prioridade","O que fazer"]
    h2=[c for c in h2 if c in df2.columns]; n2=len(h2)

    _title(ws2,1,"🔴  PACIENTES MAIS URGENTES  —  maior pendência, menor pontuação","C00000",C["branco"],14,n2,30)
    _title(ws2,2,
           f"Top {len(df2)} pacientes críticos  •  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
           "FFEBEE","C00000",9,n2,14)
    c3s=_mw(ws2,3,1,"  ▼  Priorize: vermelho = urgência imediata  |  cada linha mostra o que precisa ser feito",ec=n2)
    c3s.font=_fn(bold=True,color=C["branco"],size=9); c3s.fill=_f("C00000")
    c3s.alignment=_al("left"); c3s.border=_bd(); ws2.row_dimensions[3].height=15

    for ci,h in enumerate(h2,1):
        bg="2C0F0F" if h in ("Nome","Microárea","Telefone") else "C00000"
        cell=ws2.cell(4,ci,h)
        cell.fill=_f(bg); cell.font=_fn(bold=True,color=C["branco"],size=9)
        cell.alignment=_al("center",wrap=True); cell.border=_bd()
    ws2.row_dimensions[4].height=32

    prev2=None; dr2=5
    for _,row in df2[h2].iterrows():
        prio=str(row.get("Prioridade",""))
        if prio!=prev2 and prio in PRIO_THEME:
            bg_s,fg_s=PRIO_THEME[prio]
            try: ws2.merge_cells(f"A{dr2}:{get_column_letter(n2)}{dr2}")
            except: pass
            s=ws2.cell(dr2,1,f"  {prio}")
            s.fill=_f(bg_s); s.font=_fn(bold=True,color=fg_s,size=9)
            s.alignment=_al("left"); s.border=_bd("medium","888888"); ws2.row_dimensions[dr2].height=14; dr2+=1; prev2=prio
        bg,fg=PRIO_THEME.get(prio,(C["cinza"],C["azul"]))
        for ci,h in enumerate(h2,1):
            vs=str(row.get(h,"") or ""); vs="" if vs in {"nan","None"} else vs
            cell=ws2.cell(dr2,ci,vs); cell.border=_bd()
            if h=="Prioridade":
                bg2,fg2=PRIO_THEME.get(vs,(C["cinza"],C["azul"]))
                cell.fill=_f(bg2); cell.font=_fn(bold=True,color=fg2,size=9); cell.alignment=_al("center")
            elif h=="O que fazer":
                cell.fill=_f(bg); cell.font=_fn(size=9,color=fg); cell.alignment=_al("left",wrap=True)
            elif h=="Nome":
                cell.fill=_f(bg); cell.font=_fn(bold=True,color=fg,size=9); cell.alignment=_al("left")
            else:
                cell.fill=_f(bg); cell.font=_fn(size=9,color=fg)
                cell.alignment=_al("center" if h not in ("Indicadores",) else "left")
        nlines=str(row.get("O que fazer","")).count("\n")+1
        ws2.row_dimensions[dr2].height=max(18,14*nlines); dr2+=1

    for ci,w in enumerate([32,11,15,26,6,10,18,55],1):
        ws2.column_dimensions[get_column_letter(ci)].width=w
    ws2.freeze_panes="A5"

    # ═══════════════════════════════════════════════════════════════════
    # ABA 3: Resumo / mini-dashboard
    # ═══════════════════════════════════════════════════════════════════
    ws3=wb.create_sheet("📊 Resumo")
    _title(ws3,1,"📊  RESUMO POR INDICADOR  —  visão consolidada",C["azul"],C["branco"],14,8,30)
    _title(ws3,2,
           f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  •  "
           f"{len(ind_cols)} indicadores  •  {total_pac} pacientes únicos",
           C["azul_c"],C["azul"],9,8,16)

    # KPI cards
    _section(ws3,3,"  ▼  VISÃO GERAL",C["azul_m"],C["branco"],8)
    def _kpi3(col,lbl,val,bg,fg):
        for r in (4,5):
            try: ws3.merge_cells(f"{get_column_letter(col)}{r}:{get_column_letter(col+1)}{r}")
            except: pass
        cl=ws3.cell(4,col,lbl); cl.fill=_f(bg); cl.font=_fn(bold=True,color=fg,size=8)
        cl.alignment=_al("center"); cl.border=_bd()
        cv=ws3.cell(5,col,val); cv.fill=_f(bg); cv.font=_fn(bold=True,color=fg,size=18)
        cv.alignment=_al("center"); cv.border=_bd()
        ws3.row_dimensions[4].height=14; ws3.row_dimensions[5].height=28
    _kpi3(1,"👥 Pacientes",total_pac,C["azul_h"],C["azul"])
    _kpi3(3,"🔴 Urgentes",urgentes,C["verm_c"],C["verm_e"])
    _kpi3(5,"✔ Em dia",em_dia,C["verde_c"],C["verde_e"])
    _kpi3(7,"⚠ 2+ listas",multi,C["amarelo"],C["amar_t"])

    # Tabela por indicador
    _section(ws3,7,"  ▼  DETALHE POR INDICADOR  —  % de conclusão, pendências e barra de progresso",C["roxo"],C["branco"],8)
    hdr3=["Indicador","Total","Concluídos","Pendentes","Média pts","% Concluídos","Progresso","Status"]
    for ci,h in enumerate(hdr3,1):
        cell=ws3.cell(8,ci,h)
        cell.fill=_f(C["azul"]); cell.font=_fn(bold=True,color=C["branco"],size=9)
        cell.alignment=_al("center",wrap=True); cell.border=_bd()
    ws3.row_dimensions[8].height=30

    for ri,code in enumerate(ind_cols,9):
        vals=[str(v) for v in df[code] if str(v) not in ("—","nan","None","")]
        total=len(vals)
        conc=sum(1 for v in vals if "100" in v or "✔" in v)
        pend=total-conc
        try: media=round(sum(float(re.sub(r"[^\d.]","",v)) for v in vals)/total,1) if total else 0
        except: media=0
        pct=round(conc/total*100,1) if total else 0
        filled=int(pct/10); bar="█"*filled+"░"*(10-filled)
        if pct>=75:   bg=C["verde"];   fg=C["verde_e"];  st="🟢 Bom"
        elif pct>=50: bg=C["amarelo"]; fg=C["amar_t"];   st="🟡 Atenção"
        else:         bg=C["verm"];    fg=C["verm_t"];    st="🔴 Crítico"

        data_row=[code,total,conc,pend,media,f"{pct}%",bar,st]
        for ci,v in enumerate(data_row,1):
            cell=ws3.cell(ri,ci,v); cell.border=_bd()
            if ci==1:
                cell.fill=_f(C["azul_h"]); cell.font=_fn(bold=True,color=C["azul"],size=10)
                cell.alignment=_al("center")
            elif ci==3:
                cell.fill=_f(C["verde_c"]); cell.font=_fn(bold=True,color=C["verde_e"],size=10); cell.alignment=_al("center")
            elif ci==4:
                clr=C["verm_c"] if pend>0 else C["verde_c"]
                fcl=C["verm_e"] if pend>0 else C["verde_e"]
                cell.fill=_f(clr); cell.font=_fn(bold=True,color=fcl,size=10); cell.alignment=_al("center")
            elif ci==7:
                cell.fill=_f(bg); cell.font=XFont(name="Courier New",bold=True,color=fg,size=10)
                cell.alignment=_al("left")
            elif ci==8:
                cell.fill=_f(bg); cell.font=_fn(bold=True,color=fg,size=9); cell.alignment=_al("center")
            else:
                cell.fill=_f(C["cinza"] if ri%2==0 else C["branco"])
                cell.font=_fn(size=9); cell.alignment=_al("center")

    for ci,w in enumerate([12,10,13,12,10,13,22,12],1):
        ws3.column_dimensions[get_column_letter(ci)].width=w
    ws3.freeze_panes="A9"

    # ═══════════════════════════════════════════════════════════════════
    # ABA 4: Em Dia
    # ═══════════════════════════════════════════════════════════════════
    ws4=wb.create_sheet("🟢 Em Dia")
    df4=df[df["Pendências"]==0].sort_values("Nome")
    h4=["Nome","Microárea","Telefone","Indicadores","Qtd","Média"]
    h4=[c for c in h4 if c in df4.columns]; n4=len(h4)

    _title(ws4,1,"🟢  PACIENTES EM DIA  —  sem pendências em nenhuma lista",C["verde_e"],C["branco"],13,n4,28)
    _title(ws4,2,f"{len(df4)} pacientes concluídos  •  {datetime.now().strftime('%d/%m/%Y %H:%M')}",
           C["verde_c"],C["verde_e"],9,n4,14)
    c3v=_mw(ws4,3,1,"  ▼  Estes pacientes atingiram 100 pts em todas as listas em que estão",ec=n4)
    c3v.font=_fn(bold=True,color=C["branco"],size=9); c3v.fill=_f(C["verde_e"])
    c3v.alignment=_al("left"); c3v.border=_bd(); ws4.row_dimensions[3].height=15

    for ci,h in enumerate(h4,1):
        cell=ws4.cell(4,ci,h)
        cell.fill=_f(C["verde_e"]); cell.font=_fn(bold=True,color=C["branco"],size=9)
        cell.alignment=_al("center",wrap=True); cell.border=_bd()
    ws4.row_dimensions[4].height=28

    for ri,(_,row) in enumerate(df4[h4].iterrows(),5):
        for ci,h in enumerate(h4,1):
            vs=str(row.get(h,"") or ""); vs="" if vs in {"nan","None"} else vs
            cell=ws4.cell(ri,ci,vs); cell.border=_bd()
            cell.fill=_f(C["verde_c"] if ri%2==0 else C["branco"])
            cell.font=_fn(bold=(h=="Nome"),size=9,color=C["verde_e"])
            cell.alignment=_al("left" if h in ("Nome","Indicadores") else "center")
        ws4.row_dimensions[ri].height=16

    for ci,w in enumerate([32,11,15,26,6,8],1):
        ws4.column_dimensions[get_column_letter(ci)].width=w
    ws4.freeze_panes="A5"

    wb.save(out_path)


# ── Comparação entre pastas (usado pelo dashboard) ───────────────────────
def _collect_folder(folder: Path) -> dict:
    data = {}
    by_code: dict[str, list[Path]] = {}
    for p in folder.glob("C*.xlsx"):
        if any(x in p.name.lower() for x in ("interativa","unificad","comparad","cruzament")):
            continue
        code = _detect_code(p)
        by_code.setdefault(code, []).append(p)
    for code, paths in by_code.items():
        latest = max(paths, key=lambda p: p.stat().st_mtime)
        df = _read(latest)
        if df is not None and "Nome" in df.columns:
            df["_src"] = latest.name
            data[code] = df
    return data

def build_folder_comparison(folder_a: Path, folder_b: Path,
                             label_a="Período A", label_b="Período B") -> pd.DataFrame:
    da = _collect_folder(folder_a)
    db = _collect_folder(folder_b)
    codes = sorted(set(list(da.keys()) + list(db.keys())))
    rows = []
    for code in codes:
        def _s(df):
            if df is None: return None
            pts = pd.to_numeric(df.get("Pontuação"), errors="coerce").fillna(0)
            cls = df.get("Classificação", pd.Series(dtype=str)).fillna("").astype(str)
            return {
                "total": len(df), "media": round(float(pts.mean()),1) if len(df) else 0.0,
                "busca": int((pts<100).sum()),
                "otimo": int((cls=="Ótimo").sum()), "bom": int((cls=="Bom").sum()),
                "suf":   int((cls=="Suficiente").sum()), "reg": int((cls=="Regular").sum()),
                "arquivo": df["_src"].iloc[0] if "_src" in df.columns and len(df) else "—",
            }
        sa = _s(da.get(code)); sb = _s(db.get(code))
        def _d(k):
            if sa and sb: return sb[k]-sa[k]
            return None
        rows.append({
            "Indicador":           code,
            f"Arquivo {label_a}":  sa["arquivo"] if sa else "—",
            f"Arquivo {label_b}":  sb["arquivo"] if sb else "—",
            f"Total {label_a}":    sa["total"]   if sa else "—",
            f"Total {label_b}":    sb["total"]   if sb else "—",
            f"Média {label_a}":    sa["media"]   if sa else "—",
            f"Média {label_b}":    sb["media"]   if sb else "—",
            "Δ Média":             _d("media"),
            f"Busca {label_a}":    sa["busca"]   if sa else "—",
            f"Busca {label_b}":    sb["busca"]   if sb else "—",
            "Δ Busca":             _d("busca"),
            "Δ Ótimo":             _d("otimo"),
            "Δ Bom":               _d("bom"),
            "Δ Suficiente":        _d("suf"),
            "Δ Regular":           _d("reg"),
        })
    return pd.DataFrame(rows)

def export_folder_comparison_excel(df: pd.DataFrame, out_path: Path,
                                    label_a="Período A", label_b="Período B"):
    from datetime import datetime
    wb = Workbook()
    ws = wb.active; ws.title = "📊 Comparação Pastas"
    n = len(df.columns)
    _title(ws,1,f"APS  —  COMPARAÇÃO ENTRE PASTAS   {label_a}  ×  {label_b}",
           C["azul"],C["branco"],13,n,28)
    _title(ws,2,
           f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  •  "
           f"Verde = melhora  •  Vermelho = piora  •  Δ = diferença B−A",
           C["azul_c"],C["azul"],9,n,14)
    for ci,col in enumerate(df.columns,1):
        if col.startswith("Δ"): bg=C["roxo"]; fg=C["branco"]
        elif label_a in col: bg=C["azul_h"]; fg=C["azul"]
        elif label_b in col: bg=C["amar_m"]; fg=C["azul"]
        else: bg=C["azul_h"]; fg=C["azul"]
        cell=ws.cell(3,ci,col)
        cell.fill=_f(bg); cell.font=_fn(bold=True,color=fg,size=9)
        cell.alignment=_al("center",wrap=True); cell.border=_bd()
    ws.row_dimensions[3].height=40
    for ri,(_,row) in enumerate(df.iterrows(),4):
        for ci,col in enumerate(df.columns,1):
            val=row[col]
            vs="" if str(val) in {"nan","None"} else str(val)
            cell=ws.cell(ri,ci,val if str(val) not in {"nan","None"} else "")
            cell.border=_bd(); cell.alignment=_al("center"); cell.font=_fn(size=9)
            cell.fill=_f(C["cinza"] if ri%2==0 else C["branco"])
            if col.startswith("Δ") and val not in (None,"—",""):
                try:
                    v=float(str(val).replace(",","."))
                    invert=any(x in col for x in ("Busca","Regular","Suf"))
                    good=(v<0) if invert else (v>0)
                    bad=(v>0)  if invert else (v<0)
                    if good:   cell.fill=_f(C["verde"]);  cell.font=_fn(bold=True,color=C["verde_e"],size=9)
                    elif bad:  cell.fill=_f(C["verm"]);   cell.font=_fn(bold=True,color=C["verm_t"],size=9)
                    else:      cell.fill=_f(C["amarelo"]); cell.font=_fn(bold=True,color=C["amar_t"],size=9)
                    cell.value=f"{v:+.1f}" if isinstance(v,float) else f"{int(v):+d}"
                except: pass
    for ci,col in enumerate(df.columns,1):
        ws.column_dimensions[get_column_letter(ci)].width=(32 if "Arquivo" in col else 10 if col=="Indicador" else 14)
    ws.freeze_panes="B4"
    ws.auto_filter.ref=f"A3:{get_column_letter(n)}3"
    wb.save(out_path)

# ── UI ─────────────────────────────────────────────────────────────────────
class ComparadorPacienteApp(tk.Toplevel):
    def __init__(self,master=None,out_dir=None):
        super().__init__(master=master)
        self.title("APS – Cruzamento por Paciente")
        self.geometry("860x480")
        self.minsize(700,380)
        self.configure(bg="#EAF2F8")
        self.out_dir=out_dir or Path.home()/"Desktop"/"APS_RESULTADOS"
        self.files=[]
        self._build()
        self.lift(); self.focus_force(); self.grab_set()

    def _build(self):
        tk.Label(self,text="APS – Cruzamento por Paciente",
                 bg="#1F4E79",fg="white",font=("Segoe UI",13,"bold"),pady=10).pack(fill="x")
        tk.Label(self,
                 text="Selecione as planilhas resultado (C1–C7). O sistema cruza por nome,\n"
                      "mostra em quais listas cada paciente está e o que está pendente para ele.",
                 bg="#EAF2F8",font=("Segoe UI",9),justify="left").pack(fill="x",padx=14,pady=(8,2))

        btn=tk.Frame(self,bg="#EAF2F8"); btn.pack(fill="x",padx=14,pady=6)
        for txt,cmd,bg in [
            ("➕ Adicionar planilhas",        self._add,   "#2E75B6"),
            ("📁 Auto-detectar na pasta",     self._auto,  "#27AE60"),
            ("🗑 Limpar lista",               self._clear, "#C0392B"),
        ]:
            tk.Button(btn,text=txt,command=cmd,bg=bg,fg="white",
                      font=("Segoe UI",9,"bold")).pack(side="left",padx=(0,6))

        box=ttk.LabelFrame(self,text="Planilhas selecionadas")
        box.pack(fill="both",expand=True,padx=14,pady=4)
        box.columnconfigure(0,weight=1); box.rowconfigure(0,weight=1)
        self.tree=ttk.Treeview(box,columns=("cod","arq","mb"),show="headings",height=8)
        self.tree.heading("cod",text="Cód"); self.tree.heading("arq",text="Arquivo"); self.tree.heading("mb",text="MB")
        self.tree.column("cod",width=60,anchor="center"); self.tree.column("arq",width=560); self.tree.column("mb",width=70,anchor="e")
        self.tree.grid(row=0,column=0,sticky="nsew")
        sb=ttk.Scrollbar(box,orient="vertical",command=self.tree.yview)
        sb.grid(row=0,column=1,sticky="ns"); self.tree.configure(yscrollcommand=sb.set)

        bot=tk.Frame(self,bg="#EAF2F8"); bot.pack(fill="x",padx=14,pady=8)
        self.sv=tk.StringVar(value="Nenhuma planilha selecionada.")
        tk.Label(bot,textvariable=self.sv,bg="#EAF2F8",font=("Segoe UI",9),anchor="w").pack(side="left",fill="x",expand=True)
        tk.Button(bot,text="▶  Gerar planilha unificada",command=self._run,
                  bg="#1F4E79",fg="white",font=("Segoe UI",10,"bold")).pack(side="right")

    def _add(self):
        chosen=filedialog.askopenfilenames(title="Planilhas resultado APS",filetypes=[("Excel","*.xlsx")])
        for f in chosen:
            p=Path(f)
            if p not in self.files: self.files.append(p)
        self._refresh()

    def _auto(self):
        folder=filedialog.askdirectory(title="Pasta com resultados APS",initialdir=str(self.out_dir))
        if not folder: return
        found=[p for p in sorted(Path(folder).iterdir())
               if p.suffix.lower()==".xlsx"
               and not any(x in p.name.lower() for x in ("interativa","unificad","comparad","cruzament"))
               and re.search(r"c\d",p.name.lower())]
        if not found: messagebox.showwarning("Nada encontrado","Nenhuma planilha C1–C7 na pasta."); return
        for p in found:
            if p not in self.files: self.files.append(p)
        self._refresh()

    def _clear(self): self.files.clear(); self._refresh()

    def _refresh(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for p in self.files:
            self.tree.insert("","end",values=(_detect_code(p),p.name,f"{p.stat().st_size/1048576:.2f}"))
        n=len(self.files)
        self.sv.set(f"{n} planilha(s) selecionada(s)." if n else "Nenhuma planilha selecionada.")

    def _run(self):
        if not self.files: messagebox.showwarning("Atenção","Adicione ao menos uma planilha."); return
        self.sv.set("Processando… aguarde."); self.update()
        try:
            df=build_unified(self.files)
            if df.empty:
                messagebox.showwarning("Resultado vazio",
                    "Não foi possível cruzar. Verifique se as planilhas têm a aba '📋 Dados'."); return
            stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
            out=self.out_dir/f"CRUZAMENTO_{stamp}.xlsx"
            out.parent.mkdir(parents=True,exist_ok=True)
            export_excel(df,out)
            self.sv.set(f"✔ {out.name}")
            messagebox.showinfo("Concluído ✔",f"Planilha gerada:\n\n{out}\n\nPacientes: {len(df)}",parent=self)
            try: os.startfile(out)
            except: pass
        except Exception as exc:
            import traceback
            self.sv.set(f"Erro: {exc}")
            messagebox.showerror("Erro",f"{exc}\n\n{traceback.format_exc()}",parent=self)

def launch_comparador(master=None,out_dir=None):
    return ComparadorPacienteApp(master=master,out_dir=out_dir)

def main():
    root=tk.Tk(); root.withdraw()
    app=ComparadorPacienteApp(master=root)
    app.protocol("WM_DELETE_WINDOW",root.destroy)
    root.mainloop()

if __name__=="__main__": main()

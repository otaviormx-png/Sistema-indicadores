from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
import json
import tempfile
from pathlib import Path
import re
import sys
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from aps_utils import infer_indicator_code_from_path, INDICATOR_OUTPUT_NAMES


class ScrollableTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, bg="#F4F8FB", highlightthickness=0)
        self.vscroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vscroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vscroll.pack(side="right", fill="y")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window, width=event.width)

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass


CLASS_ORDER = ["Otimo", "Bom", "Suficiente", "Regular"]
SCORE_LEGEND = [
    ("Otimo", "76-100 pontos", "#4E79A7"),
    ("Bom", "51-75 pontos", "#F28E2B"),
    ("Suficiente", "26-50 pontos", "#2CA02C"),
    ("Regular", "1-25 pontos", "#D62728"),
    ("Critico operacional", "0 pontos", "#8B0000"),
]
PRIORITY_ORDER = ["Concluido", "Baixa", "Media", "Alta", "URGENTE", "ALTA", "MEDIA", "BAIXA"]
INDICATOR_SHORT_LABELS = {
    "C1": "C1 Mais acesso",
    "C2": "C2 Desenv. infantil",
    "C3": "C3 Gestacao/puerperio",
    "C4": "C4 Diabetes",
    "C5": "C5 Hipertensao",
    "C6": "C6 Pessoa idosa",
    "C7": "C7 Saude da mulher",
}


def indicator_display_label(code: str) -> str:
    code_up = str(code or "").upper().strip()
    if code_up in INDICATOR_SHORT_LABELS:
        return INDICATOR_SHORT_LABELS[code_up]
    friendly = INDICATOR_OUTPUT_NAMES.get(code_up, code_up).replace(".xlsx", "")
    return f"{code_up} {friendly}".strip()


@dataclass
class Snapshot:
    indicador: str
    arquivo: Path
    momento: datetime | None
    total: int
    busca_ativa: int
    media_pontuacao: float
    classes: dict[str, int]
    prioridades: dict[str, int]
    critico_zero: int


def indicator_files(results_dir: Path) -> dict[str, list[Path]]:
    grouped: dict[str, list[Path]] = {}
    for f in results_dir.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        code = infer_indicator_code_from_path(f)
        if not code:
            continue
        grouped.setdefault(code, []).append(f)
    for code in grouped:
        grouped[code].sort(key=lambda p: p.stat().st_mtime)
    return dict(sorted(grouped.items()))


def _folder_sort_key(folder: Path) -> datetime:
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(folder.name.strip(), fmt)
        except Exception:
            pass
    try:
        return datetime.fromtimestamp(folder.stat().st_mtime)
    except Exception:
        return datetime.min


def _find_previous_results_dir(results_dir: Path) -> Path | None:
    parent = results_dir.parent
    if not parent.exists():
        return None
    cur_resolved = results_dir.resolve()
    cur_key = _folder_sort_key(results_dir)
    cands: list[tuple[datetime, Path]] = []
    for d in parent.iterdir():
        if not d.is_dir():
            continue
        try:
            if d.resolve() == cur_resolved:
                continue
        except Exception:
            continue
        if not indicator_files(d):
            continue
        k = _folder_sort_key(d)
        if k < cur_key:
            cands.append((k, d))
    if not cands:
        return None
    cands.sort(key=lambda t: t[0])
    return cands[-1][1]


def parse_stamp(path: Path) -> datetime | None:
    m = re.search(r"(\d{8}_\d{6})", path.stem)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y%m%d_%H%M%S")
        except ValueError:
            pass
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return None


def read_indicator_dataframe(xlsx_path: Path) -> pd.DataFrame:
    try:
        xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"Arquivo invalido: {xlsx_path.name}") from exc
    sheet = next((s for s in xls.sheet_names if "dados" in str(s).lower()), xls.sheet_names[0])
    for header_try in (2, 1, 0):
        try:
            df = pd.read_excel(xlsx_path, sheet_name=sheet, header=header_try, engine="openpyxl")
            df = df.dropna(how="all")
            if len(df.columns) > 1:
                if "Nome" in df.columns:
                    df = df[df["Nome"].astype(str).str.strip().ne("")]
                return df.reset_index(drop=True)
        except Exception:
            continue
    try:
        return pd.read_excel(xlsx_path, sheet_name=sheet, engine="openpyxl").dropna(how="all").reset_index(drop=True)
    except Exception as exc:
        raise ValueError(f"Arquivo invalido: {xlsx_path.name}") from exc


def _norm_key(value: str) -> str:
    txt = str(value or "").strip().lower()
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"[^a-z0-9]+", "", txt)
    return txt


def _pick_col(df: pd.DataFrame, *aliases: str) -> str | None:
    norm_map = {_norm_key(c): c for c in df.columns}
    for alias in aliases:
        found = norm_map.get(_norm_key(alias))
        if found is not None:
            return found
    return None


def build_snapshot(code: str, path: Path) -> Snapshot:
    df = read_indicator_dataframe(path)
    total = len(df)
    col_pts = _pick_col(df, "Pontuacao", "Pontuação")
    col_cls = _pick_col(df, "Classificacao", "Classificação")
    pontos = pd.to_numeric(df[col_pts] if col_pts else pd.Series(dtype=float), errors="coerce").fillna(0)
    classes = Counter((df[col_cls] if col_cls else pd.Series(dtype=str)).fillna("Sem classificacao").astype(str))
    prioridades = Counter(df.get("Prioridade", pd.Series(dtype=str)).fillna("Sem prioridade").astype(str))
    busca_ativa = int((pontos < 100).sum()) if total else 0
    critico_zero = int((pontos == 0).sum()) if total else 0
    return Snapshot(
        indicador=code,
        arquivo=path,
        momento=parse_stamp(path),
        total=total,
        busca_ativa=busca_ativa,
        media_pontuacao=round(float(pontos.mean()) if total else 0.0, 1),
        classes=dict(classes),
        prioridades=dict(prioridades),
        critico_zero=critico_zero,
    )


def build_current_summary(results_dir: Path, warnings: list[str] | None = None) -> pd.DataFrame:
    rows = []
    for code, files in indicator_files(results_dir).items():
        try:
            snap = build_snapshot(code, files[-1])
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"{code}: falha ao ler {files[-1].name} ({exc})")
            continue
        row = {
            "Indicador": code,
            "Indicador Label": indicator_display_label(code),
            "Arquivo": snap.arquivo.name,
            "Total": snap.total,
            "Busca Ativa": snap.busca_ativa,
            "Media Pontuacao": snap.media_pontuacao,
            "Critico0": snap.critico_zero,
        }
        for cls in CLASS_ORDER:
            row[cls] = snap.classes.get(cls, 0)
        rows.append(row)
    return pd.DataFrame(rows)


def build_comparison_summary(results_dir: Path, warnings: list[str] | None = None) -> pd.DataFrame:
    rows = []
    prev_dir = _find_previous_results_dir(results_dir)
    prev_grouped = indicator_files(prev_dir) if prev_dir else {}
    for code, files in indicator_files(results_dir).items():
        try:
            atual = build_snapshot(code, files[-1])
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"{code}: falha comparacao atual em {files[-1].name} ({exc})")
            continue
        anterior = None
        if len(files) >= 2:
            try:
                anterior = build_snapshot(code, files[-2])
            except Exception as exc:
                if warnings is not None:
                    warnings.append(f"{code}: falha comparacao anterior em {files[-2].name} ({exc})")
        else:
            prev_files = prev_grouped.get(code, [])
            if prev_files:
                try:
                    anterior = build_snapshot(code, prev_files[-1])
                except Exception as exc:
                    if warnings is not None:
                        warnings.append(f"{code}: falha comparacao pasta anterior em {prev_files[-1].name} ({exc})")
        row = {
            "Indicador": code,
            "Indicador Label": indicator_display_label(code),
            "Arquivo Atual": atual.arquivo.name,
            "Atual": atual.media_pontuacao,
            "Anterior": anterior.media_pontuacao if anterior else None,
            "Variacao  Media": round(atual.media_pontuacao - anterior.media_pontuacao, 1) if anterior else None,
            "Variacao  Total": atual.total - anterior.total if anterior else None,
            "Variacao  Busca Ativa": atual.busca_ativa - anterior.busca_ativa if anterior else None,
        }
        for cls in CLASS_ORDER:
            row[f"Variacao  {cls}"] = (atual.classes.get(cls, 0) - anterior.classes.get(cls, 0)) if anterior else None
        rows.append(row)
    return pd.DataFrame(rows)


def load_aprazamento_summary(results_dir: Path) -> dict[str, int]:
    summary = {"total": 0, "vencido": 0, "vermelho": 0, "amarelo": 0, "verde": 0}
    path = results_dir / "aprazamento_controle.json"
    if not path.exists():
        return summary
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return summary
    patients = payload.get("patients", [])
    if not isinstance(patients, list):
        return summary
    for row in patients:
        if not isinstance(row, dict):
            continue
        semaforo = str(row.get("semaforo", "")).strip().upper()
        summary["total"] += 1
        if semaforo == "VENCIDO":
            summary["vencido"] += 1
        elif semaforo == "VERMELHO":
            summary["vermelho"] += 1
        elif semaforo == "AMARELO":
            summary["amarelo"] += 1
        elif semaforo == "VERDE":
            summary["verde"] += 1
    return summary


def build_history(results_dir: Path, indicador: str) -> pd.DataFrame:
    files = indicator_files(results_dir).get(indicador, [])
    rows = []
    for path in files:
        snap = build_snapshot(indicador, path)
        label = snap.momento.strftime("%d/%m %H:%M") if snap.momento else path.stem
        row = {
            "Momento": label,
            "Pontuacao Media": snap.media_pontuacao,
            "Total": snap.total,
            "Busca Ativa": snap.busca_ativa,
        }
        for cls in CLASS_ORDER:
            row[cls] = snap.classes.get(cls, 0)
        rows.append(row)
    return pd.DataFrame(rows)


def build_manual_summary(path: Path) -> tuple[pd.DataFrame, dict]:
    df = read_indicator_dataframe(path)
    col_pts = _pick_col(df, "Pontuacao", "Pontuação")
    col_cls = _pick_col(df, "Classificacao", "Classificação")
    pts = pd.to_numeric(df[col_pts] if col_pts else pd.Series(dtype=float), errors="coerce").fillna(0)
    classes = (df[col_cls] if col_cls else pd.Series(dtype=str)).fillna("Sem classificacao").astype(str)
    counts = {c: int((classes == c).sum()) for c in CLASS_ORDER}
    summary = {
        "arquivo": path.name,
        "total": int(len(df)),
        "media": round(float(pts.mean()) if len(df) else 0.0, 1),
        "busca": int((pts < 100).sum()) if len(df) else 0,
        **counts,
    }
    table = pd.DataFrame({
        "Classificacao": CLASS_ORDER,
        "Quantidade": [counts[c] for c in CLASS_ORDER],
    })
    return table, summary


def _norm_person_name(name: str) -> str:
    txt = str(name or "").strip().lower()
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def count_unique_patients_latest(results_dir: Path) -> int:
    seen: set[str] = set()
    grouped = indicator_files(results_dir)
    for _code, files in grouped.items():
        if not files:
            continue
        latest = files[-1]
        try:
            df = read_indicator_dataframe(latest)
        except Exception:
            continue
        if "Nome" not in df.columns:
            continue
        for raw in df["Nome"].astype(str):
            key = _norm_person_name(raw)
            if key and key != "nan":
                seen.add(key)
    return len(seen)


class APSDashboard(tk.Toplevel):
    def __init__(self, parent: tk.Misc | None, results_dir: Path):
        super().__init__(parent)
        self.results_dir = Path(results_dir)
        self.title("APS - Dashboard Operacional v2")
        try:
            self.state("zoomed")
        except Exception:
            self.geometry("1180x760")
        self.minsize(980, 640)
        self.configure(bg="#F4F8FB")
        self.summary_df = pd.DataFrame()
        self.compare_df = pd.DataFrame()
        self.history_df = pd.DataFrame()
        self.indicator_var = tk.StringVar(value="C1")
        self.folder_var = tk.StringVar(value=str(results_dir))
        self.card_vars = {
            "arquivos": tk.StringVar(value="0"),
            "linhas_brutas": tk.StringVar(value="0"),
            "linhas":   tk.StringVar(value="0"),
            "busca":    tk.StringVar(value="0"),
            "critico_zero": tk.StringVar(value="0"),
            "media":    tk.StringVar(value="0.0"),
            "delta":    tk.StringVar(value="0.0"),
            "delta_busca": tk.StringVar(value="0"),
            "ap_total": tk.StringVar(value="0"),
            "ap_vencido": tk.StringVar(value="0"),
            "ap_alerta": tk.StringVar(value="0"),
        }
        self._refresh_warnings: list[str] = []
        self.path_a = tk.StringVar()
        self.path_b = tk.StringVar()
        self.status_var = tk.StringVar(value="Dashboard pronto.")
        self.status_compare = tk.StringVar(value="Selecione dois arquivos XLSX para comparar.")
        self.compare_insights_var = tk.StringVar(value="Execute a comparacao para ver insights de transicao.")
        self.overview_alerts_var = tk.StringVar(value="Alertas operacionais: atualize o dashboard para gerar os alertas.")
        self.summary_a: dict = {}
        self.summary_b: dict = {}
        self.manual_compare_merged_df = pd.DataFrame()
        self.manual_compare_meta_df = pd.DataFrame()
        self.folder_a_var = tk.StringVar(value="(nao selecionada)")
        self.folder_b_var = tk.StringVar(value="(nao selecionada)")
        self.folder_a_name = tk.StringVar(value="Periodo A")
        self.folder_b_name = tk.StringVar(value="Periodo B")
        self.folders_status = tk.StringVar(value="Selecione as duas pastas para comparar.")
        self.folder_compare_df = pd.DataFrame()
        self.folder_compare_label_a = "Periodo A"
        self.folder_compare_label_b = "Periodo B"
        self.folder_summary_a = pd.DataFrame()
        self.folder_summary_b = pd.DataFrame()
        self.folder_class_totals_a: dict[str, int] = {}
        self.folder_class_totals_b: dict[str, int] = {}
        self.folder_compare_excel_path: Path | None = None
        self.pdf_graph_flags = {
            "panorama": tk.BooleanVar(value=True),
            "busca": tk.BooleanVar(value=True),
            "classificacao": tk.BooleanVar(value=True),
            "risco": tk.BooleanVar(value=True),
            "manual_media": tk.BooleanVar(value=False),
            "manual_classificacao": tk.BooleanVar(value=False),
        }
        self._pdf_checkbuttons: dict[str, ttk.Checkbutton] = {}
        self.action_filter_var = tk.StringVar()
        self.action_sort_var = tk.StringVar(value="Urgencia")
        self.action_class_filter_var = tk.StringVar(value="TODAS")
        self.action_priority_filter_var = tk.StringVar(value="TODAS")
        self.action_indicator_filter_var = tk.StringVar(value="TODOS")
        self.action_simple_mode_var = tk.BooleanVar(value=True)
        self.action_status_var = tk.StringVar(value="Selecione pasta ou arquivos para carregar o painel.")
        self.action_folder_var = tk.StringVar(value=str(results_dir))
        self.action_source_mode = tk.StringVar(value="pasta")
        self.action_source_info_var = tk.StringVar(value=f"Pasta: {results_dir}")
        self.action_insights_var = tk.StringVar(value="Leitura rapida: sem dados carregados.")
        self.action_top_indicators_var = tk.StringVar(value="Top indicadores criticos: sem dados.")
        self.action_top_bairros_var = tk.StringVar(value="Top bairros com maior concentracao: sem dados.")
        self.action_patient_detail_var = tk.StringVar(value="Selecione um paciente na fila para ver detalhes.")
        self.action_selected_files: list[Path] = []
        self._action_row_map: dict[str, dict] = {}
        self.action_view_df = pd.DataFrame()
        self.action_card_vars = {
            "total": tk.StringVar(value="0"),
            "urgente": tk.StringVar(value="0"),
            "alta": tk.StringVar(value="0"),
            "monitorar": tk.StringVar(value="0"),
            "concluido": tk.StringVar(value="0"),
            "media": tk.StringVar(value="0.0"),
        }
        self.unified_df = pd.DataFrame()
        self.unified_path: Path | None = None
        self.auto_unified_refresh_var = tk.BooleanVar(value=True)
        self._last_indicator_signature = None
        self._auto_refresh_after_id = None
        self._build_ui()
        self._set_startup_idle_state()
        self._update_pdf_option_states()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Janela sempre a frente
        self.lift()
        self.focus_force()

    def _set_startup_idle_state(self):
        # Abre rapido e sem varrer planilhas; usuario decide quando carregar.
        self.summary_df = pd.DataFrame()
        self.compare_df = pd.DataFrame()
        self.history_df = pd.DataFrame()
        self.folder_compare_df = pd.DataFrame()
        self.folder_summary_a = pd.DataFrame()
        self.folder_summary_b = pd.DataFrame()
        self.folder_class_totals_a = {}
        self.folder_class_totals_b = {}
        self.folder_compare_excel_path = None
        self.card_vars["arquivos"].set("0")
        self.card_vars["linhas_brutas"].set("0")
        self.card_vars["linhas"].set("0")
        self.card_vars["busca"].set("0")
        self.card_vars["critico_zero"].set("0")
        self.card_vars["media"].set("0.0")
        self.card_vars["delta"].set("0.0")
        self.card_vars["delta_busca"].set("0")
        self.card_vars["ap_total"].set("0")
        self.card_vars["ap_vencido"].set("0")
        self.card_vars["ap_alerta"].set("0")
        self.status_var.set("Dashboard pronto. Clique em Atualizar para carregar os dados.")
        self.action_status_var.set("Painel pronto. Clique em Carregar painel para ler as planilhas.")
        self.action_insights_var.set("Leitura rapida: sem dados carregados.")
        self.action_top_indicators_var.set("Top indicadores criticos: sem dados.")
        self.action_top_bairros_var.set("Top bairros com maior concentracao: sem dados.")
        self.action_patient_detail_var.set("Selecione um paciente na fila para ver detalhes.")
        self.overview_alerts_var.set("Alertas operacionais: atualize o dashboard para gerar os alertas.")
        self._update_pdf_option_states()

    def _build_ui(self):
        header_txt = "Dashboard APS v2 - Operacao e Prioridades"
        header = tk.Label(self, text=header_txt, bg="#1F4E79", fg="white", font=("Segoe UI", 16, "bold"), pady=12)
        header.pack(fill="x")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=12)

        self.tab_overview = ScrollableTab(self.notebook)
        self.tab_compare = ScrollableTab(self.notebook)
        self.tab_folders = ScrollableTab(self.notebook)
        self.notebook.add(self.tab_overview, text="Panorama")
        self.tab_actions = ScrollableTab(self.notebook)
        self.notebook.add(self.tab_actions, text="Painel de acao")
        self.notebook.add(self.tab_compare, text="Comparar arquivos")
        self.notebook.add(self.tab_folders, text="Comparar pastas")

        self._build_overview_tab(self.tab_overview.inner)
        self._build_actions_tab(self.tab_actions.inner)
        self._build_compare_tab(self.tab_compare.inner)
        self._build_folders_tab(self.tab_folders.inner)

    def _build_overview_tab(self, root):
        top = ttk.Frame(root)
        top.pack(fill="x", padx=4, pady=10)
        ttk.Button(top, text="Atualizar", command=self.refresh).pack(side="left")
        ttk.Button(top, text="Exportar relatorio PDF", command=self.exportar_relatorio_pdf).pack(side="left", padx=(8,0))

        # Seletor de fonte (pasta inteira ou arquivo unico)
        ttk.Label(top, text="  Fonte:").pack(side="left", padx=(16,4))
        ttk.Entry(top, textvariable=self.folder_var, width=38).pack(side="left")
        ttk.Button(top, text="Trocar pasta", command=self._change_folder).pack(side="left", padx=(4,16))
        ttk.Button(top, text="Selecionar arquivo", command=self._pick_single_file).pack(side="left", padx=(0, 12))

        ttk.Label(top, text="Historico:").pack(side="left", padx=(0,6))
        self.cbo_indicator = ttk.Combobox(top, textvariable=self.indicator_var, state="readonly", width=8)
        self.cbo_indicator.pack(side="left")
        ttk.Button(top, text="Ver historico", command=self.show_history_window).pack(side="left", padx=(8,0))

        pdf_opts = ttk.LabelFrame(root, text="Graficos no PDF")
        pdf_opts.pack(fill="x", padx=4, pady=(0, 8))
        self._pdf_checkbuttons["panorama"] = ttk.Checkbutton(pdf_opts, text="Grafico 1 (painel principal)", variable=self.pdf_graph_flags["panorama"])
        self._pdf_checkbuttons["panorama"].pack(side="left", padx=(8, 10))
        self._pdf_checkbuttons["busca"] = ttk.Checkbutton(pdf_opts, text="Busca ativa", variable=self.pdf_graph_flags["busca"])
        self._pdf_checkbuttons["busca"].pack(side="left", padx=(0, 10))
        self._pdf_checkbuttons["classificacao"] = ttk.Checkbutton(pdf_opts, text="Distribuicao classificacao", variable=self.pdf_graph_flags["classificacao"])
        self._pdf_checkbuttons["classificacao"].pack(side="left", padx=(0, 10))
        self._pdf_checkbuttons["risco"] = ttk.Checkbutton(pdf_opts, text="Risco por bairro", variable=self.pdf_graph_flags["risco"])
        self._pdf_checkbuttons["risco"].pack(side="left", padx=(0, 10))
        self._pdf_checkbuttons["manual_media"] = ttk.Checkbutton(pdf_opts, text="Comparar arquivos: media", variable=self.pdf_graph_flags["manual_media"])
        self._pdf_checkbuttons["manual_media"].pack(side="left", padx=(0, 10))
        self._pdf_checkbuttons["manual_classificacao"] = ttk.Checkbutton(pdf_opts, text="Comparar arquivos: classificacao", variable=self.pdf_graph_flags["manual_classificacao"])
        self._pdf_checkbuttons["manual_classificacao"].pack(side="left", padx=(0, 10))

        cards = ttk.Frame(root)
        cards.pack(fill="x", padx=4)
        self._make_card(cards, "Arquivos atuais", self.card_vars["arquivos"], 0)
        self._make_card(cards, "Pacientes brutos", self.card_vars["linhas_brutas"], 1)
        self._make_card(cards, "Pacientes unicos", self.card_vars["linhas"], 2)
        self._make_card(cards, "Busca ativa atual", self.card_vars["busca"], 3)
        self._make_card(cards, "Critico (0 pts)", self.card_vars["critico_zero"], 4)
        self._make_card(cards, "Media atual", self.card_vars["media"], 5)
        self._make_card(cards, "Variacao media vs anterior", self.card_vars["delta"], 6)
        self._make_card(cards, "Variacao busca ativa", self.card_vars["delta_busca"], 7)
        self._make_card(cards, "Apraz total", self.card_vars["ap_total"], 8)
        self._make_card(cards, "Apraz vencido", self.card_vars["ap_vencido"], 9)
        self._make_card(cards, "Apraz alerta", self.card_vars["ap_alerta"], 10)

        self._build_score_legend(root)
        self._build_overview_alerts(root)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True, padx=4, pady=12)
        self.overview_body = body
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=0)
        body.rowconfigure(1, weight=1)
        body.rowconfigure(2, weight=1)
        body.rowconfigure(3, weight=1)
        body.rowconfigure(4, weight=1)

        current_box = ttk.LabelFrame(body, text="Resumo atual por indicador")
        current_box.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        current_box.rowconfigure(0, weight=1)
        current_box.columnconfigure(0, weight=1)
        cur_cols = ("Indicador", "Total", "Busca Ativa", "Media Pontuacao", "Otimo", "Bom", "Suficiente", "Regular")
        self.tree_current = ttk.Treeview(current_box, columns=cur_cols, show="headings", height=6)
        for c in cur_cols:
            self.tree_current.heading(c, text=c)
            self.tree_current.column(c, width=100, anchor="center")
        self.tree_current.column("Indicador", width=90)
        self.tree_current.grid(row=0, column=0, sticky="nsew")
        sc1 = ttk.Scrollbar(current_box, orient="vertical", command=self.tree_current.yview)
        sc1.grid(row=0, column=1, sticky="ns")
        self.tree_current.configure(yscrollcommand=sc1.set)
        self.tree_current.bind("<<TreeviewSelect>>", self._on_current_indicator_select)

        compare_box = ttk.LabelFrame(body, text="Comparativo atual x anterior")
        compare_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        compare_box.rowconfigure(0, weight=1)
        compare_box.columnconfigure(0, weight=1)
        cmp_cols = ("Indicador", "Atual", "Anterior", "Variacao  Media", "Variacao  Total", "Variacao  Busca Ativa", "Variacao  Otimo", "Variacao  Bom", "Variacao  Suficiente", "Variacao  Regular")
        self.tree_compare = ttk.Treeview(compare_box, columns=cmp_cols, show="headings", height=6)
        for c in cmp_cols:
            self.tree_compare.heading(c, text=c)
            self.tree_compare.column(c, width=98, anchor="center")
        self.tree_compare.column("Indicador", width=90)
        self.tree_compare.grid(row=0, column=0, sticky="nsew")
        sc2 = ttk.Scrollbar(compare_box, orient="vertical", command=self.tree_compare.yview)
        sc2.grid(row=0, column=1, sticky="ns")
        self.tree_compare.configure(yscrollcommand=sc2.set)

        self.fig1 = Figure(figsize=(6, 3.2), dpi=100)
        self.ax1 = self.fig1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master=body)
        self.canvas1_widget = self.canvas1.get_tk_widget()
        self.canvas1_widget.grid(row=2, column=0, sticky="nsew", padx=(0, 6), pady=(0, 10))

        self.fig2 = Figure(figsize=(6, 3.2), dpi=100)
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, master=body)
        self.canvas2_widget = self.canvas2.get_tk_widget()
        self.canvas2_widget.grid(row=2, column=1, sticky="nsew", padx=(6, 0), pady=(0, 10))

        self.fig3 = Figure(figsize=(6, 3.4), dpi=100)
        self.ax3 = self.fig3.add_subplot(111)
        self.canvas3 = FigureCanvasTkAgg(self.fig3, master=body)
        self.canvas3_widget = self.canvas3.get_tk_widget()
        self.canvas3_widget.grid(row=3, column=0, sticky="nsew", padx=(0, 6))

        self.fig4 = Figure(figsize=(6, 3.4), dpi=100)
        self.ax4 = self.fig4.add_subplot(111)
        self.canvas4 = FigureCanvasTkAgg(self.fig4, master=body)
        self.canvas4_widget = self.canvas4.get_tk_widget()
        self.canvas4_widget.grid(row=3, column=1, sticky="nsew", padx=(6, 0))
        ttk.Label(root, textvariable=self.status_var, anchor="w", relief="sunken").pack(fill="x", padx=4, pady=(0, 8))

    def _build_actions_tab(self, root):
        top = ttk.Frame(root)
        top.pack(fill="x", padx=4, pady=10)
        ttk.Button(top, text="Selecionar pasta", command=self._pick_action_folder).pack(side="left", padx=(0, 6))
        ttk.Button(top, text="Selecionar arquivo(s)", command=self._pick_action_files).pack(side="left", padx=(0, 10))
        ttk.Label(top, text="Filtro:").pack(side="left")
        ent = ttk.Entry(top, textvariable=self.action_filter_var, width=32)
        ent.pack(side="left", padx=(6, 10))
        ent.bind("<KeyRelease>", lambda _e: self._refresh_actions_view())
        ttk.Label(top, text="Classe:").pack(side="left", padx=(0, 4))
        cbo_class = ttk.Combobox(
            top,
            textvariable=self.action_class_filter_var,
            state="readonly",
            width=12,
            values=("TODAS", "Otimo", "Bom", "Suficiente", "Regular"),
        )
        cbo_class.pack(side="left", padx=(0, 8))
        cbo_class.bind("<<ComboboxSelected>>", lambda _e: self._refresh_actions_view())
        ttk.Label(top, text="Prioridade:").pack(side="left", padx=(0, 4))
        cbo_prio = ttk.Combobox(
            top,
            textvariable=self.action_priority_filter_var,
            state="readonly",
            width=12,
            values=("TODAS", "URGENTE", "ALTA", "MONITORAR", "CONCLUIDO"),
        )
        cbo_prio.pack(side="left", padx=(0, 8))
        cbo_prio.bind("<<ComboboxSelected>>", lambda _e: self._refresh_actions_view())
        ttk.Label(top, text="Indicador:").pack(side="left", padx=(0, 4))
        self.cbo_action_indicator = ttk.Combobox(
            top,
            textvariable=self.action_indicator_filter_var,
            state="readonly",
            width=10,
            values=("TODOS",),
        )
        self.cbo_action_indicator.pack(side="left", padx=(0, 8))
        self.cbo_action_indicator.bind("<<ComboboxSelected>>", lambda _e: self._refresh_actions_view())
        ttk.Label(top, text="Ordenar:").pack(side="left")
        cbo = ttk.Combobox(top, textvariable=self.action_sort_var, state="readonly", width=14, values=("Urgencia", "Pontuacao", "Pendencias", "Alfabetica"))
        cbo.pack(side="left", padx=(6, 8))
        cbo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_actions_view())
        ttk.Checkbutton(top, text="Modo simples", variable=self.action_simple_mode_var, command=self._refresh_actions_view).pack(side="left", padx=(6, 0))
        ttk.Button(top, text="Exportar operacional", command=self.export_operational_report).pack(side="right", padx=(8, 0))
        ttk.Button(top, text="Carregar painel", command=self._load_action_data).pack(side="right")

        src = ttk.Frame(root)
        src.pack(fill="x", padx=4, pady=(0, 6))
        ttk.Label(src, textvariable=self.action_source_info_var, anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Label(root, textvariable=self.action_insights_var, anchor="w", foreground="#1F4E79").pack(fill="x", padx=4, pady=(0, 6))

        cards = ttk.Frame(root)
        cards.pack(fill="x", padx=4)
        self._make_card(cards, "Total", self.action_card_vars["total"], 0)
        self._make_card(cards, "Urgente", self.action_card_vars["urgente"], 1)
        self._make_card(cards, "Alta", self.action_card_vars["alta"], 2)
        self._make_card(cards, "Monitorar", self.action_card_vars["monitorar"], 3)
        self._make_card(cards, "Concluido", self.action_card_vars["concluido"], 4)
        self._make_card(cards, "Media", self.action_card_vars["media"], 5)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True, padx=4, pady=10)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        body.rowconfigure(2, weight=0)

        box = ttk.LabelFrame(body, text="Fila operacional (priorize aqui)")
        box.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 6))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        cols = ("Prioridade", "Classificacao", "Nome", "Bairro", "Pendencias", "Media", "Indicadores", "O que fazer")
        self.tree_actions = ttk.Treeview(box, columns=cols, show="headings", height=18)
        for c in cols:
            self.tree_actions.heading(c, text=c)
            self.tree_actions.column(c, width=120, anchor="center")
        self.tree_actions.column("Classificacao", width=105, anchor="center")
        self.tree_actions.column("Nome", width=240, anchor="w")
        self.tree_actions.column("O que fazer", width=280, anchor="w")
        self.tree_actions.tag_configure("prio_urgente", background="#FDECEA")
        self.tree_actions.tag_configure("prio_alta", background="#FFF4E5")
        self.tree_actions.tag_configure("prio_monitorar", background="#FFFBE6")
        self.tree_actions.tag_configure("prio_concluido", background="#EAF7EA")
        self.tree_actions.grid(row=0, column=0, sticky="nsew")
        sca = ttk.Scrollbar(box, orient="vertical", command=self.tree_actions.yview)
        sca.grid(row=0, column=1, sticky="ns")
        self.tree_actions.configure(yscrollcommand=sca.set)
        self.tree_actions.bind("<Double-1>", self._copy_selected_patient)
        self.tree_actions.bind("<<TreeviewSelect>>", self._on_action_select)

        self.action_fig1 = Figure(figsize=(5.5, 3.2), dpi=100)
        self.action_ax1 = self.action_fig1.add_subplot(111)
        self.action_canvas1 = FigureCanvasTkAgg(self.action_fig1, master=body)
        self.action_canvas1.get_tk_widget().grid(row=0, column=1, sticky="nsew", pady=(0, 8))

        self.action_fig2 = Figure(figsize=(5.5, 3.2), dpi=100)
        self.action_ax2 = self.action_fig2.add_subplot(111)
        self.action_canvas2 = FigureCanvasTkAgg(self.action_fig2, master=body)
        self.action_canvas2.get_tk_widget().grid(row=1, column=1, sticky="nsew")

        detail = ttk.Frame(root)
        detail.pack(fill="x", padx=4, pady=(0, 8))
        detail.columnconfigure(0, weight=1)
        detail.columnconfigure(1, weight=1)
        detail.columnconfigure(2, weight=2)
        left = ttk.LabelFrame(detail, text="Top indicadores criticos")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        mid = ttk.LabelFrame(detail, text="Top bairros com maior concentracao")
        mid.grid(row=0, column=1, sticky="nsew", padx=(0, 6))
        right = ttk.LabelFrame(detail, text="Visualizador do paciente")
        right.grid(row=0, column=2, sticky="nsew")
        ttk.Label(left, textvariable=self.action_top_indicators_var, anchor="w", justify="left").pack(fill="x", padx=8, pady=8)
        ttk.Label(mid, textvariable=self.action_top_bairros_var, anchor="w", justify="left").pack(fill="x", padx=8, pady=8)
        ttk.Label(right, textvariable=self.action_patient_detail_var, anchor="w", justify="left").pack(fill="x", padx=8, pady=8)

        ttk.Label(root, textvariable=self.action_status_var, anchor="w", relief="sunken").pack(fill="x", padx=4, pady=(0, 8))

    def _layout_overview_charts(self, folder_compare_mode: bool):
        if not hasattr(self, "canvas1_widget"):
            return
        # Limpa o grid atual dos quatro graficos para reposicionar.
        self.canvas1_widget.grid_forget()
        self.canvas2_widget.grid_forget()
        self.canvas3_widget.grid_forget()
        self.canvas4_widget.grid_forget()

        if folder_compare_mode:
            # Modo apresentacao:
            # - Distribuicao atual sobe para a primeira linha de graficos
            # - Distribuicao comparada (ax1) ocupa duas colunas na linha seguinte
            # - Risco por bairro ocupa duas colunas abaixo
            self.canvas3_widget.grid(row=2, column=0, sticky="nsew", padx=(0, 6), pady=(0, 10))
            self.canvas2_widget.grid(row=2, column=1, sticky="nsew", padx=(6, 0), pady=(0, 10))
            self.canvas1_widget.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=(0, 0), pady=(0, 10))
            self.canvas4_widget.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=(0, 0))
            self.overview_body.rowconfigure(2, weight=1)
            self.overview_body.rowconfigure(3, weight=2)
            self.overview_body.rowconfigure(4, weight=1)
        else:
            # Layout padrao (2x2)
            self.canvas1_widget.grid(row=2, column=0, sticky="nsew", padx=(0, 6), pady=(0, 10))
            self.canvas2_widget.grid(row=2, column=1, sticky="nsew", padx=(6, 0), pady=(0, 10))
            self.canvas3_widget.grid(row=3, column=0, sticky="nsew", padx=(0, 6))
            self.canvas4_widget.grid(row=3, column=1, sticky="nsew", padx=(6, 0))
            self.overview_body.rowconfigure(2, weight=1)
            self.overview_body.rowconfigure(3, weight=1)
            self.overview_body.rowconfigure(4, weight=0)

    def _resolve_action_source_files(self) -> list[Path]:
        if self.action_source_mode.get() == "arquivos" and self.action_selected_files:
            out = []
            for p in self.action_selected_files:
                if not p.exists():
                    continue
                if infer_indicator_code_from_path(p):
                    out.append(p)
            return out
        folder = Path(self.action_folder_var.get().strip() or self.results_dir)
        if not folder.exists():
            return []
        files = []
        for p in folder.glob("*.xlsx"):
            if infer_indicator_code_from_path(p):
                files.append(p)
        files.sort(key=lambda x: x.stat().st_mtime if x.exists() else 0, reverse=True)
        latest_by_code: dict[str, Path] = {}
        for p in files:
            code = infer_indicator_code_from_path(p)
            if not code:
                continue
            if code not in latest_by_code:
                latest_by_code[code] = p
        return sorted(latest_by_code.values(), key=lambda x: x.name.lower())

    def _load_action_data(self, silent: bool = False):
        paths = self._resolve_action_source_files()
        indicator_values = ["TODOS"] + sorted({str(infer_indicator_code_from_path(p) or "").upper() for p in paths if infer_indicator_code_from_path(p)})
        if hasattr(self, "cbo_action_indicator"):
            self.cbo_action_indicator["values"] = tuple(indicator_values)
            if self.action_indicator_filter_var.get().strip().upper() not in indicator_values:
                self.action_indicator_filter_var.set("TODOS")
        if not paths:
            self.unified_df = pd.DataFrame()
            self.action_view_df = pd.DataFrame()
            self._refresh_actions_view()
            if not silent:
                messagebox.showwarning("Sem arquivos", "Nao encontrei planilhas para montar o painel.")
            return
        try:
            from aps_comparador_paciente import build_unified
            raw = build_unified(paths)
            if raw is None or raw.empty:
                self.unified_df = pd.DataFrame()
                self.action_view_df = pd.DataFrame()
                self._refresh_actions_view()
                if not silent:
                    messagebox.showwarning("Sem dados", "Nao foi possivel montar dados com os arquivos selecionados.")
                return
            self.unified_df = self._prepare_unified_df(raw)
            self.unified_path = None
            self._refresh_actions_view()
            self._last_indicator_signature = self._current_indicator_signature()
            src = "arquivos selecionados" if self.action_source_mode.get() == "arquivos" else f"pasta {Path(self.action_folder_var.get()).name}"
            self.action_status_var.set(f"Painel carregado de {src} ({len(paths)} planilhas).")
        except Exception as exc:
            if not silent:
                messagebox.showerror("Erro", str(exc))

    def _pick_action_folder(self):
        initial = self.action_folder_var.get().strip() or str(self.results_dir)
        chosen = filedialog.askdirectory(title="Escolha a pasta com planilhas C1..C7", initialdir=initial)
        if not chosen:
            return
        self.action_source_mode.set("pasta")
        self.action_folder_var.set(chosen)
        self.action_selected_files = []
        self.action_source_info_var.set(f"Pasta: {chosen}")
        self._load_action_data()

    def _pick_action_files(self):
        initial = self.action_folder_var.get().strip() or str(self.results_dir)
        paths = filedialog.askopenfilenames(title="Selecione 1 ou mais planilhas para o painel", initialdir=initial, filetypes=[("Excel", "*.xlsx *.xls")])
        if not paths:
            return
        self.action_source_mode.set("arquivos")
        self.action_selected_files = [Path(p) for p in paths]
        self.action_source_info_var.set(f"Arquivos selecionados: {len(self.action_selected_files)}")
        self._load_action_data()

    def _nkey(self, value: str) -> str:
        txt = str(value or "").strip().lower()
        txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
        txt = re.sub(r"[^a-z0-9]+", "", txt)
        return txt

    def _extract_indicator_codes(self, value: str) -> list[str]:
        return sorted(set(re.findall(r"C\d+", str(value or "").upper())))

    def _norm_bairro(self, value: str) -> str:
        txt = str(value or "").strip()
        key = self._nkey(txt)
        if not txt or key in {"", "nan", "none", "na", "n/a", "sembairro", "sembairro"}:
            return "SEM BAIRRO"
        txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
        txt = re.sub(r"[\-_/]+", " ", txt)
        txt = re.sub(r"[^A-Za-z0-9 ]+", " ", txt)
        txt = re.sub(r"\s+", " ", txt).strip()
        txt = re.sub(r"^BAIRRO\s+", "", txt, flags=re.IGNORECASE)
        tokens = txt.upper().split(" ")
        replace_map = {
            "LGO": "LARGO",
            "LOT": "LOTEAMENTO",
            "CONJ": "CONJUNTO",
            "RES": "RESIDENCIAL",
            "VL": "VILA",
            "JD": "JARDIM",
            "STA": "SANTA",
            "STO": "SANTO",
            "ST": "SANTO",
        }
        txt = " ".join(replace_map.get(tok, tok) for tok in tokens if tok)
        txt = txt.strip()
        return txt or "SEM BAIRRO"

    def _canonicalize_bairro_series(self, series: pd.Series) -> pd.Series:
        if series is None or len(series) == 0:
            return pd.Series(dtype=str)
        norm_vals = series.astype(str).apply(self._norm_bairro)
        keys = norm_vals.apply(lambda v: re.sub(r"[^A-Z0-9]+", "", str(v)))
        keys = keys.replace("", "SEMBAIRRO")
        counts = keys.value_counts()
        if counts.empty:
            return norm_vals

        roots: dict[str, str] = {k: k for k in counts.index.tolist()}
        key_list = counts.index.tolist()
        for i, k1 in enumerate(key_list):
            if k1 == "SEMBAIRRO":
                continue
            for k2 in key_list[i + 1:]:
                if k2 == "SEMBAIRRO":
                    continue
                if roots.get(k2) != k2:
                    continue
                ratio = SequenceMatcher(None, k1, k2).ratio()
                if ratio >= 0.94 or (ratio >= 0.90 and (k1 in k2 or k2 in k1)):
                    roots[k2] = k1

        def _root(k: str) -> str:
            r = k
            while roots.get(r, r) != r:
                r = roots[r]
            return r

        root_keys = keys.apply(_root)
        meta = pd.DataFrame({"bairro": norm_vals.values, "root": root_keys.values})
        root_label: dict[str, str] = {}
        for root_key, grp in meta.groupby("root"):
            root_label[root_key] = str(grp["bairro"].value_counts().index[0])
        return root_keys.map(lambda rk: root_label.get(rk, "SEM BAIRRO"))

    def _resolve_col(self, df: pd.DataFrame, *aliases: str) -> str | None:
        norm_map = {self._nkey(c): c for c in df.columns}
        for a in aliases:
            found = norm_map.get(self._nkey(a))
            if found is not None:
                return found
        return None

    def _prepare_unified_df(self, raw: pd.DataFrame) -> pd.DataFrame:
        if raw is None or raw.empty:
            return pd.DataFrame()
        c_nome = self._resolve_col(raw, "Nome")
        c_prio = self._resolve_col(raw, "Prioridade")
        c_ind = self._resolve_col(raw, "Indicadores")
        c_oqf = self._resolve_col(raw, "O que fazer")
        if not all([c_nome, c_prio, c_ind, c_oqf]):
            return pd.DataFrame()
        c_bairro = self._resolve_col(raw, "Bairro", "Bairro/Localidade", "Localidade", "Microarea", "Microarea")
        c_pend = self._resolve_col(raw, "Pendencias", "Pendencias", "Pendancias")
        c_media = self._resolve_col(raw, "Media", "Media")
        if c_pend is None:
            c_pend = next((c for c in raw.columns if "pend" in self._nkey(c)), None)
        if c_media is None:
            c_media = next((c for c in raw.columns if "medi" in self._nkey(c)), None)

        df = pd.DataFrame({
            "nome": raw[c_nome].astype(str),
            "prioridade": raw[c_prio].astype(str),
            "indicadores": raw[c_ind].astype(str),
            "oqf": raw[c_oqf].astype(str),
            "bairro": raw[c_bairro].astype(str) if c_bairro else "",
            "pendencias": raw[c_pend].astype(str) if c_pend else "",
            "media": raw[c_media].astype(str) if c_media else "",
        }).fillna("")

        for col in ("nome", "prioridade", "indicadores", "oqf", "bairro", "pendencias", "media"):
            df[col] = df[col].astype(str).str.strip()
        df["bairro"] = self._canonicalize_bairro_series(df["bairro"])

        # Remove linhas de legenda/separador: paciente valido precisa ter nome e ao menos um indicador C1..C9.
        mask_valid = (
            df["nome"].ne("")
            & ~df["nome"].str.lower().eq("nan")
            & df["indicadores"].str.contains(r"C\d+", case=False, na=False)
        )
        df = df[mask_valid].copy()
        df["_prio_ord"] = df["prioridade"].apply(lambda v: {"URGENTE": 0, "ALTA": 1, "MONITORAR": 2, "CONCLUIDO": 3}.get(self._norm_prio(v), 9))
        df["_pend_num"] = pd.to_numeric(df.get("pendencias"), errors="coerce").fillna(0)
        df["_name_key"] = df["nome"].apply(_norm_person_name)
        # Mantem a linha mais relevante por paciente (prioridade pior e mais pendencias).
        df = df.sort_values(["_name_key", "_prio_ord", "_pend_num"], ascending=[True, True, False])
        df = df.drop_duplicates(subset=["_name_key"], keep="first")
        df = df.drop(columns=["_prio_ord", "_pend_num", "_name_key"], errors="ignore")
        return df.reset_index(drop=True)

    def _norm_prio(self, value: str) -> str:
        txt = str(value or "").upper()
        if "URGENTE" in txt:
            return "URGENTE"
        if "MONITOR" in txt or "MEDIA" in txt:
            return "MONITORAR"
        if "ALTA" in txt:
            return "ALTA"
        if "CONCL" in txt or "BAIXA" in txt:
            return "CONCLUIDO"
        return "MONITORAR"

    def _priority_distribution_from_latest(self) -> dict[str, int]:
        out = {"URGENTE": 0, "ALTA": 0, "MONITORAR": 0, "CONCLUIDO": 0}
        # Regra principal: distribuicao por paciente unico (sem duplicar entre C1..C7).
        try:
            base_df = self.unified_df
            if base_df is None or base_df.empty:
                from aps_comparador_paciente import build_unified
                paths = self._resolve_action_source_files()
                if paths:
                    raw = build_unified(paths)
                    base_df = self._prepare_unified_df(raw)
            if base_df is not None and not base_df.empty and "prioridade" in base_df.columns:
                prios = base_df["prioridade"].astype(str).apply(self._norm_prio)
                for p in prios:
                    out[p] = out.get(p, 0) + 1
                return out
        except Exception:
            pass

        # Fallback: por linhas das planilhas (pode duplicar pacientes).
        grouped = indicator_files(self.results_dir)
        for _code, files in grouped.items():
            if not files:
                continue
            try:
                df = read_indicator_dataframe(files[-1])
            except Exception:
                continue
            c_prio = _pick_col(df, "Prioridade")
            if not c_prio:
                continue
            for raw in df[c_prio].astype(str):
                p = self._norm_prio(raw)
                out[p] = out.get(p, 0) + 1
        return out

    def _draw_overview_variation_chart(self):
        self.ax1.clear()

        # Comparacao de pastas: torres empilhadas por indicador (C1 anterior x C1 atual, ...).
        if (
            self.folder_summary_a is not None
            and self.folder_summary_b is not None
            and (not self.folder_summary_a.empty or not self.folder_summary_b.empty)
        ):
            self._stacked_folder_compare_chart(
                self.ax1,
                self.folder_summary_a,
                self.folder_summary_b,
                self.folder_compare_label_a,
                self.folder_compare_label_b,
            )
            return

        labels: list[str] = []
        vals: list[float] = []
        title = "Variacao media (sem comparacao disponivel)"

        if self.folder_compare_df is not None and not self.folder_compare_df.empty:
            c_var = next((c for c in self.folder_compare_df.columns if self._nkey(c).startswith("variacaomedia")), None)
            if c_var is None:
                c_var = next((c for c in self.folder_compare_df.columns if self._nkey(c).startswith("deltamedia")), None)
            if c_var and "Indicador" in self.folder_compare_df.columns:
                labels = [indicator_display_label(v) for v in self.folder_compare_df["Indicador"].astype(str).tolist()]
                vals = pd.to_numeric(self.folder_compare_df[c_var], errors="coerce").fillna(0.0).tolist()
                title = f"Variacao media ({self.folder_compare_label_b} - {self.folder_compare_label_a})"

        if (not labels) and (self.compare_df is not None) and (not self.compare_df.empty):
            c_var = "Variacao  Media" if "Variacao  Media" in self.compare_df.columns else ("Delta  Media" if "Delta  Media" in self.compare_df.columns else None)
            if c_var:
                labels = self.compare_df.get("Indicador Label", self.compare_df.get("Indicador", pd.Series(dtype=str))).astype(str).tolist()
                vals = pd.to_numeric(self.compare_df[c_var], errors="coerce").fillna(0.0).tolist()
                title = "Variacao media atual x anterior"

        if labels:
            colors = ["#1B5E20" if v >= 0 else "#B71C1C" for v in vals]
            bars = self.ax1.bar(labels, vals, color=colors)
            self.ax1.axhline(0, color="#6B6B6B", linewidth=1)
            self.ax1.bar_label(bars, fmt="%+.1f", padding=2, fontsize=8)
            self.ax1.set_ylabel("Pontos")
            self.ax1.tick_params(axis="x", labelrotation=18)
        self.ax1.set_title(title)
        self.ax1.set_xlabel("Indicador")

    def _load_folder_class_totals(self, folder: Path) -> dict[str, int]:
        try:
            summary = build_current_summary(folder)
        except Exception:
            return {}
        if summary is None or summary.empty:
            return {}
        out: dict[str, int] = {}
        for cls in CLASS_ORDER:
            if cls in summary.columns:
                vals = pd.to_numeric(summary[cls], errors="coerce").fillna(0)
                out[cls] = int(vals.sum())
            else:
                out[cls] = 0
        return out

    def _stacked_folder_compare_chart(self, ax, summary_a: pd.DataFrame, summary_b: pd.DataFrame, label_a: str, label_b: str):
        ax.clear()
        if summary_a is None:
            summary_a = pd.DataFrame()
        if summary_b is None:
            summary_b = pd.DataFrame()

        map_a = {}
        map_b = {}
        if not summary_a.empty and "Indicador" in summary_a.columns:
            map_a = {str(row["Indicador"]).strip().upper(): row for _, row in summary_a.iterrows()}
        if not summary_b.empty and "Indicador" in summary_b.columns:
            map_b = {str(row["Indicador"]).strip().upper(): row for _, row in summary_b.iterrows()}
        codes = sorted(set(map_a.keys()) | set(map_b.keys()))
        if not codes:
            ax.set_title("Distribuicao por classificacao (pastas comparadas): sem dados")
            return False

        labels = []
        x = []
        pos = 0.0
        short_a = str(label_a or "Anterior").strip()
        short_b = str(label_b or "Atual").strip()
        for code in codes:
            x.extend([pos, pos + 0.42])
            labels.extend([f"{code}\n{short_a}", f"{code}\n{short_b}"])
            pos += 1.15

        class_colors = {
            "Otimo": "#4E79A7",
            "Bom": "#F28E2B",
            "Suficiente": "#2CA02C",
            "Regular": "#D62728",
        }
        bottom_a = [0] * len(codes)
        bottom_b = [0] * len(codes)

        for cls in CLASS_ORDER:
            vals_a = [int(pd.to_numeric(pd.Series([map_a.get(code, {}).get(cls, 0)]), errors="coerce").fillna(0).iloc[0]) for code in codes]
            vals_b = [int(pd.to_numeric(pd.Series([map_b.get(code, {}).get(cls, 0)]), errors="coerce").fillna(0).iloc[0]) for code in codes]
            ax.bar(
                [x[i * 2] for i in range(len(codes))],
                vals_a,
                width=0.36,
                bottom=bottom_a,
                color=class_colors.get(cls, "#7F7F7F"),
                label=cls if cls not in ax.get_legend_handles_labels()[1] else None,
            )
            ax.bar(
                [x[i * 2 + 1] for i in range(len(codes))],
                vals_b,
                width=0.36,
                bottom=bottom_b,
                color=class_colors.get(cls, "#7F7F7F"),
            )
            bottom_a = [bottom_a[i] + vals_a[i] for i in range(len(codes))]
            bottom_b = [bottom_b[i] + vals_b[i] for i in range(len(codes))]

        totals_a = bottom_a[:]
        totals_b = bottom_b[:]
        max_total = float(max(totals_a + totals_b) if (totals_a or totals_b) else 0.0)
        ax.set_xticks(x, labels)
        ax.set_title("Distribuicao por classificacao (pastas comparadas)")
        ax.set_ylabel("Pacientes")
        ax.set_xlabel("Indicador / periodo")
        ax.tick_params(axis="x", labelrotation=16)
        ax.set_ylim(0.0, max(1.0, max_total * 1.18))
        ax.legend(fontsize=8, ncols=3, loc="upper right")

        for i, code in enumerate(codes):
            xt = (x[i * 2] + x[i * 2 + 1]) / 2
            ax.text(xt, max(totals_a[i], totals_b[i]) + max(4.0, max_total * 0.02), code, ha="center", va="bottom", fontsize=8, color="#1F4E79", fontweight="bold")

        for i in range(len(codes)):
            ax.text(x[i * 2], totals_a[i] + max(3.0, max_total * 0.012), str(int(totals_a[i])), ha="center", va="bottom", fontsize=7, color="#1F1F1F")
            ax.text(x[i * 2 + 1], totals_b[i] + max(3.0, max_total * 0.012), str(int(totals_b[i])), ha="center", va="bottom", fontsize=7, color="#1F1F1F")

        return True

    def _risk_by_bairro_from_latest(self) -> pd.Series:
        by_patient: dict[str, dict] = {}
        grouped = indicator_files(self.results_dir)
        score_map = {"URGENTE": 3, "ALTA": 2, "MONITORAR": 1, "CONCLUIDO": 0}

        for _code, files in grouped.items():
            if not files:
                continue
            try:
                df = read_indicator_dataframe(files[-1])
            except Exception:
                continue
            c_nome = _pick_col(df, "Nome")
            c_bairro = _pick_col(df, "Bairro", "Bairro/Localidade", "Localidade", "Microarea", "Microárea", "Microarea")
            c_prio = _pick_col(df, "Prioridade")
            c_pts = _pick_col(df, "Pontuacao", "Pontuação")
            if not c_nome:
                continue

            for _, row in df.iterrows():
                nome_raw = str(row.get(c_nome, "") or "").strip()
                key = _norm_person_name(nome_raw)
                if not key or key == "nan":
                    continue
                bairro = self._norm_bairro(row.get(c_bairro, "") if c_bairro else "")

                if c_prio:
                    ptxt = self._norm_prio(str(row.get(c_prio, "") or ""))
                else:
                    pts = pd.to_numeric(pd.Series([row.get(c_pts, 0) if c_pts else 0]), errors="coerce").fillna(0).iloc[0]
                    if pts >= 100:
                        ptxt = "CONCLUIDO"
                    elif pts >= 75:
                        ptxt = "MONITORAR"
                    elif pts >= 50:
                        ptxt = "ALTA"
                    else:
                        ptxt = "URGENTE"
                score = int(score_map.get(ptxt, 0))

                cur = by_patient.get(key)
                if cur is None or score > int(cur.get("score", 0)):
                    by_patient[key] = {"score": score, "bairro": bairro}
                elif cur.get("bairro", "SEM BAIRRO") == "SEM BAIRRO" and bairro != "SEM BAIRRO":
                    cur["bairro"] = bairro

        if not by_patient:
            return pd.Series(dtype=float)
        temp = pd.DataFrame(list(by_patient.values()))
        if temp.empty or "bairro" not in temp.columns:
            return pd.Series(dtype=float)
        temp["bairro"] = self._canonicalize_bairro_series(temp["bairro"].astype(str))
        out = temp.groupby("bairro")["score"].sum().sort_values(ascending=False)
        out = out[out > 0]
        return out.head(10)
    def _sort_action_df(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        work = df.copy()
        work["_pend_num"] = pd.to_numeric(work.get("pendencias"), errors="coerce").fillna(0)
        work["_media_num"] = pd.to_numeric(work.get("media"), errors="coerce").fillna(0)
        work["_prio_ord"] = work.get("prioridade", "").apply(
            lambda v: {"URGENTE": 0, "ALTA": 1, "MONITORAR": 2, "CONCLUIDO": 3}.get(self._norm_prio(v), 9)
        )
        mode = self.action_sort_var.get().strip()
        if mode == "Pontuacao":
            return work.sort_values(["_media_num", "_prio_ord", "nome"], ascending=[True, True, True])
        if mode == "Pendencias":
            return work.sort_values(["_pend_num", "_prio_ord", "nome"], ascending=[False, True, True])
        if mode == "Alfabetica":
            return work.sort_values(["nome", "_prio_ord"], ascending=[True, True])
        return work.sort_values(["_prio_ord", "_pend_num", "_media_num", "nome"], ascending=[True, False, True, True])

    def _refresh_actions_view(self):
        if not hasattr(self, "tree_actions"):
            return
        for item in self.tree_actions.get_children():
            self.tree_actions.delete(item)
        self._action_row_map = {}
        if self.unified_df.empty:
            self.action_view_df = pd.DataFrame()
            self.action_status_var.set("Sem dados carregados. Selecione pasta ou arquivos.")
            for key in self.action_card_vars:
                self.action_card_vars[key].set("0" if key != "media" else "0.0")
            self.action_insights_var.set("Leitura rapida: sem dados carregados.")
            self.action_top_indicators_var.set("Top indicadores criticos: sem dados.")
            self.action_top_bairros_var.set("Top bairros com maior concentracao: sem dados.")
            self.action_patient_detail_var.set("Selecione um paciente na fila para ver detalhes.")
            self._draw_action_charts(pd.DataFrame())
            return

        df = self.unified_df.copy()
        df["_media_num"] = pd.to_numeric(df.get("media"), errors="coerce").fillna(0)
        df["_class"] = df["_media_num"].apply(self._class_from_media)
        df["_prio_norm"] = df.get("prioridade", pd.Series(dtype=str)).astype(str).apply(self._norm_prio)
        if self.action_simple_mode_var.get():
            self.tree_actions.configure(displaycolumns=("Prioridade", "Classificacao", "Nome", "Bairro", "O que fazer"))
        else:
            self.tree_actions.configure(displaycolumns=("Prioridade", "Classificacao", "Nome", "Bairro", "Pendencias", "Media", "Indicadores", "O que fazer"))
        filtro_raw = self.action_filter_var.get().strip()
        filtro = self._nkey(filtro_raw)
        if filtro:
            mask = (
                df.get("nome", "").astype(str).apply(lambda v: filtro in self._nkey(v))
                | df.get("bairro", "").astype(str).apply(lambda v: filtro in self._nkey(v))
                | df.get("indicadores", "").astype(str).apply(lambda v: filtro in self._nkey(v))
                | df.get("oqf", "").astype(str).apply(lambda v: filtro in self._nkey(v))
            )
            df = df[mask]
        class_filter = self.action_class_filter_var.get().strip().upper()
        if class_filter and class_filter != "TODAS":
            df = df[df["_class"].astype(str).str.upper() == class_filter]
        prio_filter = self.action_priority_filter_var.get().strip().upper()
        if prio_filter and prio_filter != "TODAS":
            df = df[df["_prio_norm"].astype(str).str.upper() == prio_filter]
        ind_filter = self.action_indicator_filter_var.get().strip().upper()
        if ind_filter and ind_filter != "TODOS":
            df = df[df.get("indicadores", "").astype(str).str.contains(rf"\b{re.escape(ind_filter)}\b", case=False, na=False)]
        df = self._sort_action_df(df)
        self.action_view_df = df.copy()

        prios = df["_prio_norm"]
        self.action_card_vars["total"].set(str(len(df)))
        self.action_card_vars["urgente"].set(str(int((prios == "URGENTE").sum())))
        self.action_card_vars["alta"].set(str(int((prios == "ALTA").sum())))
        self.action_card_vars["monitorar"].set(str(int((prios == "MONITORAR").sum())))
        self.action_card_vars["concluido"].set(str(int((prios == "CONCLUIDO").sum())))
        media = df["_media_num"]
        media_val = float(media.mean() if len(media) else 0)
        self.action_card_vars["media"].set(f"{media_val:.1f}")
        urg = int((prios == "URGENTE").sum())
        alta = int((prios == "ALTA").sum())
        busca = int((df["_media_num"] < 100).sum())
        zero_pts = int((df["_media_num"] == 0).sum())
        self.action_insights_var.set(
            f"Leitura rapida | Urgentes: {urg} | Altas: {alta} | Busca ativa: {busca} | Critico (0): {zero_pts} | Media geral: {media_val:.1f}"
        )

        top_limit = 180 if self.action_simple_mode_var.get() else 300
        top = df.head(top_limit)
        for i, (_, row) in enumerate(top.iterrows(), start=1):
            iid = f"r{i}"
            prio_norm = self._norm_prio(row.get("prioridade", ""))
            tag = {
                "URGENTE": "prio_urgente",
                "ALTA": "prio_alta",
                "MONITORAR": "prio_monitorar",
                "CONCLUIDO": "prio_concluido",
            }.get(prio_norm, "prio_monitorar")
            self.tree_actions.insert("", "end", iid=iid, values=(
                prio_norm,
                row.get("_class", ""),
                row.get("nome", ""),
                row.get("bairro", ""),
                row.get("pendencias", ""),
                row.get("media", ""),
                row.get("indicadores", ""),
                str(row.get("oqf", "")).replace("\n", " | "),
            ), tags=(tag,))
            self._action_row_map[iid] = row.to_dict()

        pend = pd.to_numeric(df.get("pendencias"), errors="coerce").fillna(0)
        pend_df = df[pend > 0].copy()
        ind_counts: dict[str, int] = {}
        for raw in pend_df.get("indicadores", pd.Series(dtype=str)).astype(str):
            for code in re.findall(r"C\d+", raw.upper()):
                ind_counts[code] = ind_counts.get(code, 0) + 1
        if ind_counts:
            ordered_ind = sorted(ind_counts.items(), key=lambda x: x[1], reverse=True)[:8]
            self.action_top_indicators_var.set(
                "\n".join(f"{idx+1}. {code}: {count}" for idx, (code, count) in enumerate(ordered_ind))
            )
        else:
            self.action_top_indicators_var.set("Sem pendencias por indicador no filtro.")

        risk_df = df.copy()
        risk_df["_prio"] = risk_df["prioridade"].astype(str).apply(self._norm_prio)
        risk_df["_score"] = risk_df["_prio"].map({"URGENTE": 3, "ALTA": 2, "MONITORAR": 1, "CONCLUIDO": 0}).fillna(0)
        bcol = risk_df.get("bairro", pd.Series(index=risk_df.index, dtype=str)).astype(str)
        bcol = self._canonicalize_bairro_series(bcol)
        bairros = risk_df.groupby(bcol)["_score"].sum().sort_values(ascending=False)
        bairros = bairros[bairros > 0]
        if not bairros.empty:
            ordered_b = list(bairros.head(8).items())
            self.action_top_bairros_var.set(
                "\n".join(f"{idx+1}. {bairro}: risco {int(score)}" for idx, (bairro, score) in enumerate(ordered_b))
            )
        else:
            self.action_top_bairros_var.set("Sem concentracao de risco por bairro no filtro.")

        if self.action_source_mode.get() == "arquivos":
            base = f"{len(self.action_selected_files)} arquivo(s)"
        else:
            base = f"pasta {Path(self.action_folder_var.get().strip() or self.results_dir).name}"
        self.action_status_var.set(f"Fila carregada de {base} | {len(df)} paciente(s) no filtro | exibindo {len(top)}.")
        self._draw_action_charts(df)
        if self.tree_actions.get_children():
            first = self.tree_actions.get_children()[0]
            self.tree_actions.selection_set(first)
            self._on_action_select()

    def _on_action_select(self, _evt=None):
        sel = self.tree_actions.selection()
        if not sel:
            self.action_patient_detail_var.set("Selecione um paciente na fila para ver detalhes.")
            return
        row = self._action_row_map.get(sel[0], {})
        if not row:
            self.action_patient_detail_var.set("Detalhe indisponivel para o item selecionado.")
            return
        nome = str(row.get("nome", "")).strip()
        bairro = str(row.get("bairro", "")).strip()
        prio = self._norm_prio(row.get("prioridade", ""))
        classe = str(row.get("_class", "")).strip() or self._class_from_media(row.get("media", 0))
        media = str(row.get("media", "")).strip()
        pend = str(row.get("pendencias", "")).strip()
        inds = str(row.get("indicadores", "")).strip()
        oqf = str(row.get("oqf", "")).strip().replace("\n", " | ")
        self.action_patient_detail_var.set(
            f"{nome}\n"
            f"Bairro: {bairro or 'SEM BAIRRO'} | Prioridade: {prio} | Classificacao: {classe}\n"
            f"Media: {media} | Pendencias: {pend}\n"
            f"Indicadores: {inds}\n"
            f"Acao: {oqf}"
        )

    def _draw_action_charts(self, df: pd.DataFrame):
        self.action_ax1.clear()
        self.action_ax2.clear()
        if df.empty:
            self.action_ax1.set_title("Sem dados para prioridade")
            self.action_ax2.set_title("Sem dados para pendencias por indicador")
            self.action_canvas1.draw()
            self.action_canvas2.draw()
            return

        prios = df.get("prioridade", pd.Series(dtype=str)).astype(str).apply(self._norm_prio)
        order = ["URGENTE", "ALTA", "MONITORAR", "CONCLUIDO"]
        vals = [int((prios == p).sum()) for p in order]
        self.action_ax1.barh(order, vals, color=["#C62828", "#EF6C00", "#9E7D00", "#2E7D32"])
        self.action_ax1.set_title("Distribuicao por prioridade")
        self.action_ax1.set_xlabel("Pacientes")

        pend = pd.to_numeric(df.get("pendencias"), errors="coerce").fillna(0)
        pend_df = df[pend > 0].copy()
        counts: dict[str, int] = {}
        for raw in pend_df.get("indicadores", pd.Series(dtype=str)).astype(str):
            for code in re.findall(r"C\d+", raw.upper()):
                counts[code] = counts.get(code, 0) + 1
        if counts:
            ordered = sorted(counts.items(), key=lambda x: x[1], reverse=True)[:10]
            keys = [x[0] for x in ordered][::-1]
            vals2 = [x[1] for x in ordered][::-1]
            self.action_ax2.barh(keys, vals2, color="#2E75B6")
        self.action_ax2.set_title("Top indicadores criticos")
        self.action_ax2.set_xlabel("Pacientes com pendencia")

        self.action_fig1.tight_layout()
        self.action_fig2.tight_layout()
        self.action_canvas1.draw()
        self.action_canvas2.draw()

    def _current_indicator_signature(self):
        parts = []
        if self.action_source_mode.get() == "arquivos" and self.action_selected_files:
            for p in sorted(self.action_selected_files, key=lambda x: x.name.lower()):
                if not p.exists():
                    continue
                code = infer_indicator_code_from_path(p) or p.stem.upper()
                try:
                    st = p.stat()
                    parts.append((code, p.name, int(st.st_mtime_ns), int(st.st_size)))
                except Exception:
                    continue
            return tuple(parts)

        grouped = indicator_files(Path(self.action_folder_var.get().strip() or self.results_dir))
        for code, files in sorted(grouped.items()):
            if not files:
                continue
            p = files[-1]
            try:
                st = p.stat()
                parts.append((code, p.name, int(st.st_mtime_ns), int(st.st_size)))
            except Exception:
                continue
        return tuple(parts)

    def _schedule_auto_unified_refresh(self):
        if self._auto_refresh_after_id:
            try:
                self.after_cancel(self._auto_refresh_after_id)
            except Exception:
                pass
            self._auto_refresh_after_id = None
        self._auto_refresh_after_id = self.after(5000, self._auto_unified_refresh_tick)

    def _auto_unified_refresh_tick(self):
        self._auto_refresh_after_id = None
        if not self.winfo_exists():
            return
        if not hasattr(self, "tree_actions"):
            self._schedule_auto_unified_refresh()
            return
        now_sig = self._current_indicator_signature()
        if self._last_indicator_signature is None:
            self._last_indicator_signature = now_sig
        elif now_sig != self._last_indicator_signature:
            self.action_status_var.set("Mudanca detectada nos indicadores. Recarregando painel...")
            self._last_indicator_signature = now_sig
            self._load_action_data(silent=True)
        self._schedule_auto_unified_refresh()

    def _on_close(self):
        if self._auto_refresh_after_id:
            try:
                self.after_cancel(self._auto_refresh_after_id)
            except Exception:
                pass
            self._auto_refresh_after_id = None
        self.destroy()

    def _copy_selected_patient(self, _evt=None):
        sel = self.tree_actions.selection()
        if not sel:
            return
        vals = self.tree_actions.item(sel[0], "values")
        if not vals:
            return
        name = str(vals[1])
        self.clipboard_clear()
        self.clipboard_append(name)
        self.action_status_var.set(f"Paciente copiado: {name}")

    def _build_compare_tab(self, root):
        top = ttk.Frame(root)
        top.pack(fill="x", padx=4, pady=10)
        ttk.Label(top, text="Arquivo A:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.path_a, width=100).grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        ttk.Button(top, text="Selecionar", command=self._pick_a).grid(row=0, column=2, padx=4)
        ttk.Label(top, text="Arquivo B:").grid(row=1, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.path_b, width=100).grid(row=1, column=1, padx=6, pady=4, sticky="ew")
        ttk.Button(top, text="Selecionar", command=self._pick_b).grid(row=1, column=2, padx=4)
        ttk.Button(top, text="Comparar", command=self.compare_manual).grid(row=0, column=3, rowspan=2, padx=(12, 0), sticky="ns")
        ttk.Button(top, text="Exportar comparacao", command=self.export_manual_comparison).grid(row=0, column=4, rowspan=2, padx=(8, 0), sticky="ns")
        top.columnconfigure(1, weight=1)

        cards = ttk.Frame(root)
        cards.pack(fill="x", padx=4)
        self.card_a_total = self._make_card(cards, "A - Total", tk.StringVar(value='-'), 0)
        self.card_b_total = self._make_card(cards, "B - Total", tk.StringVar(value='-'), 1)
        self.card_a_media = self._make_card(cards, "A - Media", tk.StringVar(value='-'), 2)
        self.card_b_media = self._make_card(cards, "B - Media", tk.StringVar(value='-'), 3)
        self.card_delta_manual = self._make_card(cards, "Variacao media", tk.StringVar(value='-'), 4)
        self.card_busca_manual = self._make_card(cards, "Variacao busca ativa", tk.StringVar(value='-'), 5)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True, padx=4, pady=12)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        box_table = ttk.LabelFrame(body, text="Comparacao por classificacao")
        box_table.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 8))
        box_table.rowconfigure(0, weight=1)
        box_table.columnconfigure(0, weight=1)
        cols = ("Classificacao", "Arquivo A", "Arquivo B", "Variacao", "Variacao %")
        self.tree_manual = ttk.Treeview(box_table, columns=cols, show="headings")
        for c in cols:
            self.tree_manual.heading(c, text=c)
            self.tree_manual.column(c, width=110, anchor="center")
        self.tree_manual.column("Classificacao", width=160)
        self.tree_manual.tag_configure("delta_up", foreground="#1B5E20")
        self.tree_manual.tag_configure("delta_down", foreground="#B71C1C")
        self.tree_manual.tag_configure("delta_same", foreground="#1F4E79")
        self.tree_manual.grid(row=0, column=0, sticky="nsew")
        sc = ttk.Scrollbar(box_table, orient="vertical", command=self.tree_manual.yview)
        sc.grid(row=0, column=1, sticky="ns")
        self.tree_manual.configure(yscrollcommand=sc.set)

        self.figm1 = Figure(figsize=(5.5, 3.6), dpi=100)
        self.axm1 = self.figm1.add_subplot(111)
        self.canvasm1 = FigureCanvasTkAgg(self.figm1, master=body)
        self.canvasm1.get_tk_widget().grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 8))

        self.figm2 = Figure(figsize=(11.2, 3.4), dpi=100)
        self.axm2 = self.figm2.add_subplot(111)
        self.canvasm2 = FigureCanvasTkAgg(self.figm2, master=body)
        self.canvasm2.get_tk_widget().grid(row=1, column=0, columnspan=2, sticky="nsew")

        insight_box = ttk.LabelFrame(root, text="Resumo da comparacao")
        insight_box.pack(fill="x", padx=4, pady=(0, 8))
        ttk.Label(insight_box, textvariable=self.compare_insights_var, anchor="w", justify="left").pack(fill="x", padx=8, pady=8)

        ttk.Label(root, textvariable=self.status_compare, anchor="w", relief="sunken").pack(fill="x", padx=4, pady=(0, 8))

    def _make_card(self, parent, title, var, col):
        frame = tk.Frame(parent, bg="#DCE6F1", bd=1, relief="solid")
        frame.grid(row=0, column=col, sticky="ew", padx=6, pady=6)
        parent.columnconfigure(col, weight=1)
        tk.Label(frame, text=title, bg="#DCE6F1", fg="#1F1F1F", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(frame, textvariable=var, bg="#DCE6F1", fg="#1F4E79", font=("Segoe UI", 18, "bold")).pack(anchor="w", padx=12, pady=(0, 10))
        return var

    def _build_score_legend(self, root):
        legend = ttk.LabelFrame(root, text="Legenda de pontuacao e cores")
        legend.pack(fill="x", padx=4, pady=(0, 8))
        tk.Label(
            legend,
            text="Classificacao oficial em 4 niveis + destaque operacional para 0 ponto:",
            bg="#F4F8FB",
            fg="#1F1F1F",
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", padx=8, pady=(6, 4))
        row = tk.Frame(legend, bg="#F4F8FB")
        row.pack(fill="x", padx=8, pady=(0, 8))
        for idx, (label, faixa, color) in enumerate(SCORE_LEGEND):
            item = tk.Frame(row, bg="#F4F8FB")
            item.grid(row=0, column=idx, sticky="w", padx=(0, 12))
            tk.Label(item, width=2, bg=color, relief="solid", bd=1).pack(side="left", padx=(0, 6))
            tk.Label(item, text=f"{label}: {faixa}", bg="#F4F8FB", fg="#1F1F1F", font=("Segoe UI", 9)).pack(side="left")
        for idx in range(len(SCORE_LEGEND)):
            row.columnconfigure(idx, weight=1)

    def _build_overview_alerts(self, root):
        box = ttk.LabelFrame(root, text="Alertas operacionais")
        box.pack(fill="x", padx=4, pady=(0, 8))
        ttk.Label(box, textvariable=self.overview_alerts_var, anchor="w", justify="left").pack(fill="x", padx=8, pady=8)

    def _update_pdf_option_states(self):
        overview_ready = self.summary_df is not None and not self.summary_df.empty
        manual_ready = self.manual_compare_merged_df is not None and not self.manual_compare_merged_df.empty
        availability = {
            "panorama": overview_ready,
            "busca": overview_ready,
            "classificacao": overview_ready,
            "risco": overview_ready,
            "manual_media": manual_ready,
            "manual_classificacao": manual_ready,
        }
        for key, var in self.pdf_graph_flags.items():
            enabled = availability.get(key, True)
            if not enabled:
                var.set(False)
            chk = self._pdf_checkbuttons.get(key)
            if chk is not None:
                chk.configure(state=("normal" if enabled else "disabled"))

    def _summary_row_from_snapshot(self, code: str, snap: Snapshot) -> dict:
        row = {
            "Indicador": code,
            "Indicador Label": indicator_display_label(code),
            "Arquivo": snap.arquivo.name,
            "Total": snap.total,
            "Busca Ativa": snap.busca_ativa,
            "Media Pontuacao": snap.media_pontuacao,
            "Critico0": snap.critico_zero,
        }
        for cls in CLASS_ORDER:
            row[cls] = int(snap.classes.get(cls, 0))
        return row

    def _compare_row_from_snapshots(self, code: str, atual: Snapshot, anterior: Snapshot) -> dict:
        row = {
            "Indicador": code,
            "Indicador Label": indicator_display_label(code),
            "Arquivo Atual": atual.arquivo.name,
            "Atual": atual.media_pontuacao,
            "Anterior": anterior.media_pontuacao,
            "Variacao  Media": round(atual.media_pontuacao - anterior.media_pontuacao, 1),
            "Variacao  Total": atual.total - anterior.total,
            "Variacao  Busca Ativa": atual.busca_ativa - anterior.busca_ativa,
        }
        for cls in CLASS_ORDER:
            row[f"Variacao  {cls}"] = int(atual.classes.get(cls, 0)) - int(anterior.classes.get(cls, 0))
        return row

    def _build_single_file_summary(self, source_file: Path, warnings: list[str] | None = None) -> pd.DataFrame:
        try:
            code = (infer_indicator_code_from_path(source_file) or source_file.stem).upper()
            snap = build_snapshot(code, source_file)
            return pd.DataFrame([self._summary_row_from_snapshot(code, snap)])
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"{source_file.name}: falha ao ler ({exc})")
            return pd.DataFrame()

    def _build_single_file_comparison(self, source_file: Path, warnings: list[str] | None = None) -> pd.DataFrame:
        code = (infer_indicator_code_from_path(source_file) or "").upper()
        if not code:
            return pd.DataFrame()
        try:
            atual = build_snapshot(code, source_file)
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"{code}: falha comparacao no arquivo atual ({exc})")
            return pd.DataFrame()
        grouped = indicator_files(source_file.parent)
        files = grouped.get(code, [])
        if len(files) < 2:
            return pd.DataFrame()
        cur_resolved = source_file.resolve()
        prev_path = None
        for idx, candidate in enumerate(files):
            if candidate.resolve() != cur_resolved:
                continue
            if idx > 0:
                prev_path = files[idx - 1]
            break
        if prev_path is None:
            prev_path = files[-2]
        try:
            anterior = build_snapshot(code, prev_path)
            return pd.DataFrame([self._compare_row_from_snapshots(code, atual, anterior)])
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"{code}: falha ao ler comparativo anterior ({exc})")
            return pd.DataFrame()

    def _pick_single_file(self):
        initial = str(Path(self.folder_var.get().strip()).parent) if Path(self.folder_var.get().strip()).is_file() else (self.folder_var.get().strip() or str(self.results_dir))
        selected = filedialog.askopenfilename(
            title="Selecione um arquivo de indicador (C1..C7)",
            initialdir=initial,
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        if not selected:
            return
        self.folder_var.set(selected)
        self.refresh()

    def _class_from_media(self, value) -> str:
        pts = pd.to_numeric(pd.Series([value]), errors="coerce").fillna(0).iloc[0]
        if pts > 75:
            return "Otimo"
        if pts > 50:
            return "Bom"
        if pts > 25:
            return "Suficiente"
        return "Regular"

    def _change_folder(self):
        folder = filedialog.askdirectory(
            title="Selecione a pasta com os resultados APS",
            initialdir=str(self.results_dir))
        if folder:
            self.results_dir = Path(folder)
            self.folder_var.set(folder)
            self.refresh()

    def _on_current_indicator_select(self, _evt=None):
        sel = self.tree_current.selection()
        if not sel:
            return
        code = str(sel[0]).strip().upper()
        if not re.fullmatch(r"C\d+", code):
            return
        self.action_indicator_filter_var.set(code)
        if self.unified_df is None or self.unified_df.empty:
            self._load_action_data(silent=True)
        self._refresh_actions_view()
        try:
            self.notebook.select(self.tab_actions)
        except Exception:
            pass
        self.action_status_var.set(f"Filtro aplicado via panorama: {code}.")

    def _refresh_overview_alerts(self):
        if self.summary_df is None or self.summary_df.empty:
            self.overview_alerts_var.set("Alertas operacionais: sem dados carregados.")
            return
        lines: list[str] = []

        crit = int(pd.to_numeric(self.summary_df.get("Critico0"), errors="coerce").fillna(0).sum())
        lines.append(f"Pacientes criticos (0 pts): {crit}")

        if self.compare_df is not None and not self.compare_df.empty:
            d_media = pd.to_numeric(self.compare_df.get("Variacao  Media"), errors="coerce")
            d_busca = pd.to_numeric(self.compare_df.get("Variacao  Busca Ativa"), errors="coerce")
            if not d_media.dropna().empty:
                worst_idx = d_media.fillna(0).idxmin()
                worst_row = self.compare_df.loc[worst_idx]
                worst_lbl = str(worst_row.get("Indicador Label", worst_row.get("Indicador", "-")))
                worst_val = float(pd.to_numeric(pd.Series([d_media.loc[worst_idx]]), errors="coerce").fillna(0).iloc[0])
                lines.append(f"Maior queda de media: {worst_lbl} ({worst_val:+.1f})")
            pos_busca = self.compare_df[d_busca > 0] if d_busca is not None else pd.DataFrame()
            if pos_busca is not None and not pos_busca.empty:
                top_busca = pos_busca.sort_values("Variacao  Busca Ativa", ascending=False).iloc[0]
                b_lbl = str(top_busca.get("Indicador Label", top_busca.get("Indicador", "-")))
                b_val = int(pd.to_numeric(pd.Series([top_busca.get("Variacao  Busca Ativa", 0)]), errors="coerce").fillna(0).iloc[0])
                lines.append(f"Aumento de busca ativa: {b_lbl} (+{b_val})")

        bairros = self._risk_by_bairro_from_latest()
        if bairros is not None and not bairros.empty:
            top_bairro = str(bairros.index[0])
            top_risk = int(float(bairros.iloc[0]))
            lines.append(f"Bairro com maior risco: {top_bairro} (score {top_risk})")

        self.overview_alerts_var.set("\n".join(lines) if lines else "Alertas operacionais: sem achados relevantes.")

    def refresh(self):
        source_path = Path(self.folder_var.get().strip() or self.results_dir)
        single_file_mode = source_path.is_file()
        if single_file_mode:
            self.results_dir = source_path.parent
            source_code = (infer_indicator_code_from_path(source_path) or source_path.stem).upper()
            files_grouped = {source_code: [source_path]}
        else:
            self.results_dir = source_path
            files_grouped = indicator_files(self.results_dir)
        self._last_indicator_signature = self._current_indicator_signature()
        codes = list(files_grouped.keys())
        self.cbo_indicator["values"] = codes or ["C1"]
        if self.indicator_var.get() not in codes and codes:
            self.indicator_var.set(codes[0])

        warnings: list[str] = []
        if single_file_mode:
            self.summary_df = self._build_single_file_summary(source_path, warnings=warnings)
            self.compare_df = self._build_single_file_comparison(source_path, warnings=warnings)
        else:
            self.summary_df = build_current_summary(self.results_dir, warnings=warnings)
            self.compare_df = build_comparison_summary(self.results_dir, warnings=warnings)
        ap_summary = load_aprazamento_summary(self.results_dir)
        self.card_vars["ap_total"].set(str(ap_summary["total"]))
        self.card_vars["ap_vencido"].set(str(ap_summary["vencido"]))
        self.card_vars["ap_alerta"].set(str(ap_summary["vermelho"] + ap_summary["amarelo"]))
        self._refresh_warnings = warnings
        self._load_action_data(silent=True)

        for tree in (self.tree_current, self.tree_compare):
            for item in tree.get_children():
                tree.delete(item)

        if self.summary_df.empty:
            self._layout_overview_charts(folder_compare_mode=False)
            for key, value in [("arquivos", "0"), ("linhas_brutas", "0"), ("linhas", "0"), ("busca", "0"), ("critico_zero", "0"), ("media", "0.0"), ("delta", "0.0"), ("delta_busca", "0")]:
                self.card_vars[key].set(value)
            for ax, canvas, title in [
                (self.ax1, self.canvas1, "Sem dados"),
                (self.ax2, self.canvas2, "Sem dados"),
                (self.ax3, self.canvas3, "Sem dados"),
                (self.ax4, self.canvas4, "Sem dados"),
            ]:
                ax.clear(); ax.set_title(title); canvas.draw()
            self._refresh_actions_view()
            msg = "Sem dados na pasta selecionada."
            if warnings:
                msg = f"{msg} {len(warnings)} arquivo(s) ignorado(s)."
            self.status_var.set(msg)
            self._refresh_overview_alerts()
            self._update_pdf_option_states()
            return

        for _, row in self.summary_df.iterrows():
            row_code = str(row.get("Indicador", "")).strip().upper()
            values = (
                row.get("Indicador Label", row.get("Indicador")), row.get("Total"), row.get("Busca Ativa"), row.get("Media Pontuacao"),
                row.get("Otimo", 0), row.get("Bom", 0), row.get("Suficiente", 0), row.get("Regular", 0)
            )
            if re.fullmatch(r"C\d+", row_code):
                self.tree_current.insert("", "end", iid=row_code, values=values)
            else:
                self.tree_current.insert("", "end", values=values)

        if not self.compare_df.empty:
            for _, row in self.compare_df.iterrows():
                vals = [row.get("Indicador Label", row.get("Indicador"))] + [row.get(c) for c in ["Atual", "Anterior", "Variacao  Media", "Variacao  Total", "Variacao  Busca Ativa", "Variacao  Otimo", "Variacao  Bom", "Variacao  Suficiente", "Variacao  Regular"]]
                self.tree_compare.insert("", "end", values=tuple("-" if pd.isna(v) else v for v in vals))

        self.card_vars["arquivos"].set(str(len(self.summary_df)))
        self.card_vars["linhas_brutas"].set(str(int(self.summary_df["Total"].sum())))
        self.card_vars["linhas"].set(str(count_unique_patients_latest(self.results_dir)))
        self.card_vars["busca"].set(str(int(self.summary_df["Busca Ativa"].sum())))
        self.card_vars["critico_zero"].set(str(int(pd.to_numeric(self.summary_df.get("Critico0"), errors="coerce").fillna(0).sum())))
        self.card_vars["media"].set(f"{self.summary_df['Media Pontuacao'].mean():.1f}")

        valid_delta = pd.to_numeric(self.compare_df.get("Variacao  Media"), errors="coerce") if not self.compare_df.empty else pd.Series(dtype=float)
        valid_busca = pd.to_numeric(self.compare_df.get("Variacao  Busca Ativa"), errors="coerce") if not self.compare_df.empty else pd.Series(dtype=float)
        delta_mean = float(valid_delta.dropna().mean()) if not valid_delta.dropna().empty else 0.0
        delta_busca = int(round(float(valid_busca.dropna().sum()))) if not valid_busca.dropna().empty else 0
        self.card_vars["delta"].set(f"{delta_mean:+.1f}")
        self.card_vars["delta_busca"].set(f"{delta_busca:+d}")

        self._refresh_overview_charts()
        self._refresh_overview_alerts()
        self._refresh_actions_view()
        src_label = f"arquivo {source_path.name}" if single_file_mode else f"{len(self.summary_df)} indicador(es)"
        msg = f"Panorama atualizado com {src_label}."
        if warnings:
            msg = f"{msg} {len(warnings)} arquivo(s) ignorado(s)."
        self.status_var.set(msg)
        self._update_pdf_option_states()
        self._schedule_auto_unified_refresh()

    def _refresh_overview_charts(self):
        folder_compare_mode = (
            self.folder_summary_a is not None
            and self.folder_summary_b is not None
            and (not self.folder_summary_a.empty or not self.folder_summary_b.empty)
        )
        self._layout_overview_charts(folder_compare_mode=folder_compare_mode)

        self.ax2.clear()
        self.ax3.clear()
        self.ax4.clear()

        # Grafico 1: variacao (comparacao de periodos/arquivos) no lugar da distribuicao por prioridade.
        self._draw_overview_variation_chart()

        labels = self.summary_df["Indicador"].apply(indicator_display_label)
        bars_busca = self.ax2.bar(labels, self.summary_df["Busca Ativa"])
        self.ax2.set_title("Busca ativa atual por indicador")
        self.ax2.set_xlabel("Indicador")
        self.ax2.set_ylabel("Pendentes")
        self.ax2.tick_params(axis="x", labelrotation=18)
        max_busca = float(pd.to_numeric(self.summary_df["Busca Ativa"], errors="coerce").fillna(0).max() if len(self.summary_df) else 0.0)
        self.ax2.set_ylim(0.0, max(1.0, max_busca * 1.12))
        self.ax2.bar_label(bars_busca, fmt="%d", padding=2, fontsize=8)

        stacked_cols = [c for c in CLASS_ORDER if c in self.summary_df.columns]
        bottom = None
        for cls in stacked_cols:
            vals = self.summary_df[cls]
            self.ax3.bar(labels, vals, bottom=bottom, label=cls)
            bottom = vals if bottom is None else bottom + vals
        self.ax3.set_title("Distribuicao atual por classificacao")
        self.ax3.set_xlabel("Indicador")
        self.ax3.set_ylabel("Pacientes")
        self.ax3.tick_params(axis="x", labelrotation=18)
        self.ax3.legend(fontsize=8)
        if bottom is not None:
            totals = pd.to_numeric(bottom, errors="coerce").fillna(0)
            max_total = float(totals.max() if len(totals) else 0.0)
            self.ax3.set_ylim(0.0, max(1.0, max_total * 1.16))
            for xi, tv in enumerate(totals.tolist()):
                y_txt = tv + max(2.0, max_total * 0.015)
                self.ax3.text(xi, y_txt, f"{int(tv)}", ha="center", va="bottom", fontsize=8, color="#1F4E79")

        # Risco por bairro visivel mesmo sem carregar o painel de acao.
        borough = pd.Series(dtype=float)
        if self.unified_df is not None and not self.unified_df.empty:
            risk_df = self.unified_df.copy()
            risk_df["_prio"] = risk_df["prioridade"].astype(str).apply(self._norm_prio)
            risk_df["_score"] = risk_df["_prio"].map({"URGENTE": 3, "ALTA": 2, "MONITORAR": 1, "CONCLUIDO": 0}).fillna(0)
            bcol = risk_df.get("bairro", pd.Series(index=risk_df.index, dtype=str)).astype(str)
            bcol = self._canonicalize_bairro_series(bcol)
            borough = risk_df.groupby(bcol)["_score"].sum().sort_values(ascending=False).head(10)
            borough = borough[borough > 0]
        if borough.empty:
            borough = self._risk_by_bairro_from_latest()
        if not borough.empty:
            keys = list(borough.index)[::-1]
            vals = [float(borough[k]) for k in keys]
            colors = ["#FDECEA" if v >= 25 else "#FFF4E5" if v >= 15 else "#FFFBE6" if v >= 8 else "#EAF7EA" for v in vals]
            self.ax4.barh(keys, vals, color=colors)
            self.ax4.set_title("Risco por bairro (top 10)")
            self.ax4.set_xlabel("Score de risco")
        else:
            self.ax4.set_title("Risco por bairro: sem dados")

        self.fig1.tight_layout()
        self.canvas1.draw()
        self.fig2.tight_layout()
        self.canvas2.draw()
        self.fig3.tight_layout()
        self.canvas3.draw()
        self.fig4.tight_layout()
        self.canvas4.draw()

    def _pick_a(self):
        path = filedialog.askopenfilename(title="Selecione o Arquivo A", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.path_a.set(path)

    def _pick_b(self):
        path = filedialog.askopenfilename(title="Selecione o Arquivo B", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.path_b.set(path)

    def _norm_class(self, value: str) -> str:
        txt = str(value or "").strip().lower()
        txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
        base = re.sub(r"\s+", " ", txt).strip()
        if "otimo" in base:
            return "Otimo"
        if "bom" in base:
            return "Bom"
        if "suficiente" in base:
            return "Suficiente"
        if "regular" in base:
            return "Regular"
        return "Regular"

    def _manual_metrics_by_patient(self, path: Path) -> dict[str, dict]:
        df = read_indicator_dataframe(path)
        c_nome = _pick_col(df, "Nome") or "Nome"
        c_pts = _pick_col(df, "Pontuacao", "Pontuação")
        c_cls = _pick_col(df, "Classificacao", "Classificação")
        if c_nome not in df.columns:
            return {}
        pts = pd.to_numeric(df[c_pts] if c_pts else pd.Series(index=df.index, dtype=float), errors="coerce").fillna(0)
        cls = df[c_cls].astype(str).fillna("") if c_cls else pd.Series(index=df.index, dtype=str)
        out: dict[str, dict] = {}
        for i, raw_name in enumerate(df[c_nome].astype(str)):
            key = _norm_person_name(raw_name)
            if not key or key == "nan":
                continue
            p = float(pts.iloc[i]) if i < len(pts) else 0.0
            c = self._norm_class(cls.iloc[i] if i < len(cls) else "")
            item = {"nome": str(raw_name).strip(), "pts": p, "cls": c, "busca": p < 100}
            prev = out.get(key)
            if prev is None or p > float(prev.get("pts", 0)):
                out[key] = item
        return out

    def _build_compare_insights(self, path_a: Path, path_b: Path) -> str:
        a = self._manual_metrics_by_patient(path_a)
        b = self._manual_metrics_by_patient(path_b)
        common = set(a.keys()) & set(b.keys())
        class_rank = {c: i for i, c in enumerate(CLASS_ORDER)}

        entered_busca = 0
        left_busca = 0
        improved = 0
        worsened = 0
        deltas = []
        for k in common:
            pa = a[k]
            pb = b[k]
            if not pa["busca"] and pb["busca"]:
                entered_busca += 1
            if pa["busca"] and not pb["busca"]:
                left_busca += 1
            ra = class_rank.get(pa["cls"], 9)
            rb = class_rank.get(pb["cls"], 9)
            if rb < ra:
                improved += 1
            elif rb > ra:
                worsened += 1
            deltas.append((pb["pts"] - pa["pts"], pb["nome"]))

        deltas.sort(key=lambda x: x[0], reverse=True)
        best = [f"{name} ({delta:+.1f})" for delta, name in deltas[:3] if delta > 0]
        worst = [f"{name} ({delta:+.1f})" for delta, name in sorted(deltas, key=lambda x: x[0])[:3] if delta < 0]
        best_txt = ", ".join(best) if best else "sem ganhos relevantes"
        worst_txt = ", ".join(worst) if worst else "sem quedas relevantes"
        return (
            f"Entradas na busca ativa: {entered_busca} | Saidas da busca ativa: {left_busca}\n"
            f"Melhoraram classificacao: {improved} | Pioraram classificacao: {worsened}\n"
            f"Top melhoras: {best_txt}\n"
            f"Top quedas: {worst_txt}"
        )

    def compare_manual(self):
        try:
            path_a = Path(self.path_a.get())
            path_b = Path(self.path_b.get())
            if not path_a.exists() or not path_b.exists():
                raise FileNotFoundError("Selecione dois arquivos XLSX validos.")
            table_a, self.summary_a = build_manual_summary(path_a)
            table_b, self.summary_b = build_manual_summary(path_b)

            merged = table_a.merge(table_b, on="Classificacao", how="outer", suffixes=(" A", " B")).fillna(0)
            merged["Arquivo A"] = merged["Quantidade A"].astype(int)
            merged["Arquivo B"] = merged["Quantidade B"].astype(int)
            merged["Variacao"] = merged["Arquivo B"] - merged["Arquivo A"]
            base_a = merged["Arquivo A"].replace(0, pd.NA)
            merged["Variacao %"] = ((merged["Variacao"] / base_a) * 100).round(1)
            merged = merged[["Classificacao", "Arquivo A", "Arquivo B", "Variacao", "Variacao %"]]
            self.manual_compare_merged_df = merged.copy()
            self.manual_compare_meta_df = pd.DataFrame([{
                "Arquivo A": str(path_a),
                "Arquivo B": str(path_b),
                "Total A": int(self.summary_a.get("total", 0)),
                "Total B": int(self.summary_b.get("total", 0)),
                "Media A": float(self.summary_a.get("media", 0)),
                "Media B": float(self.summary_b.get("media", 0)),
                "Variacao Media": float(self.summary_b.get("media", 0)) - float(self.summary_a.get("media", 0)),
                "Busca A": int(self.summary_a.get("busca", 0)),
                "Busca B": int(self.summary_b.get("busca", 0)),
                "Variacao Busca": int(self.summary_b.get("busca", 0)) - int(self.summary_a.get("busca", 0)),
            }])

            for item in self.tree_manual.get_children():
                self.tree_manual.delete(item)
            for _, row in merged.iterrows():
                delta_v = int(row.get("Variacao", 0))
                tag = "delta_same"
                if delta_v > 0:
                    tag = "delta_up"
                elif delta_v < 0:
                    tag = "delta_down"
                delta_pct = row.get("Variacao %")
                delta_pct_txt = "-" if pd.isna(delta_pct) else f"{float(delta_pct):+.1f}%"
                self.tree_manual.insert(
                    "",
                    "end",
                    values=[
                        row.get("Classificacao", ""),
                        int(row.get("Arquivo A", 0)),
                        int(row.get("Arquivo B", 0)),
                        f"{delta_v:+d}",
                        delta_pct_txt,
                    ],
                    tags=(tag,),
                )

            self.card_a_total.set(str(self.summary_a["total"]))
            self.card_b_total.set(str(self.summary_b["total"]))
            self.card_a_media.set(f"{self.summary_a['media']:.1f}")
            self.card_b_media.set(f"{self.summary_b['media']:.1f}")
            self.card_delta_manual.set(f"{self.summary_b['media'] - self.summary_a['media']:+.1f}")
            self.card_busca_manual.set(f"{self.summary_b['busca'] - self.summary_a['busca']:+d}")

            self.axm1.clear()
            medias = [float(self.summary_a["media"]), float(self.summary_b["media"])]
            bars = self.axm1.bar(["A", "B"], medias, color=["#2E75B6", "#ED7D31"])
            self.axm1.set_title("Pontuacao media")
            self.axm1.set_ylabel("Pontos")
            lo = min(medias)
            hi = max(medias)
            span = max(1.0, hi - lo)
            pad = max(0.8, span * 0.9)
            y0 = max(0.0, lo - pad)
            y1 = min(100.0, hi + pad)
            if y1 - y0 < 2.0:
                y1 = min(100.0, y0 + 2.0)
            self.axm1.set_ylim(y0, y1)
            self.axm1.bar_label(bars, fmt="%.1f", padding=3)
            delta_media = medias[1] - medias[0]
            delta_color = "#1B5E20" if delta_media > 0 else ("#B71C1C" if delta_media < 0 else "#1F4E79")
            self.axm1.text(
                0.5,
                0.96,
                f"Variacao: {delta_media:+.1f}",
                transform=self.axm1.transAxes,
                ha="center",
                va="top",
                fontsize=10,
                fontweight="bold",
                color=delta_color,
            )
            self.canvasm1.draw_idle()

            self.axm2.clear()
            order_rank = {c: i for i, c in enumerate(CLASS_ORDER)}
            merged_plot = merged.copy()
            merged_plot["_ord"] = merged_plot["Classificacao"].map(order_rank).fillna(99)
            merged_plot = merged_plot.sort_values("_ord")
            cls = merged_plot["Classificacao"].tolist()
            deltas = merged_plot["Variacao"].astype(float).tolist()
            colors = ["#1B5E20" if d > 0 else "#B71C1C" if d < 0 else "#1F4E79" for d in deltas]
            self.axm2.axvline(0, color="#6B6B6B", linewidth=1)
            bars2 = self.axm2.barh(cls, deltas, color=colors)
            self.axm2.bar_label(bars2, fmt="%+.0f", padding=4)
            self.axm2.set_title("Variacao por classificacao (B - A)")
            self.axm2.set_xlabel("Variacao de pacientes")
            self.canvasm2.draw_idle()
            self.compare_insights_var.set(self._build_compare_insights(path_a, path_b))

            self._update_pdf_option_states()
            self.status_compare.set("Comparacao concluida com sucesso.")
        except Exception as exc:
            self._update_pdf_option_states()
            messagebox.showerror("Erro", str(exc))
            self.status_compare.set(f"Erro: {exc}")

    def export_manual_comparison(self):
        try:
            if self.manual_compare_merged_df.empty:
                raise RuntimeError("Execute a comparacao antes de exportar.")
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            initial_name = f"COMPARACAO_ARQUIVOS_{stamp}.xlsx"
            out = filedialog.asksaveasfilename(
                title="Salvar comparacao de arquivos",
                defaultextension=".xlsx",
                initialdir=str(self.results_dir),
                initialfile=initial_name,
                filetypes=[("Excel", "*.xlsx")],
            )
            if not out:
                return
            out_path = Path(out)
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                self.manual_compare_meta_df.to_excel(writer, sheet_name="Resumo", index=False)
                self.manual_compare_merged_df.to_excel(writer, sheet_name="Classificacao", index=False)
            generated = [str(out_path)]
            if hasattr(self, "figm1"):
                p1 = out_path.with_name(out_path.stem + "_grafico_media.png")
                self.figm1.savefig(p1, dpi=180, bbox_inches="tight")
                generated.append(str(p1))
            if hasattr(self, "figm2"):
                p2 = out_path.with_name(out_path.stem + "_grafico_classificacao.png")
                self.figm2.savefig(p2, dpi=180, bbox_inches="tight")
                generated.append(str(p2))
            self.status_compare.set(f"Comparacao exportada: {out_path.name}")
            messagebox.showinfo("Exportacao concluida", "Arquivos gerados:\n\n" + "\n".join(generated))
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))
            self.status_compare.set(f"Erro ao exportar: {exc}")

    def export_operational_report(self):
        try:
            base_df = self.action_view_df.copy() if self.action_view_df is not None and not self.action_view_df.empty else self.unified_df.copy()
            if base_df is None or base_df.empty:
                raise RuntimeError("Carregue o painel operacional antes de exportar.")

            base_df["_media_num"] = pd.to_numeric(base_df.get("media"), errors="coerce").fillna(0)
            base_df["_class"] = base_df["_media_num"].apply(self._class_from_media)
            base_df["_prio_norm"] = base_df.get("prioridade", pd.Series(index=base_df.index, dtype=str)).astype(str).apply(self._norm_prio)
            bcol = base_df.get("bairro", pd.Series(index=base_df.index, dtype=str)).astype(str)
            base_df["_bairro_norm"] = self._canonicalize_bairro_series(bcol)

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            initial_name = f"OPERACIONAL_{stamp}.xlsx"
            out = filedialog.asksaveasfilename(
                title="Salvar exportacao operacional",
                defaultextension=".xlsx",
                initialdir=str(self.results_dir),
                initialfile=initial_name,
                filetypes=[("Excel", "*.xlsx")],
            )
            if not out:
                return
            out_path = Path(out)

            cols_export = ["nome", "_bairro_norm", "_prio_norm", "_class", "_media_num", "pendencias", "indicadores", "oqf"]
            work = base_df[cols_export].copy()
            work = work.rename(
                columns={
                    "nome": "Nome",
                    "_bairro_norm": "Bairro",
                    "_prio_norm": "Prioridade",
                    "_class": "Classificacao",
                    "_media_num": "Pontos",
                    "pendencias": "Pendencias",
                    "indicadores": "Indicadores",
                    "oqf": "Acao",
                }
            )
            work["Pontos"] = work["Pontos"].round(0).astype(int)

            urgentes = work[work["Prioridade"] == "URGENTE"].copy()
            zero_pts = work[work["Pontos"] == 0].copy()
            regulares = work[work["Classificacao"] == "Regular"].copy()

            risk_map = {"URGENTE": 3, "ALTA": 2, "MONITORAR": 1, "CONCLUIDO": 0}
            bairros = work.copy()
            bairros["_risk"] = bairros["Prioridade"].map(risk_map).fillna(0).astype(int)
            top_bairros = (
                bairros.groupby("Bairro", dropna=False)
                .agg(
                    Total=("Nome", "count"),
                    Urgentes=("Prioridade", lambda s: int((s == "URGENTE").sum())),
                    ZeroPts=("Pontos", lambda s: int((s == 0).sum())),
                    Regulares=("Classificacao", lambda s: int((s == "Regular").sum())),
                    ScoreRisco=("_risk", "sum"),
                )
                .reset_index()
                .sort_values(["ScoreRisco", "Urgentes", "ZeroPts", "Total"], ascending=[False, False, False, False])
            )

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                work.to_excel(writer, sheet_name="FilaFiltrada", index=False)
                urgentes.to_excel(writer, sheet_name="Urgentes", index=False)
                zero_pts.to_excel(writer, sheet_name="ZeroPontos", index=False)
                regulares.to_excel(writer, sheet_name="Regulares", index=False)
                top_bairros.to_excel(writer, sheet_name="TopBairros", index=False)

            self.action_status_var.set(f"Exportacao operacional concluida: {out_path.name}")
            messagebox.showinfo(
                "Exportacao operacional",
                f"Arquivo gerado:\n{out_path}\n\n"
                f"Fila filtrada: {len(work)}\nUrgentes: {len(urgentes)}\n0 pontos: {len(zero_pts)}\nRegulares: {len(regulares)}",
            )
        except Exception as exc:
            messagebox.showerror("Erro ao exportar operacional", str(exc))



    def _pdf_cell_text(self, col_name: str, value) -> str:
        txt = self._fix_text(value)
        col = str(col_name or "")
        if "Arquivo" in col:
            try:
                txt = Path(str(txt)).name
            except Exception:
                txt = str(txt)
        txt = str(txt).replace("\r", " ").replace("\n", " ")
        txt = re.sub(r"\s+", " ", txt).strip()
        max_len = 28 if "Arquivo" in col else 36
        if len(txt) > max_len:
            txt = txt[: max_len - 3] + "..."
        return txt

    def _df_to_table_data(self, df: pd.DataFrame, max_rows: int | None = None):
        if df is None or df.empty:
            return [["Sem dados"]]
        frame = df.copy()
        if max_rows:
            frame = frame.head(max_rows)
        frame = frame.fillna("")
        frame.columns = [self._fix_text(c) for c in frame.columns]
        for col in frame.columns:
            frame[col] = frame[col].map(lambda v: self._pdf_cell_text(col, v))
        rows = frame.astype(str).values.tolist()
        data = [list(frame.columns)] + rows
        return data

    def _num_from_text(self, value, default: float = 0.0) -> float:
        txt = str(value or "").replace(",", ".")
        m = re.search(r"[-+]?\d+(?:\.\d+)?", txt)
        if not m:
            return default
        try:
            return float(m.group(0))
        except Exception:
            return default

    def _fix_text(self, value) -> str:
        s = str(value if value is not None else "")
        try:
            repaired = s.encode("latin1", errors="ignore").decode("utf-8", errors="ignore")
            if repaired and (repaired.count("Ã") + repaired.count("â")) < (s.count("Ã") + s.count("â")):
                s = repaired
        except Exception:
            pass
        return s

    def _kpi_status(self, value: float, positive_good: bool = True) -> str:
        if value == 0:
            return "Estavel"
        improved = value > 0 if positive_good else value < 0
        return "Melhorou" if improved else "Piorou"

    def _make_pdf_table(self, df: pd.DataFrame, title: str, max_rows: int | None = None, small: bool = False):
        styles = getSampleStyleSheet()
        elems = []
        if df is None or df.empty:
            elems.extend([Paragraph(title, styles["Heading2"]), Spacer(1, 0.18 * cm)])
            data = self._df_to_table_data(pd.DataFrame({"Sem dados": ["-"]}), max_rows=1)
            tbl = Table(data, repeatRows=1, colWidths=[26.8 * cm])
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7C9D6")),
            ]))
            elems.append(tbl)
            elems.append(Spacer(1, 0.35 * cm))
            return elems

        frame = df.copy()
        if max_rows:
            frame = frame.head(max_rows)

        all_cols = list(frame.columns)
        fixed_cols = [c for c in all_cols if str(c) in {"Indicador", "Classificacao"}]
        if not fixed_cols and all_cols:
            fixed_cols = [all_cols[0]]
        fixed_cols = fixed_cols[:1]
        max_cols_per_part = 8
        if len(all_cols) <= max_cols_per_part:
            col_groups = [all_cols]
        else:
            rest_cols = [c for c in all_cols if c not in fixed_cols]
            chunk = max(1, max_cols_per_part - len(fixed_cols))
            col_groups = [fixed_cols + rest_cols[i:i + chunk] for i in range(0, len(rest_cols), chunk)]

        total_parts = len(col_groups)
        total_w = 26.8 * cm

        for idx, cols in enumerate(col_groups, start=1):
            part_title = title if total_parts == 1 else f"{title} (parte {idx}/{total_parts})"
            elems.extend([Paragraph(part_title, styles["Heading2"]), Spacer(1, 0.14 * cm)])
            data = self._df_to_table_data(frame[cols], max_rows=None)
            headers = [str(h) for h in data[0]]
            ncols = max(1, len(headers))

            if ncols == 1:
                widths = [total_w]
            else:
                base_widths = []
                for h in headers:
                    if "Arquivo" in h:
                        base_widths.append(4.0)
                    elif "Indicador Label" in h:
                        base_widths.append(3.8)
                    elif h in {"Indicador", "Classificacao"}:
                        base_widths.append(2.6)
                    else:
                        base_widths.append(2.0)
                scale = total_w / (sum(base_widths) * cm)
                widths = [(w * cm * scale) for w in base_widths]

            tbl = Table(data, repeatRows=1, colWidths=widths)
            font_size = 7 if small else 8
            if ncols >= 7:
                font_size = 6.5 if small else 7
            style_cmds = [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("LEADING", (0, 0), (-1, -1), font_size + 2),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7C9D6")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#EEF4F8")]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]

            for cidx, h in enumerate(headers):
                if ("Indicador" in h) or ("Classificacao" in h) or ("Arquivo" in h):
                    style_cmds.append(("ALIGN", (cidx, 1), (cidx, -1), "LEFT"))

            delta_cols = [i for i, h in enumerate(headers) if ("Variacao" in h) or ("Delta" in h) or ("Δ" in h)]
            for r in range(1, len(data)):
                for c in delta_cols:
                    val = self._num_from_text(data[r][c], default=0.0)
                    if val > 0:
                        style_cmds.append(("TEXTCOLOR", (c, r), (c, r), colors.HexColor("#1B5E20")))
                        style_cmds.append(("FONTNAME", (c, r), (c, r), "Helvetica-Bold"))
                    elif val < 0:
                        style_cmds.append(("TEXTCOLOR", (c, r), (c, r), colors.HexColor("#B71C1C")))
                        style_cmds.append(("FONTNAME", (c, r), (c, r), "Helvetica-Bold"))
                    else:
                        style_cmds.append(("TEXTCOLOR", (c, r), (c, r), colors.HexColor("#1F4E79")))

            tbl.setStyle(TableStyle(style_cmds))
            elems.append(tbl)
            elems.append(Spacer(1, 0.28 * cm))
        return elems

    def exportar_relatorio_pdf(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta para salvar o relatorio e os graficos")
        if not pasta:
            return

        pasta = Path(pasta)
        pasta.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        png1 = pasta / f"{timestamp}_grafico_panorama_principal.png"
        png2 = pasta / f"{timestamp}_grafico_busca_ativa_por_indicador.png"
        png3 = pasta / f"{timestamp}_grafico_distribuicao_classificacao.png"
        png4 = pasta / f"{timestamp}_grafico_risco_por_bairro.png"
        excel_path = pasta / f"{timestamp}_Dados_Relatorio_APS.xlsx"

        try:
            include = self.pdf_graph_flags
            manual_ready = self.manual_compare_merged_df is not None and not self.manual_compare_merged_df.empty
            if include["panorama"].get():
                self.fig1.savefig(png1, dpi=180, bbox_inches="tight")
            if include["busca"].get():
                self.fig2.savefig(png2, dpi=180, bbox_inches="tight")
            if include["classificacao"].get():
                self.fig3.savefig(png3, dpi=180, bbox_inches="tight")
            if include["risco"].get():
                self.fig4.savefig(png4, dpi=180, bbox_inches="tight")

            compare_png1 = None
            compare_png2 = None
            if include["manual_media"].get() and manual_ready and hasattr(self, "figm1"):
                compare_png1 = pasta / f"{timestamp}_grafico_comparacao_media.png"
                self.figm1.savefig(compare_png1, dpi=180, bbox_inches="tight")
            if include["manual_classificacao"].get() and manual_ready and hasattr(self, "figm2"):
                compare_png2 = pasta / f"{timestamp}_grafico_comparacao_classificacao.png"
                self.figm2.savefig(compare_png2, dpi=180, bbox_inches="tight")

            pdf_path = pasta / f"{timestamp}_Relatorio_Dashboard_APS.pdf"
            styles = getSampleStyleSheet()
            styles.add(ParagraphStyle(
                name="APSBody",
                parent=styles["BodyText"],
                fontName="Helvetica",
                fontSize=9,
                leading=12,
                spaceAfter=6,
            ))
            styles.add(ParagraphStyle(
                name="APSBodySmall",
                parent=styles["BodyText"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                spaceAfter=4,
            ))

            doc = SimpleDocTemplate(
                str(pdf_path),
                pagesize=landscape(A4),
                rightMargin=1.2 * cm,
                leftMargin=1.2 * cm,
                topMargin=1.2 * cm,
                bottomMargin=1.2 * cm,
            )

            story = []
            delta_media_val = self._num_from_text(self.card_vars["delta"].get(), default=0.0)
            delta_busca_val = self._num_from_text(self.card_vars["delta_busca"].get(), default=0.0)
            status_media = self._kpi_status(delta_media_val, positive_good=True)
            status_busca = self._kpi_status(delta_busca_val, positive_good=False)

            story.append(Paragraph(self._fix_text("RELATORIO EXECUTIVO - DASHBOARD APS"), styles["Title"]))
            story.append(Paragraph(
                self._fix_text(f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')} | Pasta de resultados: {self.results_dir}"),
                styles["APSBody"]
            ))
            story.append(Spacer(1, 0.25 * cm))

            resumo_text = (
                f"<b>Leitura rapida:</b> Media geral: <b>{self.card_vars['media'].get()}</b> "
                f"(variacao {delta_media_val:+.1f} - {status_media}). "
                f"Busca ativa: <b>{self.card_vars['busca'].get()}</b> "
                f"(variacao {int(delta_busca_val):+d} - {status_busca}). "
                f"Critico operacional (0 pts): <b>{self.card_vars['critico_zero'].get()}</b>."
            )
            story.append(Paragraph(resumo_text, styles["APSBody"]))
            story.append(Spacer(1, 0.15 * cm))

            cards_df = pd.DataFrame([{
                "Arquivos atuais": self.card_vars["arquivos"].get(),
                "Pacientes brutos": self.card_vars["linhas_brutas"].get(),
                "Pacientes unicos": self.card_vars["linhas"].get(),
                "Busca ativa atual": self.card_vars["busca"].get(),
                "Critico (0 pts)": self.card_vars["critico_zero"].get(),
                "Media atual": self.card_vars["media"].get(),
                "Variacao media vs anterior": f"{delta_media_val:+.1f}",
                "Variacao busca ativa": f"{int(delta_busca_val):+d}",
            }])

            folder_cmp_df = pd.DataFrame()
            if self.folder_compare_excel_path is not None and Path(self.folder_compare_excel_path).exists():
                try:
                    folder_cmp_df = pd.read_excel(Path(self.folder_compare_excel_path), sheet_name=0, header=2)
                    folder_cmp_df = folder_cmp_df.dropna(how="all").reset_index(drop=True)
                except Exception:
                    folder_cmp_df = pd.DataFrame()
            if folder_cmp_df.empty and self.folder_compare_df is not None and not self.folder_compare_df.empty:
                folder_cmp_df = self.folder_compare_df.copy()

            manual_df_export = pd.DataFrame()
            manual_tree_export = pd.DataFrame()
            if manual_ready and self.path_a.get().strip() and self.path_b.get().strip():
                manual_df_export = pd.DataFrame([{
                    "Arquivo A": self.summary_a.get("arquivo", "-"),
                    "Total A": self.summary_a.get("total", "-"),
                    "Media A": self.summary_a.get("media", "-"),
                    "Busca ativa A": self.summary_a.get("busca", "-"),
                    "Arquivo B": self.summary_b.get("arquivo", "-"),
                    "Total B": self.summary_b.get("total", "-"),
                    "Media B": self.summary_b.get("media", "-"),
                    "Busca ativa B": self.summary_b.get("busca", "-"),
                }])
                rows_export = []
                for item in self.tree_manual.get_children():
                    rows_export.append(self.tree_manual.item(item, "values"))
                if rows_export:
                    manual_tree_export = pd.DataFrame(rows_export, columns=["Classificacao", "Arquivo A", "Arquivo B", "Variacao", "Variacao %"])

            selected_graphs_df = pd.DataFrame([{
                "Panorama principal": "SIM" if include["panorama"].get() else "NAO",
                "Busca ativa": "SIM" if include["busca"].get() else "NAO",
                "Distribuicao classificacao": "SIM" if include["classificacao"].get() else "NAO",
                "Risco por bairro": "SIM" if include["risco"].get() else "NAO",
                "Comparar arquivos - media": "SIM" if (include["manual_media"].get() and manual_ready) else "NAO",
                "Comparar arquivos - classificacao": "SIM" if (include["manual_classificacao"].get() and manual_ready) else "NAO",
            }])

            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                cards_df.to_excel(writer, sheet_name="Resumo_Executivo", index=False)
                selected_graphs_df.to_excel(writer, sheet_name="Graficos_Selecionados", index=False)
                if self.summary_df is not None and not self.summary_df.empty:
                    self.summary_df.to_excel(writer, sheet_name="Resumo_Atual", index=False)
                if self.compare_df is not None and not self.compare_df.empty:
                    self.compare_df.to_excel(writer, sheet_name="Comparativo_Atual_Anterior", index=False)
                if not folder_cmp_df.empty:
                    folder_cmp_df.to_excel(writer, sheet_name="Comparacao_Pastas", index=False)
                if not self.folder_summary_a.empty:
                    self.folder_summary_a.to_excel(writer, sheet_name="Pastas_Periodo_A", index=False)
                if not self.folder_summary_b.empty:
                    self.folder_summary_b.to_excel(writer, sheet_name="Pastas_Periodo_B", index=False)
                if not manual_df_export.empty:
                    manual_df_export.to_excel(writer, sheet_name="Comparacao_Manual_Resumo", index=False)
                if not manual_tree_export.empty:
                    manual_tree_export.to_excel(writer, sheet_name="Comparacao_Manual_Classe", index=False)

            story.extend(self._make_pdf_table(cards_df, "Resumo executivo", max_rows=1, small=False))
            story.append(Paragraph(self._fix_text(f"Base de dados do relatorio (Excel): {excel_path}"), styles["APSBodySmall"]))

            if self.summary_df is not None and not self.summary_df.empty:
                story.extend(self._make_pdf_table(self.summary_df, "Tabela 1 - Resumo atual por indicador", max_rows=30, small=True))

            if self.compare_df is not None and not self.compare_df.empty:
                story.extend(self._make_pdf_table(self.compare_df, "Tabela 2 - Comparativo atual x anterior", max_rows=30, small=True))
            if not folder_cmp_df.empty:
                src_text = (
                    f"Arquivo base da comparacao de pastas: {self.folder_compare_excel_path}"
                    if self.folder_compare_excel_path is not None
                    else "Arquivo base da comparacao de pastas: dados em memoria"
                )
                story.append(Paragraph(self._fix_text(src_text), styles["APSBodySmall"]))
                story.extend(self._make_pdf_table(folder_cmp_df, "Tabela 3 - Comparacao entre pastas", max_rows=30, small=True))

            img_list = []
            if include["panorama"].get() and png1.exists():
                img_list.append((png1, "Grafico - Panorama principal"))
            if include["busca"].get() and png2.exists():
                img_list.append((png2, "Grafico - Busca ativa atual por indicador"))
            if include["classificacao"].get() and png3.exists():
                img_list.append((png3, "Grafico - Distribuicao atual por classificacao"))
            if include["risco"].get() and png4.exists():
                img_list.append((png4, "Grafico - Risco por bairro"))
            if img_list:
                story.append(PageBreak())
                for idx_img, (img_path, title) in enumerate(img_list, start=1):
                    story.append(Paragraph(self._fix_text(title), styles["Heading2"]))
                    story.append(Spacer(1, 0.12 * cm))
                    story.append(Image(str(img_path), width=24.5 * cm, height=9.2 * cm))
                    story.append(Spacer(1, 0.28 * cm))
                    if idx_img % 2 == 0 and idx_img < len(img_list):
                        story.append(PageBreak())

            if manual_ready and self.path_a.get().strip() and self.path_b.get().strip():
                story.append(PageBreak())
                story.append(Paragraph("Comparacao manual entre arquivos", styles["Heading1"]))
                story.append(Paragraph(
                    "Sessao para comparar dois arquivos especificos (A x B), com foco em media, busca ativa e distribuicao por classificacao.",
                    styles["APSBodySmall"],
                ))

                story.extend(self._make_pdf_table(manual_df_export, "Tabela 4 - Resumo da comparacao manual", max_rows=1, small=True))
                if not manual_tree_export.empty:
                    story.extend(self._make_pdf_table(manual_tree_export, "Tabela 5 - Distribuicao por classificacao", max_rows=20, small=True))

                if compare_png1:
                    story.append(Paragraph("Grafico 5 - Comparacao de pontuacao media", styles["Heading2"]))
                    story.append(Image(str(compare_png1), width=24.5 * cm, height=8.8 * cm))
                    story.append(Spacer(1, 0.25 * cm))
                if compare_png2:
                    story.append(Paragraph("Grafico 6 - Comparacao por classificacao", styles["Heading2"]))
                    story.append(Image(str(compare_png2), width=24.5 * cm, height=8.8 * cm))

            doc.build(story)
            messagebox.showinfo(
                "Exportacao concluida",
                f"Arquivos gerados com sucesso em:\n{pasta}\n\nPDF: {pdf_path.name}\nExcel de dados: {excel_path.name}",
            )
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))
    def _build_folders_tab(self, root):
        tk.Label(root,
                 text="Compare os resultados de dois periodos diferentes (ex: mes anterior x mes atual).",
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", padx=8, pady=(10, 4))

        frm = ttk.LabelFrame(root, text="Selecionar pastas")
        frm.pack(fill="x", padx=8, pady=6)
        frm.columnconfigure(1, weight=1)

        for row_i, (lbl_var, name_var, pick_cmd, label) in enumerate([
            (self.folder_a_var, self.folder_a_name, self._pick_folder_a, "Periodo A (anterior):"),
            (self.folder_b_var, self.folder_b_name, self._pick_folder_b, "Periodo B (atual):"),
        ]):
            tk.Label(frm, text=label, font=("Segoe UI", 9, "bold")).grid(
                row=row_i*2, column=0, sticky="w", padx=8, pady=(8, 0))
            ttk.Entry(frm, textvariable=lbl_var, state="readonly").grid(
                row=row_i*2, column=1, sticky="ew", padx=6)
            tk.Button(frm, text="Escolher pasta", command=pick_cmd,
                      bg="#2E75B6", fg="white", font=("Segoe UI", 8)).grid(
                row=row_i*2, column=2, padx=4)
            tk.Label(frm, text="  Rotulo:").grid(row=row_i*2+1, column=0, sticky="e", padx=8, pady=(0, 6))
            ttk.Entry(frm, textvariable=name_var, width=22).grid(
                row=row_i*2+1, column=1, sticky="w", padx=6, pady=(0, 6))

        info = ttk.LabelFrame(root, text="O que sera comparado")
        info.pack(fill="x", padx=8, pady=4)
        tk.Label(info,
                 text=("- Le automaticamente os arquivos C1-C7 mais recentes de cada pasta\n"
                       "- Compara total de pacientes, pontuacao media, busca ativa e classificacoes\n"
                       "- Variacao verde = melhora  |  Variacao vermelha = piora  |  Exporta planilha com cores"),
                 font=("Segoe UI", 9), justify="left").pack(padx=10, pady=8, anchor="w")

        bot = tk.Frame(root)
        bot.pack(fill="x", padx=8, pady=8)
        tk.Label(bot, textvariable=self.folders_status, font=("Segoe UI", 9), anchor="w").pack(
            side="left", fill="x", expand=True)
        tk.Button(bot, text=">  Gerar comparacao entre pastas", command=self._run_folders,
                  bg="#7030A0", fg="white", font=("Segoe UI", 10, "bold")).pack(side="right")

    def _pick_folder_a(self):
        f = filedialog.askdirectory(title="Pasta Periodo A", initialdir=str(self.results_dir))
        if f: self.folder_a_var.set(f)

    def _pick_folder_b(self):
        f = filedialog.askdirectory(title="Pasta Periodo B", initialdir=str(self.results_dir))
        if f: self.folder_b_var.set(f)

    def _export_folder_comparison_graph(
        self,
        df: pd.DataFrame,
        out_png: Path,
        label_a: str,
        label_b: str,
        summary_a: pd.DataFrame | None = None,
        summary_b: pd.DataFrame | None = None,
    ) -> Path | None:
        if (df is None or df.empty) and (
            (summary_a is None or summary_a.empty) and (summary_b is None or summary_b.empty)
        ):
            return None
        if summary_a is None:
            summary_a = pd.DataFrame()
        if summary_b is None:
            summary_b = pd.DataFrame()
        if summary_a.empty and summary_b.empty:
            # Fallback: se nao vierem os summaries, tenta gerar por indicador a partir da pasta atual.
            # Nao interrompe a exportacao; apenas produz o grafico vazio com titulo.
            summary_a = pd.DataFrame(columns=["Indicador"])
            summary_b = pd.DataFrame(columns=["Indicador"])

        fig = Figure(figsize=(14, 6.2), dpi=150)
        ax = fig.add_subplot(111)
        ok = self._stacked_folder_compare_chart(ax, summary_a, summary_b, label_a, label_b)
        if not ok:
            return None
        fig.tight_layout()
        fig.savefig(out_png, dpi=180, bbox_inches="tight")
        return out_png

    def _run_folders(self):
        fa = self.folder_a_var.get().strip()
        fb = self.folder_b_var.get().strip()
        if fa in ("(nao selecionada)", "") or fb in ("(nao selecionada)", ""):
            messagebox.showwarning("Atencao", "Selecione as duas pastas.")
            return
        self.folders_status.set("Processando..."); self.update()
        try:
            from aps_comparador_paciente import build_folder_comparison, export_folder_comparison_excel
            la = self.folder_a_name.get().strip() or "Periodo A"
            lb = self.folder_b_name.get().strip() or "Periodo B"
            df = build_folder_comparison(Path(fa), Path(fb), la, lb)
            if df.empty:
                messagebox.showwarning("Resultado vazio", "Nenhum arquivo C1-C7 encontrado nas pastas.")
                return
            # Leva a comparacao de pastas para o Panorama (torres por C anterior x atual).
            self.folder_summary_a = build_current_summary(Path(fa))
            self.folder_summary_b = build_current_summary(Path(fb))
            self.folder_compare_df = df.copy()
            self.folder_compare_label_a = la
            self.folder_compare_label_b = lb
            self.folder_class_totals_a = self._load_folder_class_totals(Path(fa))
            self.folder_class_totals_b = self._load_folder_class_totals(Path(fb))
            if hasattr(self, "summary_df") and self.summary_df is not None and not self.summary_df.empty:
                self._refresh_overview_charts()
                try:
                    self.notebook.select(self.tab_overview)
                except Exception:
                    pass
            import datetime as dt
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            out = self.results_dir / f"COMPARACAO_PASTAS_{stamp}.xlsx"
            export_folder_comparison_excel(df, out, la, lb)
            self.folder_compare_excel_path = out
            out_png = self.results_dir / f"COMPARACAO_PASTAS_{stamp}_grafico.png"
            gpath = self._export_folder_comparison_graph(
                df,
                out_png,
                la,
                lb,
                summary_a=self.folder_summary_a,
                summary_b=self.folder_summary_b,
            )
            self.folders_status.set(f"OK {out.name}")
            msg = f"Comparacao gerada:\n\n{out}"
            if gpath:
                msg += f"\n{gpath}"
            messagebox.showinfo("Concluido OK", msg)
            try:
                import os; os.startfile(out)
            except Exception:
                pass
        except Exception as exc:
            import traceback
            self.folders_status.set(f"Erro: {exc}")
            messagebox.showerror("Erro", f"{exc}\n\n{traceback.format_exc()}")

    def show_history_window(self):
        code = self.indicator_var.get()
        hist = build_history(self.results_dir, code)
        if hist.empty:
            messagebox.showinfo("Historico", f"Sem historico para {code}.")
            return
        win = tk.Toplevel(self)
        win.title(f"Historico - {code}")
        win.geometry("960x520")
        cols = tuple(hist.columns)
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=110, anchor="center")
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        for _, row in hist.iterrows():
            tree.insert("", "end", values=tuple(row[c] for c in cols))


def launch_dashboard(results_dir: Path | None = None):
    results_dir = Path(results_dir or (Path.home() / "Desktop" / "APS_RESULTADOS"))
    if not results_dir.exists():
        raise FileNotFoundError(f"Pasta de resultados nao encontrada: {results_dir}")

    root = tk._default_root
    if root is None:
        root = tk.Tk()
        root.withdraw()
    win = APSDashboard(root, results_dir)
    return win


def main():
    results_dir = None
    for arg in sys.argv[1:]:
        low = str(arg).strip().lower()
        if not str(arg).startswith("--"):
            results_dir = Path(arg)
    root = tk.Tk()
    root.withdraw()
    try:
        launch_dashboard(results_dir)
        root.mainloop()
    except Exception as exc:
        messagebox.showerror("Erro", str(exc))


if __name__ == "__main__":
    main()



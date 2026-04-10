from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
import tempfile
from pathlib import Path
import re
import sys
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


CLASS_ORDER = ["Ótimo", "Bom", "Suficiente", "Regular", "Ruim", "Crítico"]
PRIORITY_ORDER = ["Concluído", "Baixa", "Média", "Alta", "🔴 URGENTE", "🟠 ALTA", "🟡 MÉDIA", "🟢 BAIXA"]


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


def indicator_files(results_dir: Path) -> dict[str, list[Path]]:
    grouped: dict[str, list[Path]] = {}
    for f in results_dir.glob("C*.xlsx"):
        m = re.match(r"^(C\d+)_(.+)\.xlsx$", f.name, flags=re.I)
        if not m:
            continue
        code = m.group(1).upper()
        grouped.setdefault(code, []).append(f)
    for code in grouped:
        grouped[code].sort(key=lambda p: p.stat().st_mtime)
    return dict(sorted(grouped.items()))


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
    xls = pd.ExcelFile(xlsx_path)
    sheet = next((s for s in xls.sheet_names if str(s).startswith("📋 Dados")), xls.sheet_names[0])
    for header_try in (2, 1, 0):
        try:
            df = pd.read_excel(xlsx_path, sheet_name=sheet, header=header_try)
            df = df.dropna(how="all")
            if len(df.columns) > 1:
                if "Nome" in df.columns:
                    df = df[df["Nome"].astype(str).str.strip().ne("")]
                return df.reset_index(drop=True)
        except Exception:
            continue
    return pd.read_excel(xlsx_path, sheet_name=sheet).dropna(how="all").reset_index(drop=True)


def build_snapshot(code: str, path: Path) -> Snapshot:
    df = read_indicator_dataframe(path)
    total = len(df)
    pontos = pd.to_numeric(df.get("Pontuação"), errors="coerce").fillna(0)
    classes = Counter(df.get("Classificação", pd.Series(dtype=str)).fillna("Sem classificação").astype(str))
    prioridades = Counter(df.get("Prioridade", pd.Series(dtype=str)).fillna("Sem prioridade").astype(str))
    busca_ativa = int((pontos < 100).sum()) if total else 0
    return Snapshot(
        indicador=code,
        arquivo=path,
        momento=parse_stamp(path),
        total=total,
        busca_ativa=busca_ativa,
        media_pontuacao=round(float(pontos.mean()) if total else 0.0, 1),
        classes=dict(classes),
        prioridades=dict(prioridades),
    )


def build_current_summary(results_dir: Path) -> pd.DataFrame:
    rows = []
    for code, files in indicator_files(results_dir).items():
        snap = build_snapshot(code, files[-1])
        row = {
            "Indicador": code,
            "Arquivo": snap.arquivo.name,
            "Total": snap.total,
            "Busca Ativa": snap.busca_ativa,
            "Média Pontuação": snap.media_pontuacao,
        }
        for cls in CLASS_ORDER:
            row[cls] = snap.classes.get(cls, 0)
        rows.append(row)
    return pd.DataFrame(rows)


def build_comparison_summary(results_dir: Path) -> pd.DataFrame:
    rows = []
    for code, files in indicator_files(results_dir).items():
        atual = build_snapshot(code, files[-1])
        anterior = build_snapshot(code, files[-2]) if len(files) >= 2 else None
        row = {
            "Indicador": code,
            "Arquivo Atual": atual.arquivo.name,
            "Atual": atual.media_pontuacao,
            "Anterior": anterior.media_pontuacao if anterior else None,
            "Δ Média": round(atual.media_pontuacao - anterior.media_pontuacao, 1) if anterior else None,
            "Δ Total": atual.total - anterior.total if anterior else None,
            "Δ Busca Ativa": atual.busca_ativa - anterior.busca_ativa if anterior else None,
        }
        for cls in CLASS_ORDER:
            row[f"Δ {cls}"] = atual.classes.get(cls, 0) - (anterior.classes.get(cls, 0) if anterior else 0)
        rows.append(row)
    return pd.DataFrame(rows)


def build_history(results_dir: Path, indicador: str) -> pd.DataFrame:
    files = indicator_files(results_dir).get(indicador, [])
    rows = []
    for path in files:
        snap = build_snapshot(indicador, path)
        label = snap.momento.strftime("%d/%m %H:%M") if snap.momento else path.stem
        row = {
            "Momento": label,
            "Pontuação Média": snap.media_pontuacao,
            "Total": snap.total,
            "Busca Ativa": snap.busca_ativa,
        }
        for cls in CLASS_ORDER:
            row[cls] = snap.classes.get(cls, 0)
        rows.append(row)
    return pd.DataFrame(rows)


def build_manual_summary(path: Path) -> tuple[pd.DataFrame, dict]:
    df = read_indicator_dataframe(path)
    pts = pd.to_numeric(df.get("Pontuação"), errors="coerce").fillna(0)
    classes = df.get("Classificação", pd.Series(dtype=str)).fillna("Sem classificação").astype(str)
    counts = {c: int((classes == c).sum()) for c in CLASS_ORDER}
    summary = {
        "arquivo": path.name,
        "total": int(len(df)),
        "media": round(float(pts.mean()) if len(df) else 0.0, 1),
        "busca": int((pts < 100).sum()) if len(df) else 0,
        **counts,
    }
    table = pd.DataFrame({
        "Classificação": CLASS_ORDER,
        "Quantidade": [counts[c] for c in CLASS_ORDER],
    })
    return table, summary


class APSDashboard(tk.Toplevel):
    def __init__(self, parent: tk.Misc | None, results_dir: Path):
        super().__init__(parent)
        self.results_dir = Path(results_dir)
        self.title("APS - Dashboard Evolutivo + Comparador")
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
            "linhas":   tk.StringVar(value="0"),
            "busca":    tk.StringVar(value="0"),
            "media":    tk.StringVar(value="0.0"),
            "delta":    tk.StringVar(value="0.0"),
            "delta_busca": tk.StringVar(value="0"),
        }
        self.path_a = tk.StringVar()
        self.path_b = tk.StringVar()
        self.status_compare = tk.StringVar(value="Selecione dois arquivos XLSX para comparar.")
        self.summary_a: dict = {}
        self.summary_b: dict = {}
        self.folder_a_var = tk.StringVar(value="(não selecionada)")
        self.folder_b_var = tk.StringVar(value="(não selecionada)")
        self.folder_a_name = tk.StringVar(value="Período A")
        self.folder_b_name = tk.StringVar(value="Período B")
        self.folders_status = tk.StringVar(value="Selecione as duas pastas para comparar.")
        self._build_ui()
        self.refresh()
        # Janela sempre à frente
        self.lift()
        self.focus_force()

    def _build_ui(self):
        header = tk.Label(self, text="Dashboard APS - Atual x Anterior + Comparador Manual", bg="#1F4E79", fg="white", font=("Segoe UI", 16, "bold"), pady=12)
        header.pack(fill="x")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=12)

        self.tab_overview = ScrollableTab(self.notebook)
        self.tab_compare = ScrollableTab(self.notebook)
        self.tab_folders = ScrollableTab(self.notebook)
        self.notebook.add(self.tab_overview, text="Panorama")
        self.notebook.add(self.tab_compare, text="Comparar arquivos")
        self.notebook.add(self.tab_folders, text="📁 Comparar pastas")

        self._build_overview_tab(self.tab_overview.inner)
        self._build_compare_tab(self.tab_compare.inner)
        self._build_folders_tab(self.tab_folders.inner)

    def _build_overview_tab(self, root):
        top = ttk.Frame(root)
        top.pack(fill="x", padx=4, pady=10)
        ttk.Button(top, text="Atualizar", command=self.refresh).pack(side="left")
        ttk.Button(top, text="Exportar relatório PDF", command=self.exportar_relatorio_pdf).pack(side="left", padx=(8,0))

        # ── Seletor de pasta ──
        ttk.Label(top, text="  Pasta:").pack(side="left", padx=(16,4))
        ttk.Entry(top, textvariable=self.folder_var, width=38).pack(side="left")
        ttk.Button(top, text="Trocar pasta", command=self._change_folder).pack(side="left", padx=(4,16))

        ttk.Label(top, text="Histórico:").pack(side="left", padx=(0,6))
        self.cbo_indicator = ttk.Combobox(top, textvariable=self.indicator_var, state="readonly", width=8)
        self.cbo_indicator.pack(side="left")
        ttk.Button(top, text="Ver histórico", command=self.show_history_window).pack(side="left", padx=(8,0))

        cards = ttk.Frame(root)
        cards.pack(fill="x", padx=4)
        self._make_card(cards, "Arquivos atuais", self.card_vars["arquivos"], 0)
        self._make_card(cards, "Pacientes atuais", self.card_vars["linhas"], 1)
        self._make_card(cards, "Busca ativa atual", self.card_vars["busca"], 2)
        self._make_card(cards, "Média atual", self.card_vars["media"], 3)
        self._make_card(cards, "Δ média vs anterior", self.card_vars["delta"], 4)
        self._make_card(cards, "Δ busca ativa", self.card_vars["delta_busca"], 5)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True, padx=4, pady=12)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)
        body.rowconfigure(3, weight=1)

        current_box = ttk.LabelFrame(body, text="Resumo atual por indicador")
        current_box.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        current_box.rowconfigure(0, weight=1)
        current_box.columnconfigure(0, weight=1)
        cur_cols = ("Indicador", "Total", "Busca Ativa", "Média Pontuação", "Ótimo", "Bom", "Suficiente", "Regular", "Ruim", "Crítico")
        self.tree_current = ttk.Treeview(current_box, columns=cur_cols, show="headings", height=6)
        for c in cur_cols:
            self.tree_current.heading(c, text=c)
            self.tree_current.column(c, width=100, anchor="center")
        self.tree_current.column("Indicador", width=90)
        self.tree_current.grid(row=0, column=0, sticky="nsew")
        sc1 = ttk.Scrollbar(current_box, orient="vertical", command=self.tree_current.yview)
        sc1.grid(row=0, column=1, sticky="ns")
        self.tree_current.configure(yscrollcommand=sc1.set)

        compare_box = ttk.LabelFrame(body, text="Comparativo atual x anterior")
        compare_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        compare_box.rowconfigure(0, weight=1)
        compare_box.columnconfigure(0, weight=1)
        cmp_cols = ("Indicador", "Atual", "Anterior", "Δ Média", "Δ Total", "Δ Busca Ativa", "Δ Ótimo", "Δ Bom", "Δ Suficiente", "Δ Regular", "Δ Ruim", "Δ Crítico")
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
        self.canvas1.get_tk_widget().grid(row=2, column=0, sticky="nsew", padx=(0, 6), pady=(0, 10))

        self.fig2 = Figure(figsize=(6, 3.2), dpi=100)
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, master=body)
        self.canvas2.get_tk_widget().grid(row=2, column=1, sticky="nsew", padx=(6, 0), pady=(0, 10))

        self.fig3 = Figure(figsize=(6, 3.4), dpi=100)
        self.ax3 = self.fig3.add_subplot(111)
        self.canvas3 = FigureCanvasTkAgg(self.fig3, master=body)
        self.canvas3.get_tk_widget().grid(row=3, column=0, sticky="nsew", padx=(0, 6))

        self.fig4 = Figure(figsize=(6, 3.4), dpi=100)
        self.ax4 = self.fig4.add_subplot(111)
        self.canvas4 = FigureCanvasTkAgg(self.fig4, master=body)
        self.canvas4.get_tk_widget().grid(row=3, column=1, sticky="nsew", padx=(6, 0))

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
        top.columnconfigure(1, weight=1)

        cards = ttk.Frame(root)
        cards.pack(fill="x", padx=4)
        self.card_a_total = self._make_card(cards, "A - Total", tk.StringVar(value='-'), 0)
        self.card_b_total = self._make_card(cards, "B - Total", tk.StringVar(value='-'), 1)
        self.card_a_media = self._make_card(cards, "A - Média", tk.StringVar(value='-'), 2)
        self.card_b_media = self._make_card(cards, "B - Média", tk.StringVar(value='-'), 3)
        self.card_delta_manual = self._make_card(cards, "Δ Média", tk.StringVar(value='-'), 4)
        self.card_busca_manual = self._make_card(cards, "Δ Busca Ativa", tk.StringVar(value='-'), 5)

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True, padx=4, pady=12)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        box_table = ttk.LabelFrame(body, text="Comparação por classificação")
        box_table.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 8))
        box_table.rowconfigure(0, weight=1)
        box_table.columnconfigure(0, weight=1)
        cols = ("Classificação", "Arquivo A", "Arquivo B", "Δ")
        self.tree_manual = ttk.Treeview(box_table, columns=cols, show="headings")
        for c in cols:
            self.tree_manual.heading(c, text=c)
            self.tree_manual.column(c, width=120, anchor="center")
        self.tree_manual.column("Classificação", width=160)
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

        ttk.Label(root, textvariable=self.status_compare, anchor="w", relief="sunken").pack(fill="x", padx=4, pady=(0, 8))

    def _make_card(self, parent, title, var, col):
        frame = tk.Frame(parent, bg="#DCE6F1", bd=1, relief="solid")
        frame.grid(row=0, column=col, sticky="ew", padx=6, pady=6)
        parent.columnconfigure(col, weight=1)
        tk.Label(frame, text=title, bg="#DCE6F1", fg="#1F1F1F", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(frame, textvariable=var, bg="#DCE6F1", fg="#1F4E79", font=("Segoe UI", 18, "bold")).pack(anchor="w", padx=12, pady=(0, 10))
        return var

    def _change_folder(self):
        folder = filedialog.askdirectory(
            title="Selecione a pasta com os resultados APS",
            initialdir=str(self.results_dir))
        if folder:
            self.results_dir = Path(folder)
            self.folder_var.set(folder)
            self.refresh()

    def refresh(self):
        files_grouped = indicator_files(self.results_dir)
        codes = list(files_grouped.keys())
        self.cbo_indicator["values"] = codes or ["C1"]
        if self.indicator_var.get() not in codes and codes:
            self.indicator_var.set(codes[0])

        self.summary_df = build_current_summary(self.results_dir)
        self.compare_df = build_comparison_summary(self.results_dir)

        for tree in (self.tree_current, self.tree_compare):
            for item in tree.get_children():
                tree.delete(item)

        if self.summary_df.empty:
            for key, value in [("arquivos", "0"), ("linhas", "0"), ("busca", "0"), ("media", "0.0"), ("delta", "0.0"), ("delta_busca", "0")]:
                self.card_vars[key].set(value)
            for ax, canvas, title in [
                (self.ax1, self.canvas1, "Sem dados"),
                (self.ax2, self.canvas2, "Sem dados"),
                (self.ax3, self.canvas3, "Sem dados"),
                (self.ax4, self.canvas4, "Sem dados"),
            ]:
                ax.clear(); ax.set_title(title); canvas.draw()
            return

        for _, row in self.summary_df.iterrows():
            self.tree_current.insert("", "end", values=(
                row.get("Indicador"), row.get("Total"), row.get("Busca Ativa"), row.get("Média Pontuação"),
                row.get("Ótimo", 0), row.get("Bom", 0), row.get("Suficiente", 0), row.get("Regular", 0), row.get("Ruim", 0), row.get("Crítico", 0)
            ))

        if not self.compare_df.empty:
            for _, row in self.compare_df.iterrows():
                vals = [row.get(c) for c in ["Indicador", "Atual", "Anterior", "Δ Média", "Δ Total", "Δ Busca Ativa", "Δ Ótimo", "Δ Bom", "Δ Suficiente", "Δ Regular", "Δ Ruim", "Δ Crítico"]]
                self.tree_compare.insert("", "end", values=tuple("-" if pd.isna(v) else v for v in vals))

        self.card_vars["arquivos"].set(str(len(self.summary_df)))
        self.card_vars["linhas"].set(str(int(self.summary_df["Total"].sum())))
        self.card_vars["busca"].set(str(int(self.summary_df["Busca Ativa"].sum())))
        self.card_vars["media"].set(f"{self.summary_df['Média Pontuação'].mean():.1f}")

        valid_delta = pd.to_numeric(self.compare_df.get("Δ Média"), errors="coerce") if not self.compare_df.empty else pd.Series(dtype=float)
        valid_busca = pd.to_numeric(self.compare_df.get("Δ Busca Ativa"), errors="coerce") if not self.compare_df.empty else pd.Series(dtype=float)
        delta_mean = float(valid_delta.dropna().mean()) if not valid_delta.dropna().empty else 0.0
        delta_busca = int(round(float(valid_busca.dropna().sum()))) if not valid_busca.dropna().empty else 0
        self.card_vars["delta"].set(f"{delta_mean:+.1f}")
        self.card_vars["delta_busca"].set(f"{delta_busca:+d}")

        self._refresh_overview_charts()

    def _refresh_overview_charts(self):
        self.ax1.clear()
        self.ax1.bar(self.summary_df["Indicador"], self.summary_df["Média Pontuação"])
        self.ax1.set_title("Pontuação média atual por indicador")
        self.ax1.set_xlabel("Indicador")
        self.ax1.set_ylabel("Média")
        self.fig1.tight_layout()
        self.canvas1.draw()

        self.ax2.clear()
        self.ax2.bar(self.summary_df["Indicador"], self.summary_df["Busca Ativa"])
        self.ax2.set_title("Busca ativa atual por indicador")
        self.ax2.set_xlabel("Indicador")
        self.ax2.set_ylabel("Pendentes")
        self.fig2.tight_layout()
        self.canvas2.draw()

        self.ax3.clear()
        stacked_cols = [c for c in CLASS_ORDER if c in self.summary_df.columns]
        bottom = None
        for cls in stacked_cols:
            vals = self.summary_df[cls]
            self.ax3.bar(self.summary_df["Indicador"], vals, bottom=bottom, label=cls)
            bottom = vals if bottom is None else bottom + vals
        self.ax3.set_title("Distribuição atual por classificação")
        self.ax3.set_xlabel("Indicador")
        self.ax3.set_ylabel("Pacientes")
        self.ax3.legend(fontsize=8)
        self.fig3.tight_layout()
        self.canvas3.draw()

        self.ax4.clear()
        if not self.compare_df.empty:
            x = self.compare_df["Indicador"]
            self.ax4.axhline(0, linewidth=1)
            self.ax4.bar(x, self.compare_df["Δ Média"].fillna(0))
            self.ax4.set_title("Variação da média: atual x anterior")
            self.ax4.set_xlabel("Indicador")
            self.ax4.set_ylabel("Δ média")
        else:
            self.ax4.set_title("Sem comparativo ainda")
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

    def compare_manual(self):
        try:
            path_a = Path(self.path_a.get())
            path_b = Path(self.path_b.get())
            if not path_a.exists() or not path_b.exists():
                raise FileNotFoundError("Selecione dois arquivos XLSX válidos.")
            table_a, self.summary_a = build_manual_summary(path_a)
            table_b, self.summary_b = build_manual_summary(path_b)

            merged = table_a.merge(table_b, on="Classificação", how="outer", suffixes=(" A", " B")).fillna(0)
            merged["Arquivo A"] = merged["Quantidade A"].astype(int)
            merged["Arquivo B"] = merged["Quantidade B"].astype(int)
            merged["Δ"] = merged["Arquivo B"] - merged["Arquivo A"]
            merged = merged[["Classificação", "Arquivo A", "Arquivo B", "Δ"]]

            for item in self.tree_manual.get_children():
                self.tree_manual.delete(item)
            for _, row in merged.iterrows():
                self.tree_manual.insert("", "end", values=list(row))

            self.card_a_total.set(str(self.summary_a["total"]))
            self.card_b_total.set(str(self.summary_b["total"]))
            self.card_a_media.set(f"{self.summary_a['media']:.1f}")
            self.card_b_media.set(f"{self.summary_b['media']:.1f}")
            self.card_delta_manual.set(f"{self.summary_b['media'] - self.summary_a['media']:+.1f}")
            self.card_busca_manual.set(str(self.summary_b['busca'] - self.summary_a['busca']))

            self.axm1.clear()
            self.axm1.bar(["A", "B"], [self.summary_a["media"], self.summary_b["media"]])
            self.axm1.set_title("Pontuação média")
            self.axm1.set_ylabel("Pontos")
            self.canvasm1.draw_idle()

            self.axm2.clear()
            x = range(len(merged))
            self.axm2.bar([i - 0.2 for i in x], merged["Arquivo A"], width=0.4, label="Arquivo A")
            self.axm2.bar([i + 0.2 for i in x], merged["Arquivo B"], width=0.4, label="Arquivo B")
            self.axm2.set_xticks(list(x))
            self.axm2.set_xticklabels(merged["Classificação"].tolist(), rotation=0)
            self.axm2.set_title("Distribuição por classificação")
            self.axm2.legend()
            self.canvasm2.draw_idle()

            self.status_compare.set("Comparação concluída com sucesso.")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            self.status_compare.set(f"Erro: {exc}")



    def _df_to_table_data(self, df: pd.DataFrame, max_rows: int | None = None):
        if df is None or df.empty:
            return [["Sem dados"]]
        frame = df.copy()
        if max_rows:
            frame = frame.head(max_rows)
        frame = frame.fillna("")
        data = [list(frame.columns)] + frame.astype(str).values.tolist()
        return data

    def _make_pdf_table(self, df: pd.DataFrame, title: str, max_rows: int | None = None, small: bool = False):
        styles = getSampleStyleSheet()
        elems = [Paragraph(title, styles["Heading2"]), Spacer(1, 0.18 * cm)]
        data = self._df_to_table_data(df, max_rows=max_rows)
        tbl = Table(data, repeatRows=1)
        font_size = 7 if small else 8
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), font_size),
            ("LEADING", (0,0), (-1,-1), font_size + 2),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#B7C9D6")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor("#EEF4F8")]),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        elems.append(tbl)
        elems.append(Spacer(1, 0.35 * cm))
        return elems

    def exportar_relatorio_pdf(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta para salvar o relatório e os gráficos")
        if not pasta:
            return

        pasta = Path(pasta)
        pasta.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        png1 = pasta / f"{timestamp}_grafico_pontuacao_por_indicador.png"
        png2 = pasta / f"{timestamp}_grafico_busca_ativa_por_indicador.png"
        png3 = pasta / f"{timestamp}_grafico_distribuicao_classificacao.png"
        png4 = pasta / f"{timestamp}_grafico_delta_media.png"

        try:
            self.fig1.savefig(png1, dpi=180, bbox_inches="tight")
            self.fig2.savefig(png2, dpi=180, bbox_inches="tight")
            self.fig3.savefig(png3, dpi=180, bbox_inches="tight")
            self.fig4.savefig(png4, dpi=180, bbox_inches="tight")

            compare_png1 = None
            compare_png2 = None
            if hasattr(self, "figm1"):
                compare_png1 = pasta / f"{timestamp}_grafico_comparacao_media.png"
                self.figm1.savefig(compare_png1, dpi=180, bbox_inches="tight")
            if hasattr(self, "figm2"):
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

            doc = SimpleDocTemplate(
                str(pdf_path),
                pagesize=landscape(A4),
                rightMargin=1.2 * cm,
                leftMargin=1.2 * cm,
                topMargin=1.2 * cm,
                bottomMargin=1.2 * cm,
            )

            story = []
            story.append(Paragraph("Relatório Profissional - Dashboard APS", styles["Title"]))
            story.append(Paragraph(
                f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')} | Pasta de resultados: {self.results_dir}",
                styles["APSBody"]
            ))
            story.append(Spacer(1, 0.25 * cm))

            cards_df = pd.DataFrame([{
                "Arquivos atuais": self.card_vars["arquivos"].get(),
                "Pacientes atuais": self.card_vars["linhas"].get(),
                "Busca ativa atual": self.card_vars["busca"].get(),
                "Média atual": self.card_vars["media"].get(),
                "Δ média vs anterior": self.card_vars["delta"].get(),
                "Δ busca ativa": self.card_vars["delta_busca"].get(),
            }])
            story.extend(self._make_pdf_table(cards_df, "Resumo executivo", max_rows=1, small=False))

            if self.summary_df is not None and not self.summary_df.empty:
                story.extend(self._make_pdf_table(self.summary_df, "Tabela 1 - Resumo atual por indicador", max_rows=20, small=True))

            if self.compare_df is not None and not self.compare_df.empty:
                story.extend(self._make_pdf_table(self.compare_df, "Tabela 2 - Comparativo atual x anterior", max_rows=20, small=True))

            for img_path, title in [
                (png1, "Gráfico 1 - Pontuação média atual por indicador"),
                (png2, "Gráfico 2 - Busca ativa atual por indicador"),
                (png3, "Gráfico 3 - Distribuição atual por classificação"),
                (png4, "Gráfico 4 - Variação da média: atual x anterior"),
            ]:
                story.append(Paragraph(title, styles["Heading2"]))
                story.append(Spacer(1, 0.12 * cm))
                story.append(Image(str(img_path), width=24.5 * cm, height=9.2 * cm))
                story.append(Spacer(1, 0.28 * cm))

            if self.path_a.get().strip() and self.path_b.get().strip():
                story.append(PageBreak())
                story.append(Paragraph("Comparação manual entre arquivos", styles["Heading1"]))

                manual_df = pd.DataFrame([{
                    "Arquivo A": self.summary_a.get("arquivo", "-"),
                    "Total A": self.summary_a.get("total", "-"),
                    "Média A": self.summary_a.get("media", "-"),
                    "Busca ativa A": self.summary_a.get("busca", "-"),
                    "Arquivo B": self.summary_b.get("arquivo", "-"),
                    "Total B": self.summary_b.get("total", "-"),
                    "Média B": self.summary_b.get("media", "-"),
                    "Busca ativa B": self.summary_b.get("busca", "-"),
                }])
                story.extend(self._make_pdf_table(manual_df, "Tabela 3 - Resumo da comparação manual", max_rows=1, small=True))

                rows = []
                for item in self.tree_manual.get_children():
                    rows.append(self.tree_manual.item(item, "values"))
                if rows:
                    df_manual_tree = pd.DataFrame(rows, columns=["Classificação", "Arquivo A", "Arquivo B", "Δ"])
                    story.extend(self._make_pdf_table(df_manual_tree, "Tabela 4 - Distribuição por classificação", max_rows=20, small=True))

                if compare_png1:
                    story.append(Paragraph("Gráfico 5 - Comparação de pontuação média", styles["Heading2"]))
                    story.append(Image(str(compare_png1), width=24.5 * cm, height=8.8 * cm))
                    story.append(Spacer(1, 0.25 * cm))
                if compare_png2:
                    story.append(Paragraph("Gráfico 6 - Comparação por classificação", styles["Heading2"]))
                    story.append(Image(str(compare_png2), width=24.5 * cm, height=8.8 * cm))

            doc.build(story)
            messagebox.showinfo("Exportação concluída", f"Relatório PDF e gráficos exportados com sucesso em:\n{pasta}")
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))
    def _build_folders_tab(self, root):
        tk.Label(root,
                 text="Compare os resultados de dois períodos diferentes (ex: mês anterior × mês atual).",
                 font=("Segoe UI", 9), anchor="w").pack(fill="x", padx=8, pady=(10, 4))

        frm = ttk.LabelFrame(root, text="Selecionar pastas")
        frm.pack(fill="x", padx=8, pady=6)
        frm.columnconfigure(1, weight=1)

        for row_i, (lbl_var, name_var, pick_cmd, label) in enumerate([
            (self.folder_a_var, self.folder_a_name, self._pick_folder_a, "Período A (anterior):"),
            (self.folder_b_var, self.folder_b_name, self._pick_folder_b, "Período B (atual):"),
        ]):
            tk.Label(frm, text=label, font=("Segoe UI", 9, "bold")).grid(
                row=row_i*2, column=0, sticky="w", padx=8, pady=(8, 0))
            ttk.Entry(frm, textvariable=lbl_var, state="readonly").grid(
                row=row_i*2, column=1, sticky="ew", padx=6)
            tk.Button(frm, text="Escolher pasta", command=pick_cmd,
                      bg="#2E75B6", fg="white", font=("Segoe UI", 8)).grid(
                row=row_i*2, column=2, padx=4)
            tk.Label(frm, text="  Rótulo:").grid(row=row_i*2+1, column=0, sticky="e", padx=8, pady=(0, 6))
            ttk.Entry(frm, textvariable=name_var, width=22).grid(
                row=row_i*2+1, column=1, sticky="w", padx=6, pady=(0, 6))

        info = ttk.LabelFrame(root, text="O que será comparado")
        info.pack(fill="x", padx=8, pady=4)
        tk.Label(info,
                 text=("• Lê automaticamente os arquivos C1–C7 mais recentes de cada pasta\n"
                       "• Compara total de pacientes, pontuação média, busca ativa e classificações\n"
                       "• Δ verde = melhora   •   Δ vermelho = piora   •   Exporta planilha com cores automáticas"),
                 font=("Segoe UI", 9), justify="left").pack(padx=10, pady=8, anchor="w")

        bot = tk.Frame(root)
        bot.pack(fill="x", padx=8, pady=8)
        tk.Label(bot, textvariable=self.folders_status, font=("Segoe UI", 9), anchor="w").pack(
            side="left", fill="x", expand=True)
        tk.Button(bot, text="▶  Gerar comparação entre pastas", command=self._run_folders,
                  bg="#7030A0", fg="white", font=("Segoe UI", 10, "bold")).pack(side="right")

    def _pick_folder_a(self):
        f = filedialog.askdirectory(title="Pasta Período A", initialdir=str(self.results_dir))
        if f: self.folder_a_var.set(f)

    def _pick_folder_b(self):
        f = filedialog.askdirectory(title="Pasta Período B", initialdir=str(self.results_dir))
        if f: self.folder_b_var.set(f)

    def _run_folders(self):
        fa = self.folder_a_var.get().strip()
        fb = self.folder_b_var.get().strip()
        if fa in ("(não selecionada)", "") or fb in ("(não selecionada)", ""):
            messagebox.showwarning("Atenção", "Selecione as duas pastas.")
            return
        self.folders_status.set("Processando…"); self.update()
        try:
            from aps_comparador_paciente import build_folder_comparison, export_folder_comparison_excel
            la = self.folder_a_name.get().strip() or "Período A"
            lb = self.folder_b_name.get().strip() or "Período B"
            df = build_folder_comparison(Path(fa), Path(fb), la, lb)
            if df.empty:
                messagebox.showwarning("Resultado vazio", "Nenhum arquivo C1–C7 encontrado nas pastas.")
                return
            import datetime as dt
            stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            out = self.results_dir / f"COMPARACAO_PASTAS_{stamp}.xlsx"
            export_folder_comparison_excel(df, out, la, lb)
            self.folders_status.set(f"✔ {out.name}")
            messagebox.showinfo("Concluído ✔", f"Comparação gerada:\n\n{out}")
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
            messagebox.showinfo("Histórico", f"Sem histórico para {code}.")
            return
        win = tk.Toplevel(self)
        win.title(f"Histórico - {code}")
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
        raise FileNotFoundError(f"Pasta de resultados não encontrada: {results_dir}")

    root = tk._default_root
    if root is None:
        root = tk.Tk()
        root.withdraw()
    win = APSDashboard(root, results_dir)
    return win


def main():
    results_dir = None
    if len(sys.argv) > 1:
        results_dir = Path(sys.argv[1])
    root = tk.Tk()
    root.withdraw()
    try:
        launch_dashboard(results_dir)
        root.mainloop()
    except Exception as exc:
        messagebox.showerror("Erro", str(exc))


if __name__ == "__main__":
    main()

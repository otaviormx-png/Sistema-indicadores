from __future__ import annotations

import argparse
import json
from collections import Counter
from datetime import datetime
import os
from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from aps_utils import infer_indicator_code_from_path
import aps_aprazamento
import aps_comparador_paciente


CLASS_ORDER = ["Otimo", "Bom", "Suficiente", "Regular", "Ruim", "Critico"]
CODE_LABEL = {
    "C1": "C1 Mais acesso",
    "C2": "C2 Infantil",
    "C3": "C3 Gestacao",
    "C4": "C4 Diabetes",
    "C5": "C5 Hipertensao",
    "C6": "C6 Idoso",
    "C7": "C7 Mulher",
}


def _norm_text(value) -> str:
    txt = str(value or "").strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def _norm_person(value) -> str:
    txt = _norm_text(value).lower()
    txt = re.sub(r"[^a-z0-9]+", "", txt)
    return txt


def _pick_col(df: pd.DataFrame, *aliases: str) -> str | None:
    norm_map = {re.sub(r"[^a-z0-9]+", "", str(c).lower()): c for c in df.columns}
    for alias in aliases:
        key = re.sub(r"[^a-z0-9]+", "", alias.lower())
        if key in norm_map:
            return norm_map[key]
    return None


def indicator_files(results_dir: Path) -> dict[str, Path]:
    latest: dict[str, Path] = {}
    for p in results_dir.glob("*.xlsx"):
        # Ignora lock files temporarios do Excel.
        if p.name.startswith("~$"):
            continue
        code = infer_indicator_code_from_path(p)
        if not code:
            continue
        cur = latest.get(code)
        if cur is None or p.stat().st_mtime > cur.stat().st_mtime:
            latest[code] = p
    return dict(sorted(latest.items()))


def read_indicator_dataframe(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheet = next((s for s in xls.sheet_names if "dados" in str(s).lower()), xls.sheet_names[0])
    for header_try in (2, 1, 0):
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=header_try, engine="openpyxl")
            df = df.dropna(how="all")
            if len(df.columns) > 1:
                if "Nome" in df.columns:
                    df = df[df["Nome"].astype(str).str.strip().ne("")]
                return df.reset_index(drop=True)
        except Exception:
            continue
    return pd.read_excel(path, sheet_name=sheet, engine="openpyxl").dropna(how="all").reset_index(drop=True)


def build_indicator_summary(results_dir: Path, warnings: list[str] | None = None) -> pd.DataFrame:
    rows = []
    for code, path in indicator_files(results_dir).items():
        try:
            df = read_indicator_dataframe(path)
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"{path.name}: {exc}")
            continue
        total = len(df)
        col_pts = _pick_col(df, "Pontuacao", "Pontuação")
        col_cls = _pick_col(df, "Classificacao", "Classificação")
        pts = pd.to_numeric(df[col_pts] if col_pts else pd.Series(dtype=float), errors="coerce").fillna(0.0)
        cls = (df[col_cls] if col_cls else pd.Series(dtype=str)).fillna("Regular").astype(str)
        c = Counter(cls)
        row = {
            "Indicador": code,
            "Label": CODE_LABEL.get(code, code),
            "Arquivo": path.name,
            "Total": int(total),
            "Busca Ativa": int((pts < 100).sum()) if total else 0,
            "Media": round(float(pts.mean()) if total else 0.0, 1),
        }
        for k in CLASS_ORDER:
            row[k] = int(c.get(k, 0))
        rows.append(row)
    return pd.DataFrame(rows)


def load_aprazamento_summary(results_dir: Path) -> dict[str, int]:
    out = {"total": 0, "vencido": 0, "vermelho": 0, "amarelo": 0, "verde": 0, "sem_data": 0}
    path = results_dir / "aprazamento_controle.json"
    if not path.exists():
        return out
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return out
    patients = payload.get("patients", [])
    out["total"] = int(len(patients))
    for rec in patients:
        sem = str(rec.get("semaphore", "")).upper().strip()
        if sem == "VENCIDO":
            out["vencido"] += 1
        elif sem == "VERMELHO":
            out["vermelho"] += 1
        elif sem == "AMARELO":
            out["amarelo"] += 1
        elif sem == "VERDE":
            out["verde"] += 1
        else:
            out["sem_data"] += 1
    return out


def _norm_prio(raw: str) -> str:
    txt = str(raw or "").upper()
    if "URG" in txt or "ALTA" == txt.strip():
        return "URGENTE"
    if "ALTA" in txt:
        return "ALTA"
    if "MONI" in txt or "MEDIA" in txt or "MÉDIA" in txt or "BAIXA" in txt:
        return "MONITORAR"
    if "CONCL" in txt or "OTIMO" in txt or "ÓTIMO" in txt:
        return "CONCLUIDO"
    return "MONITORAR"


def _prio_order(prio: str) -> int:
    return {"URGENTE": 0, "ALTA": 1, "MONITORAR": 2, "CONCLUIDO": 3}.get(prio, 9)


def build_operational_queue(results_dir: Path) -> pd.DataFrame:
    files = list(indicator_files(results_dir).values())
    if not files:
        return pd.DataFrame()
    raw = aps_comparador_paciente.build_unified(files)
    if raw is None or raw.empty:
        return pd.DataFrame()

    df = raw.copy()
    df["prioridade_norm"] = df.get("prioridade", pd.Series(dtype=str)).map(_norm_prio)
    df["pontuacao_num"] = pd.to_numeric(df.get("pontuacao_media", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    df["pendencias_num"] = pd.to_numeric(df.get("pendencias_total", pd.Series(dtype=float)), errors="coerce").fillna(0).astype(int)
    df["nome_norm"] = df.get("nome", pd.Series(dtype=str)).map(_norm_person)
    df = df.sort_values(
        ["prioridade_norm", "pendencias_num", "pontuacao_num", "nome_norm"],
        key=lambda s: s.map(_prio_order) if s.name == "prioridade_norm" else s,
        ascending=[True, False, True, True],
    )

    out = pd.DataFrame(
        {
            "Prioridade": df["prioridade_norm"],
            "Nome": df.get("nome", ""),
            "Bairro": df.get("bairro", "").fillna(""),
            "Pendencias": df["pendencias_num"],
            "Media": df["pontuacao_num"].round(1),
            "Indicadores": df.get("indicadores", "").fillna(""),
            "Acoes": df.get("acoes", "").fillna(""),
        }
    )
    return out


class APSDashboardV3(tk.Toplevel):
    def __init__(self, parent: tk.Misc | None, results_dir: Path):
        super().__init__(parent)
        self.results_dir = Path(results_dir)
        self.title("APS - Dashboard v3")
        self.geometry("1320x840")
        self.minsize(1060, 680)
        self.configure(bg="#EEF4F8")

        self.status_var = tk.StringVar(value="Pronto para atualizar.")
        self.path_var = tk.StringVar(value=str(self.results_dir))
        self.queue_filter_var = tk.StringVar()
        self.queue_sort_var = tk.StringVar(value="Urgencia")

        self.card_vars = {
            "indicadores": tk.StringVar(value="0"),
            "pacientes": tk.StringVar(value="0"),
            "unicos": tk.StringVar(value="0"),
            "busca": tk.StringVar(value="0"),
            "media": tk.StringVar(value="0.0"),
            "ap_total": tk.StringVar(value="0"),
            "ap_vencido": tk.StringVar(value="0"),
            "ap_alerta": tk.StringVar(value="0"),
        }

        self.summary_df = pd.DataFrame()
        self.queue_df = pd.DataFrame()

        self._build_ui()
        self._set_idle_state()

    def _build_ui(self):
        top = tk.Frame(self, bg="#1F4E79")
        top.pack(fill="x")
        tk.Label(top, text="APS - DASHBOARD V3", bg="#1F4E79", fg="white", font=("Segoe UI", 14, "bold"), pady=10).pack(fill="x")

        bar = tk.Frame(self, bg="#DCE6F1")
        bar.pack(fill="x", padx=8, pady=(6, 4))
        tk.Label(bar, text="Pasta:", bg="#DCE6F1").pack(side="left")
        tk.Entry(bar, textvariable=self.path_var).pack(side="left", fill="x", expand=True, padx=(6, 6))
        ttk.Button(bar, text="Escolher", command=self._choose_folder).pack(side="left")
        ttk.Button(bar, text="Atualizar", command=self.refresh).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Exportar Excel", command=self._export_excel).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Comparar pastas", command=self._compare_folders).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Abrir Aprazamento", command=self._open_aprazamento).pack(side="left", padx=(6, 0))

        cards = tk.Frame(self, bg="#EEF4F8")
        cards.pack(fill="x", padx=8, pady=(2, 4))
        self._card(cards, "Indicadores", self.card_vars["indicadores"], 0)
        self._card(cards, "Pacientes", self.card_vars["pacientes"], 1)
        self._card(cards, "Unicos", self.card_vars["unicos"], 2)
        self._card(cards, "Busca ativa", self.card_vars["busca"], 3)
        self._card(cards, "Media", self.card_vars["media"], 4)
        self._card(cards, "Apraz total", self.card_vars["ap_total"], 5)
        self._card(cards, "Apraz vencido", self.card_vars["ap_vencido"], 6)
        self._card(cards, "Apraz alerta", self.card_vars["ap_alerta"], 7)

        body = tk.Frame(self, bg="#EEF4F8")
        body.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        left_top = ttk.LabelFrame(body, text="Resumo por indicador")
        left_top.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        left_top.rowconfigure(0, weight=1)
        left_top.columnconfigure(0, weight=1)
        cols = ("Indicador", "Total", "Busca Ativa", "Media", "Otimo", "Bom", "Suficiente", "Regular")
        self.tree_summary = ttk.Treeview(left_top, columns=cols, show="headings", height=9)
        for c in cols:
            self.tree_summary.heading(c, text=c)
            self.tree_summary.column(c, width=95, anchor="center")
        self.tree_summary.column("Indicador", width=140, anchor="w")
        self.tree_summary.grid(row=0, column=0, sticky="nsew")
        sy = ttk.Scrollbar(left_top, orient="vertical", command=self.tree_summary.yview)
        self.tree_summary.configure(yscrollcommand=sy.set)
        sy.grid(row=0, column=1, sticky="ns")

        right_top = ttk.LabelFrame(body, text="Fila operacional")
        right_top.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))
        right_top.columnconfigure(0, weight=1)
        right_top.rowconfigure(1, weight=1)
        f = tk.Frame(right_top)
        f.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        tk.Label(f, text="Buscar:").pack(side="left")
        ent = tk.Entry(f, textvariable=self.queue_filter_var, width=28)
        ent.pack(side="left", padx=(6, 8))
        ent.bind("<KeyRelease>", lambda _e: self._refresh_queue_view())
        tk.Label(f, text="Ordenar:").pack(side="left")
        cb = ttk.Combobox(f, textvariable=self.queue_sort_var, values=("Urgencia", "Pendencias", "Pontuacao", "Alfabetica"), state="readonly", width=12)
        cb.pack(side="left", padx=(6, 0))
        cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_queue_view())

        qcols = ("Prioridade", "Nome", "Bairro", "Pendencias", "Media", "Indicadores", "Acoes")
        self.tree_queue = ttk.Treeview(right_top, columns=qcols, show="headings", height=10)
        for c in qcols:
            self.tree_queue.heading(c, text=c)
            self.tree_queue.column(c, width=100, anchor="center")
        self.tree_queue.column("Nome", width=220, anchor="w")
        self.tree_queue.column("Acoes", width=240, anchor="w")
        self.tree_queue.tag_configure("urgente", background="#FDECEA")
        self.tree_queue.tag_configure("alta", background="#FFF4E5")
        self.tree_queue.tag_configure("monitorar", background="#FFFBE6")
        self.tree_queue.tag_configure("concluido", background="#EAF7EA")
        self.tree_queue.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        qy = ttk.Scrollbar(right_top, orient="vertical", command=self.tree_queue.yview)
        self.tree_queue.configure(yscrollcommand=qy.set)
        qy.place(relx=1.0, rely=0.15, relheight=0.82, anchor="ne")

        left_bottom = ttk.LabelFrame(body, text="Busca ativa por indicador")
        left_bottom.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(6, 0))
        self.fig1 = Figure(figsize=(6, 3.0), dpi=100)
        self.ax1 = self.fig1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master=left_bottom)
        self.canvas1.get_tk_widget().pack(fill="both", expand=True)

        right_bottom = ttk.LabelFrame(body, text="Distribuicao de classificacao")
        right_bottom.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(6, 0))
        self.fig2 = Figure(figsize=(6, 3.0), dpi=100)
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, master=right_bottom)
        self.canvas2.get_tk_widget().pack(fill="both", expand=True)

        tk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken").pack(fill="x", padx=8, pady=(0, 8))

    def _set_idle_state(self):
        self.summary_df = pd.DataFrame()
        self.queue_df = pd.DataFrame()
        for iid in self.tree_summary.get_children():
            self.tree_summary.delete(iid)
        for iid in self.tree_queue.get_children():
            self.tree_queue.delete(iid)
        for key in self.card_vars:
            self.card_vars[key].set("0" if key != "media" else "0.0")
        self.ax1.clear()
        self.ax2.clear()
        self.ax1.set_title("Clique em Atualizar para carregar")
        self.ax2.set_title("Clique em Atualizar para carregar")
        self.canvas1.draw()
        self.canvas2.draw()
        self.status_var.set("Dashboard pronto. Clique em Atualizar para carregar os dados.")

    def _card(self, parent, title: str, var: tk.StringVar, col: int):
        box = tk.Frame(parent, bg="#DCE6F1", bd=1, relief="solid")
        box.grid(row=0, column=col, sticky="ew", padx=4, pady=4)
        parent.columnconfigure(col, weight=1)
        tk.Label(box, text=title, bg="#DCE6F1", fg="#1F1F1F", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=8, pady=(8, 2))
        tk.Label(box, textvariable=var, bg="#DCE6F1", fg="#1F4E79", font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=8, pady=(0, 8))

    def _choose_folder(self):
        chosen = filedialog.askdirectory(title="Selecione a pasta de resultados", initialdir=self.path_var.get().strip() or str(self.results_dir))
        if not chosen:
            return
        self.path_var.set(chosen)
        self.refresh()

    def _open_aprazamento(self):
        try:
            win = aps_aprazamento.launch_aprazamento(master=self, base_dir=self.results_dir, auto_import=True)
            if win:
                win.lift()
                win.focus_force()
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _export_excel(self):
        if (self.summary_df is None or self.summary_df.empty) and (self.queue_df is None or self.queue_df.empty):
            messagebox.showwarning("Sem dados", "Atualize o dashboard antes de exportar.")
            return
        out_dir = filedialog.askdirectory(
            title="Escolha a pasta para exportar",
            initialdir=str(self.results_dir),
        )
        if not out_dir:
            return
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = out_dir / f"DASHBOARD_V3_{ts}.xlsx"
        ap = load_aprazamento_summary(self.results_dir)
        ap_df = pd.DataFrame(
            [
                {
                    "Apraz total": ap["total"],
                    "Apraz vencido": ap["vencido"],
                    "Apraz vermelho": ap["vermelho"],
                    "Apraz amarelo": ap["amarelo"],
                    "Apraz verde": ap["verde"],
                    "Apraz sem data": ap["sem_data"],
                    "Fonte": str(self.results_dir),
                    "Gerado em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                }
            ]
        )
        try:
            with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
                if self.summary_df is not None and not self.summary_df.empty:
                    self.summary_df.to_excel(writer, sheet_name="Resumo_Indicadores", index=False)
                if self.queue_df is not None and not self.queue_df.empty:
                    self.queue_df.to_excel(writer, sheet_name="Fila_Operacional", index=False)
                ap_df.to_excel(writer, sheet_name="Aprazamento", index=False)
            self.status_var.set(f"Exportado: {out_file}")
            try:
                os.startfile(out_file)
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))

    def _compare_folders(self):
        folder_a = filedialog.askdirectory(
            title="Selecione a pasta Periodo A (anterior)",
            initialdir=str(self.results_dir),
        )
        if not folder_a:
            return
        folder_b = filedialog.askdirectory(
            title="Selecione a pasta Periodo B (atual)",
            initialdir=str(self.results_dir),
        )
        if not folder_b:
            return
        pa = Path(folder_a)
        pb = Path(folder_b)
        label_a = pa.name or "Periodo A"
        label_b = pb.name or "Periodo B"
        try:
            df_cmp = aps_comparador_paciente.build_folder_comparison(pa, pb, label_a, label_b)
        except Exception as exc:
            messagebox.showerror("Erro ao comparar pastas", str(exc))
            return
        if df_cmp is None or df_cmp.empty:
            messagebox.showwarning("Sem dados", "Nao foram encontrados arquivos C1..C7 para comparar nas pastas selecionadas.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = self.results_dir / f"COMPARACAO_PASTAS_V3_{ts}.xlsx"
        try:
            aps_comparador_paciente.export_folder_comparison_excel(df_cmp, out_file, label_a, label_b)
            self.status_var.set(f"Comparacao de pastas gerada: {out_file.name}")
            try:
                os.startfile(out_file)
            except Exception:
                pass
            delta_media = pd.to_numeric(df_cmp.get("Variacao Media"), errors="coerce").dropna()
            delta_busca = pd.to_numeric(df_cmp.get("Variacao Busca"), errors="coerce").dropna()
            dm = float(delta_media.mean()) if not delta_media.empty else 0.0
            db = int(round(float(delta_busca.sum()))) if not delta_busca.empty else 0
            messagebox.showinfo(
                "Comparacao concluida",
                f"Arquivo gerado:\n{out_file}\n\nMedia (B-A): {dm:+.1f}\nBusca ativa (B-A): {db:+d}",
            )
        except Exception as exc:
            messagebox.showerror("Erro ao exportar comparacao", str(exc))

    def _refresh_queue_view(self):
        for iid in self.tree_queue.get_children():
            self.tree_queue.delete(iid)
        if self.queue_df is None or self.queue_df.empty:
            return

        df = self.queue_df.copy()
        term = self.queue_filter_var.get().strip().lower()
        if term:
            mask = (
                df["Nome"].astype(str).str.lower().str.contains(term, na=False)
                | df["Bairro"].astype(str).str.lower().str.contains(term, na=False)
                | df["Indicadores"].astype(str).str.lower().str.contains(term, na=False)
            )
            df = df[mask]

        sort_mode = self.queue_sort_var.get()
        if sort_mode == "Pendencias":
            df = df.sort_values(["Pendencias", "Media"], ascending=[False, True])
        elif sort_mode == "Pontuacao":
            df = df.sort_values(["Media", "Pendencias"], ascending=[True, False])
        elif sort_mode == "Alfabetica":
            df = df.sort_values(["Nome"], ascending=[True])
        else:
            df = df.sort_values(["Prioridade", "Pendencias", "Media"], key=lambda s: s.map(_prio_order) if s.name == "Prioridade" else s, ascending=[True, False, True])

        for _, row in df.head(400).iterrows():
            pr = str(row.get("Prioridade", ""))
            tag = {
                "URGENTE": "urgente",
                "ALTA": "alta",
                "MONITORAR": "monitorar",
                "CONCLUIDO": "concluido",
            }.get(pr, "monitorar")
            self.tree_queue.insert(
                "",
                "end",
                values=(
                    pr,
                    row.get("Nome", ""),
                    row.get("Bairro", ""),
                    int(row.get("Pendencias", 0)),
                    float(row.get("Media", 0.0)),
                    row.get("Indicadores", ""),
                    row.get("Acoes", ""),
                ),
                tags=(tag,),
            )

    def _refresh_charts(self):
        self.ax1.clear()
        self.ax2.clear()
        if self.summary_df is None or self.summary_df.empty:
            self.ax1.set_title("Sem dados")
            self.ax2.set_title("Sem dados")
            self.canvas1.draw()
            self.canvas2.draw()
            return

        labels = self.summary_df["Label"].astype(str).tolist()
        busca = pd.to_numeric(self.summary_df["Busca Ativa"], errors="coerce").fillna(0).tolist()
        bars = self.ax1.bar(labels, busca)
        self.ax1.set_title("Busca ativa por indicador")
        self.ax1.set_ylabel("Pacientes")
        self.ax1.tick_params(axis="x", labelrotation=20)
        self.ax1.bar_label(bars, fmt="%d", padding=2, fontsize=8)

        bottom = None
        for cls in ("Otimo", "Bom", "Suficiente", "Regular"):
            vals = pd.to_numeric(self.summary_df.get(cls, pd.Series(dtype=float)), errors="coerce").fillna(0)
            self.ax2.bar(labels, vals, bottom=bottom, label=cls)
            bottom = vals if bottom is None else (bottom + vals)
        self.ax2.set_title("Distribuicao de classificacao")
        self.ax2.set_ylabel("Pacientes")
        self.ax2.tick_params(axis="x", labelrotation=20)
        self.ax2.legend(fontsize=8)

        self.fig1.tight_layout()
        self.fig2.tight_layout()
        self.canvas1.draw()
        self.canvas2.draw()

    def refresh(self):
        self.results_dir = Path(self.path_var.get().strip() or self.results_dir)
        if not self.results_dir.exists():
            messagebox.showerror("Erro", f"Pasta nao encontrada: {self.results_dir}")
            return
        try:
            warnings: list[str] = []
            self.summary_df = build_indicator_summary(self.results_dir, warnings=warnings)
            self.queue_df = build_operational_queue(self.results_dir)
            ap = load_aprazamento_summary(self.results_dir)
        except Exception as exc:
            messagebox.showerror("Erro ao atualizar", str(exc))
            return

        for iid in self.tree_summary.get_children():
            self.tree_summary.delete(iid)
        for _, row in self.summary_df.iterrows():
            self.tree_summary.insert(
                "",
                "end",
                values=(
                    row.get("Label", row.get("Indicador", "")),
                    int(row.get("Total", 0)),
                    int(row.get("Busca Ativa", 0)),
                    float(row.get("Media", 0.0)),
                    int(row.get("Otimo", 0)),
                    int(row.get("Bom", 0)),
                    int(row.get("Suficiente", 0)),
                    int(row.get("Regular", 0)),
                ),
            )

        total = int(self.summary_df["Total"].sum()) if not self.summary_df.empty else 0
        busca = int(self.summary_df["Busca Ativa"].sum()) if not self.summary_df.empty else 0
        media = float(self.summary_df["Media"].mean()) if not self.summary_df.empty else 0.0
        unique = 0
        if not self.queue_df.empty and "Nome" in self.queue_df.columns:
            unique = int(self.queue_df["Nome"].astype(str).str.lower().str.replace(r"\\s+", " ", regex=True).nunique())

        self.card_vars["indicadores"].set(str(len(self.summary_df)))
        self.card_vars["pacientes"].set(str(total))
        self.card_vars["unicos"].set(str(unique))
        self.card_vars["busca"].set(str(busca))
        self.card_vars["media"].set(f"{media:.1f}")
        self.card_vars["ap_total"].set(str(ap["total"]))
        self.card_vars["ap_vencido"].set(str(ap["vencido"]))
        self.card_vars["ap_alerta"].set(str(ap["vermelho"] + ap["amarelo"]))

        self._refresh_queue_view()
        self._refresh_charts()
        status = f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | Pasta: {self.results_dir}"
        if warnings:
            status += f" | Ignorados {len(warnings)} arquivo(s) invalido(s)"
        self.status_var.set(status)


def launch_dashboard(results_dir: Path | None = None):
    results_dir = Path(results_dir or (Path.home() / "Desktop" / "APS_RESULTADOS"))
    if not results_dir.exists():
        raise FileNotFoundError(f"Pasta de resultados nao encontrada: {results_dir}")

    root = tk._default_root
    if root is None:
        root = tk.Tk()
        root.withdraw()
    return APSDashboardV3(root, results_dir)


def main():
    parser = argparse.ArgumentParser(description="APS Dashboard v3")
    parser.add_argument("results_dir", nargs="?", default=str(Path.home() / "Desktop" / "APS_RESULTADOS"))
    args = parser.parse_args()

    root = tk.Tk()
    root.withdraw()
    try:
        launch_dashboard(Path(args.results_dir))
        root.mainloop()
    except Exception as exc:
        messagebox.showerror("Erro", str(exc))


if __name__ == "__main__":
    main()

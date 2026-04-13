from __future__ import annotations

import re
import threading
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import aps_comparador_paciente as comparador
from aps_utils import infer_indicator_code_from_path


def _nkey(text: str) -> str:
    s = str(text or "").strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip()


def _to_num(value) -> float:
    txt = str(value or "").replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", txt)
    if not m:
        return 0.0
    try:
        return float(m.group(0))
    except Exception:
        return 0.0


class DashboardLiteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("APS - Dashboard Lite")
        self.geometry("1300x820")
        self.configure(bg="#EEF4F8")
        try:
            self.state("zoomed")
        except Exception:
            pass

        self.source_dir_var = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.filter_var = tk.StringVar()
        self.sort_var = tk.StringVar(value="Urgencia")
        self.status_var = tk.StringVar(value="Selecione uma pasta ou arquivos para carregar.")

        self.card_total = tk.StringVar(value="0")
        self.card_urg = tk.StringVar(value="0")
        self.card_alta = tk.StringVar(value="0")
        self.card_mon = tk.StringVar(value="0")
        self.card_conc = tk.StringVar(value="0")
        self.card_media = tk.StringVar(value="0.0")

        self.records: list[dict] = []
        self.filtered: list[dict] = []
        self.current_files: list[Path] = []
        self._busy = False

        self._build_ui()

    def _build_ui(self):
        top = tk.Frame(self, bg="#EEF4F8")
        top.pack(fill="x", padx=12, pady=12)
        tk.Label(top, text="APS - DASHBOARD LITE", bg="#1F4E79", fg="white", font=("Segoe UI", 13, "bold"), pady=10).pack(fill="x")

        controls = tk.Frame(self, bg="#EEF4F8")
        controls.pack(fill="x", padx=12, pady=(0, 10))
        tk.Button(controls, text="Selecionar pasta", command=self.choose_folder).pack(side="left")
        tk.Button(controls, text="Selecionar arquivos", command=self.choose_files).pack(side="left", padx=(6, 0))
        tk.Button(controls, text="Carregar", command=self.load_dashboard).pack(side="left", padx=(6, 0))
        tk.Label(controls, text="Filtro:", bg="#EEF4F8").pack(side="left", padx=(16, 4))
        ent = tk.Entry(controls, textvariable=self.filter_var, width=32)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda _e: self._apply_filter())
        tk.Label(controls, text="Ordenar:", bg="#EEF4F8").pack(side="left", padx=(10, 4))
        cb = ttk.Combobox(controls, textvariable=self.sort_var, values=("Urgencia", "Pontuacao", "Alfabetica"), state="readonly", width=12)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda _e: self._apply_filter())

        cards = tk.Frame(self, bg="#EEF4F8")
        cards.pack(fill="x", padx=12, pady=(0, 10))
        for i in range(6):
            cards.columnconfigure(i, weight=1)
        self._card(cards, 0, "Total", self.card_total, "#DCEAF5", "#1F4E79")
        self._card(cards, 1, "Urgente", self.card_urg, "#FDECEA", "#C62828")
        self._card(cards, 2, "Alta", self.card_alta, "#FFF4E5", "#B35C00")
        self._card(cards, 3, "Monitorar", self.card_mon, "#FFFBE6", "#8D6E00")
        self._card(cards, 4, "Concluido", self.card_conc, "#EAF7EA", "#2E7D32")
        self._card(cards, 5, "Media", self.card_media, "#EFE3FF", "#6A1B9A")

        body = tk.Frame(self, bg="#EEF4F8")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        cols = ("prio", "nome", "bairro", "media", "qtd", "pend", "indicadores")
        self.tree = ttk.Treeview(body, columns=cols, show="headings")
        specs = [
            ("prio", "Prioridade", 120, "center"),
            ("nome", "Nome", 290, "w"),
            ("bairro", "Bairro", 150, "center"),
            ("media", "Media", 90, "center"),
            ("qtd", "Qtd", 70, "center"),
            ("pend", "Pendencias", 100, "center"),
            ("indicadores", "Indicadores", 220, "w"),
        ]
        for key, title, w, anc in specs:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=w, anchor=anc)
        self.tree.tag_configure("urgente", background="#FDECEA")
        self.tree.tag_configure("alta", background="#FFF4E5")
        self.tree.tag_configure("monitorar", background="#FFFBE6")
        self.tree.tag_configure("concluido", background="#EAF7EA")
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        yscroll.grid(row=0, column=1, sticky="ns")

        tk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken").pack(fill="x", side="bottom")

    def _card(self, parent, col: int, title: str, var: tk.StringVar, bg: str, fg: str):
        frm = tk.Frame(parent, bg=bg, bd=1, relief="solid")
        frm.grid(row=0, column=col, sticky="nsew", padx=4, pady=4)
        tk.Label(frm, text=title, bg=bg, fg=fg, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        tk.Label(frm, textvariable=var, bg=bg, fg=fg, font=("Segoe UI", 18, "bold")).pack(anchor="w", padx=10, pady=(0, 8))

    def choose_folder(self):
        chosen = filedialog.askdirectory(title="Escolha a pasta das planilhas", initialdir=self.source_dir_var.get().strip() or str(Path.home()))
        if chosen:
            self.source_dir_var.set(chosen)
            self.status_var.set(f"Pasta selecionada: {chosen}")

    def choose_files(self):
        files = filedialog.askopenfilenames(
            title="Selecione as planilhas",
            filetypes=[("Excel", "*.xlsx *.xls")],
            initialdir=self.source_dir_var.get().strip() or str(Path.home()),
        )
        if not files:
            return
        self.current_files = [Path(p) for p in files]
        self.status_var.set(f"{len(self.current_files)} arquivo(s) selecionado(s). Clique em Carregar.")

    def _files_from_folder(self) -> list[Path]:
        folder = Path(self.source_dir_var.get().strip())
        if not folder.exists():
            return []
        out = []
        for p in folder.glob("*.xlsx"):
            n = p.name.lower()
            if not infer_indicator_code_from_path(p):
                continue
            if "backup" in n or "interativa" in n or "cruz" in n or "compar" in n or "unificad" in n:
                continue
            out.append(p)
        return out

    def _set_busy(self, value: bool):
        self._busy = value
        self.configure(cursor="watch" if value else "")

    def load_dashboard(self):
        if self._busy:
            return
        files = [p for p in self.current_files if p.exists()]
        if not files:
            files = self._files_from_folder()
        if not files:
            messagebox.showwarning("Sem arquivos", "Nenhuma planilha valida encontrada.", parent=self)
            return
        self._set_busy(True)
        self.status_var.set("Carregando dashboard lite...")

        def worker():
            df = comparador.build_unified(files)
            if df is None or df.empty:
                return []
            records = []
            for _, row in df.iterrows():
                rec = {
                    "nome": str(row.get("Nome", "") or ""),
                    "bairro": str(row.get("Bairro", "") or ""),
                    "media": _to_num(row.get("Média", row.get("MÃ©dia", 0))),
                    "prio": str(row.get("Prioridade", "") or ""),
                    "qtd": int(_to_num(row.get("Qtd", 0))),
                    "pend": int(_to_num(row.get("Pendências", row.get("PendÃªncias", 0)))),
                    "indicadores": str(row.get("Indicadores", "") or ""),
                }
                records.append(rec)
            return records

        def on_ok(records):
            self._set_busy(False)
            self.records = records
            self._apply_filter()
            self.status_var.set(f"Dashboard carregado ({len(records)} pacientes).")

        def on_err(exc):
            self._set_busy(False)
            messagebox.showerror("Erro", str(exc), parent=self)
            self.status_var.set(f"Falha ao carregar: {exc}")

        def runner():
            try:
                data = worker()
                self.after(0, lambda: on_ok(data))
            except Exception as exc:
                self.after(0, lambda: on_err(exc))

        threading.Thread(target=runner, daemon=True).start()

    def _prio_norm(self, value: str) -> str:
        t = str(value or "").upper()
        if "URGENTE" in t:
            return "URGENTE"
        if "ALTA" in t:
            return "ALTA"
        if "MONITOR" in t:
            return "MONITORAR"
        if "CONCL" in t:
            return "CONCLUIDO"
        return t

    def _sort_key(self, rec: dict):
        mode = self.sort_var.get().strip()
        prio_order = {"URGENTE": 0, "ALTA": 1, "MONITORAR": 2, "CONCLUIDO": 3}
        pk = prio_order.get(self._prio_norm(rec.get("prio", "")), 9)
        if mode == "Alfabetica":
            return (_nkey(rec.get("nome", "")), pk, rec.get("media", 0))
        if mode == "Pontuacao":
            return (rec.get("media", 0), pk, _nkey(rec.get("nome", "")))
        return (pk, rec.get("media", 0), _nkey(rec.get("nome", "")))

    def _apply_filter(self):
        term = _nkey(self.filter_var.get())
        data = []
        for rec in self.records:
            hay = " ".join(
                [
                    _nkey(rec.get("nome", "")),
                    _nkey(rec.get("bairro", "")),
                    _nkey(rec.get("indicadores", "")),
                ]
            )
            if term and term not in hay:
                continue
            data.append(rec)
        data.sort(key=self._sort_key)
        self.filtered = data

        self.tree.delete(*self.tree.get_children(""))
        for rec in data:
            p = self._prio_norm(rec.get("prio", ""))
            tag = {"URGENTE": "urgente", "ALTA": "alta", "MONITORAR": "monitorar", "CONCLUIDO": "concluido"}.get(p, "monitorar")
            self.tree.insert(
                "",
                "end",
                values=(p, rec.get("nome", ""), rec.get("bairro", ""), int(round(rec.get("media", 0))), rec.get("qtd", 0), rec.get("pend", 0), rec.get("indicadores", "")),
                tags=(tag,),
            )
        self._refresh_cards()

    def _refresh_cards(self):
        total = len(self.filtered)
        urg = sum(1 for r in self.filtered if self._prio_norm(r.get("prio", "")) == "URGENTE")
        alta = sum(1 for r in self.filtered if self._prio_norm(r.get("prio", "")) == "ALTA")
        mon = sum(1 for r in self.filtered if self._prio_norm(r.get("prio", "")) == "MONITORAR")
        conc = sum(1 for r in self.filtered if self._prio_norm(r.get("prio", "")) == "CONCLUIDO")
        media = sum(r.get("media", 0) for r in self.filtered) / total if total else 0
        self.card_total.set(str(total))
        self.card_urg.set(str(urg))
        self.card_alta.set(str(alta))
        self.card_mon.set(str(mon))
        self.card_conc.set(str(conc))
        self.card_media.set(f"{media:.1f}")


def main():
    app = DashboardLiteApp()
    app.mainloop()


if __name__ == "__main__":
    main()

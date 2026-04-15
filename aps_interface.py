
from __future__ import annotations

import os
import subprocess
import sys
import threading
import traceback
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from sistema_aps import ROOT_DESKTOP, OUT_DIR, desktop_files, get_indicators
import aps_dashboard
import aps_clonador_interativo
import aps_comparador_paciente
import aps_aprazamento
import aps_log
import aps_tema

aps_tema.init()


def _norm_ascii(text: str) -> str:
    txt = str(text or "").strip().lower()
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode("ascii")
    return txt


def _validate_input_file(path: Path) -> str | None:
    """Verifica rapidamente se o arquivo parece um bruto valido do e-SUS.
    Retorna string de erro ou None se ok."""
    if not path.exists():
        return f"Arquivo nao encontrado: {path.name}"
    if path.stat().st_size == 0:
        return f"Arquivo vazio: {path.name}"
    if path.suffix.lower() == ".csv":
        try:
            for enc in ["utf-8-sig", "utf-8", "latin1", "cp1252"]:
                try:
                    lines = path.read_text(encoding=enc, errors="ignore").splitlines()
                    break
                except Exception:
                    continue
            data = [l for l in lines if l.strip()]
            if len(data) < 2:
                return f"{path.name}: menos de 2 linhas com conteudo."
            if "nome" not in data[0].lower():
                return f"{path.name}: coluna 'Nome' nao encontrada no cabecalho."
        except Exception as exc:
            return f"{path.name}: erro ao ler - {exc}"
    return None


class APSInterface(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("APS SUITE - Painel Principal v4")
        try:
            self.state("zoomed")
        except Exception:
            self.geometry("1180x760")
        self.minsize(980, 640)
        self.configure(bg="#EAF2F8")

        self.input_dir = tk.StringVar(value=str(ROOT_DESKTOP))
        self.output_dir = tk.StringVar(value=str(OUT_DIR))
        self.status_var = tk.StringVar(value="UI atualizada v4 carregada.")
        self.summary_var = tk.StringVar(value="Nenhum processamento executado ainda.")
        self.quick_files_var = tk.StringVar(value="0")
        self.quick_selected_var = tk.StringVar(value="0")
        self.quick_ok_var = tk.StringVar(value="0")
        self.quick_error_var = tk.StringVar(value="0")
        self.quick_last_var = tk.StringVar(value="Nenhum")
        self.indicator_vars: dict[str, tk.BooleanVar] = {}
        self.processing = False
        self._last_result_path: Path | None = None

        self._build_ui()
        self._refresh_input_files()

    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Title.TLabel", background="#1F4E79", foreground="white",
                        font=("Segoe UI", 16, "bold"), padding=12)
        style.configure("Section.TLabelframe", background="#EAF2F8")
        style.configure("Section.TLabelframe.Label", background="#EAF2F8",
                        foreground="#1F4E79", font=("Segoe UI", 11, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), foreground="#1F4E79")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 9))
        aps_tema.patch_style(style)

        ttk.Label(self, text="APS SUITE - PAINEL PRINCIPAL v4",
                  style="Title.TLabel", anchor="center").pack(fill="x")
        tk.Label(self, text="VERSAO VISUAL NOVA - Resumo rapido com cards - Ultimo resultado destacado",
                 bg="#D9EAF7", fg="#1F4E79", font=("Segoe UI", 10, "bold"), pady=6).pack(fill="x")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=12, pady=12)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self._main_canvas = tk.Canvas(body, bg="#EAF2F8", highlightthickness=0)
        self._main_canvas.grid(row=0, column=0, sticky="nsew")
        vscroll = ttk.Scrollbar(body, orient="vertical", command=self._main_canvas.yview)
        vscroll.grid(row=0, column=1, sticky="ns")
        self._main_canvas.configure(yscrollcommand=vscroll.set)

        main = ttk.Frame(self._main_canvas)
        self._main_canvas_window = self._main_canvas.create_window((0, 0), window=main, anchor="nw")
        main.bind("<Configure>", lambda _e: self._main_canvas.configure(scrollregion=self._main_canvas.bbox("all")))
        self._main_canvas.bind("<Configure>", lambda e: self._main_canvas.itemconfigure(self._main_canvas_window, width=e.width))
        self._main_canvas.bind("<Enter>", lambda _e: self.bind_all("<MouseWheel>", self._on_main_mousewheel))
        self._main_canvas.bind("<Leave>", lambda _e: self.unbind_all("<MouseWheel>"))

        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        right = ttk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)

        self._build_left_panel(left)
        self._build_right_panel(right)

    def _on_main_mousewheel(self, event):
        try:
            delta = int(-1 * (event.delta / 120))
        except Exception:
            delta = -1
        self._main_canvas.yview_scroll(delta, "units")

    def _build_left_panel(self, parent):
        # Pastas
        box_paths = ttk.LabelFrame(parent, text="Pastas", style="Section.TLabelframe")
        box_paths.pack(fill="x", pady=(0, 10))
        ttk.Label(box_paths, text="Entrada (onde estao os brutos):").grid(
            row=0, column=0, sticky="w", padx=10, pady=(10, 4))
        ttk.Entry(box_paths, textvariable=self.input_dir, width=42).grid(
            row=1, column=0, padx=10, pady=(0, 8))
        ttk.Button(box_paths, text="Escolher", command=self._choose_input).grid(
            row=1, column=1, padx=8, pady=(0, 8))
        ttk.Label(box_paths, text="Saida (onde salvar resultados):").grid(
            row=2, column=0, sticky="w", padx=10, pady=(2, 4))
        ttk.Entry(box_paths, textvariable=self.output_dir, width=42).grid(
            row=3, column=0, padx=10, pady=(0, 10))
        ttk.Button(box_paths, text="Escolher", command=self._choose_output).grid(
            row=3, column=1, padx=8, pady=(0, 10))

        # Indicadores
        box_ind = ttk.LabelFrame(parent, text="Indicadores", style="Section.TLabelframe")
        box_ind.pack(fill="x", pady=(0, 10))
        ttk.Button(box_ind, text="Marcar todos",
                   command=lambda: self._set_all(True)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ttk.Button(box_ind, text="Desmarcar",
                   command=lambda: self._set_all(False)).grid(row=0, column=1, padx=10, pady=10, sticky="w")
        for i, (cfg, _) in enumerate(get_indicators(), start=1):
            var = tk.BooleanVar(value=True)
            self.indicator_vars[cfg.code] = var
            ttk.Checkbutton(box_ind,
                            text=f"{cfg.code} - {cfg.titulo.split('|')[0].strip()}",
                            variable=var,
                            command=lambda: self.quick_selected_var.set(str(sum(1 for v in self.indicator_vars.values() if v.get())))).grid(row=i, column=0, columnspan=2,
                                               sticky="w", padx=12, pady=2)
        self.quick_selected_var.set(str(sum(1 for v in self.indicator_vars.values() if v.get())))

        # Acoes
        box_actions = ttk.LabelFrame(parent, text="Acoes", style="Section.TLabelframe")
        box_actions.pack(fill="x", pady=(0, 10))
        for text, cmd in [
            ("Atualizar arquivos",         self._refresh_input_files),
            ("Processar selecionados",     self._run_selected),
            ("Abrir pasta de resultados",  self._open_output),
            ("Abrir dashboard",            self._open_dashboard),
            ("Abrir editor da planilha APS", self._open_editor_planilha),
            ("Abrir controle de aprazamento", self._open_aprazamento),
            ("Ver historico de execucoes", self._open_historico),
        ]:
            ttk.Button(box_actions, text=text, command=cmd,
                       style="Primary.TButton").pack(fill="x", padx=10, pady=5)
        # Resumo rapido
        box_status = ttk.LabelFrame(parent, text="Painel executivo", style="Section.TLabelframe")
        box_status.pack(fill="both", expand=True)

        cards = ttk.Frame(box_status)
        cards.pack(fill="x", padx=10, pady=(10, 6))
        cards.columnconfigure((0,1), weight=1)

        self._make_summary_card(cards, "Brutos", self.quick_files_var, 0, 0)
        self._make_summary_card(cards, "Selecionados", self.quick_selected_var, 0, 1)
        self._make_summary_card(cards, "Gerados OK", self.quick_ok_var, 1, 0)
        self._make_summary_card(cards, "Erros", self.quick_error_var, 1, 1)

        ttk.Label(box_status, text="Ultimo resultado:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(6, 0))
        ttk.Label(box_status, textvariable=self.quick_last_var, wraplength=280, justify="left").pack(fill="x", padx=10)

        ttk.Separator(box_status, orient="horizontal").pack(fill="x", padx=10, pady=8)
        ttk.Label(box_status, textvariable=self.summary_var,
                  wraplength=280, justify="left").pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _build_right_panel(self, parent):
        # Arquivos detectados
        box_files = ttk.LabelFrame(parent, text="Arquivos brutos detectados",
                                   style="Section.TLabelframe")
        box_files.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        box_files.rowconfigure(0, weight=1)
        box_files.columnconfigure(0, weight=1)

        cols = ("arquivo", "ext", "mb", "validacao")
        self.tree_files = ttk.Treeview(box_files, columns=cols, show="headings", height=7)
        self.tree_files.heading("arquivo", text="Arquivo")
        self.tree_files.heading("ext", text="Ext.")
        self.tree_files.heading("mb", text="MB")
        self.tree_files.heading("validacao", text="Validacao")
        self.tree_files.column("arquivo", width=380)
        self.tree_files.column("ext", width=55, anchor="center")
        self.tree_files.column("mb", width=65, anchor="e")
        self.tree_files.column("validacao", width=200)
        self.tree_files.tag_configure("erro", foreground="#C00000")
        self.tree_files.tag_configure("ok", foreground="#276221")
        self.tree_files.grid(row=0, column=0, sticky="nsew")
        y1 = ttk.Scrollbar(box_files, orient="vertical", command=self.tree_files.yview)
        y1.grid(row=0, column=1, sticky="ns")
        self.tree_files.configure(yscrollcommand=y1.set)

        # Barra de progresso
        box_prog = ttk.LabelFrame(parent, text="Progresso", style="Section.TLabelframe")
        box_prog.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        box_prog.columnconfigure(0, weight=1)
        self._prog_label = ttk.Label(box_prog, text="Aguardando processamento...")
        self._prog_label.grid(row=0, column=0, sticky="w", padx=10, pady=(6, 2))
        self._progressbar = ttk.Progressbar(box_prog, orient="horizontal",
                                            mode="determinate")
        self._progressbar.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))

        # Log
        box_log = ttk.LabelFrame(parent, text="Log de execucao", style="Section.TLabelframe")
        box_log.grid(row=2, column=0, sticky="nsew")
        box_log.rowconfigure(0, weight=1)
        box_log.columnconfigure(0, weight=1)
        self.txt_log = tk.Text(box_log, wrap="word", font=("Consolas", 10),
                               bg="#FFFFFF", fg="#1F1F1F")
        self.txt_log.grid(row=0, column=0, sticky="nsew")
        y2 = ttk.Scrollbar(box_log, orient="vertical", command=self.txt_log.yview)
        y2.grid(row=0, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=y2.set)

        ttk.Label(parent, textvariable=self.status_var,
                  anchor="w", relief="sunken").grid(row=3, column=0, sticky="ew", pady=(8, 0))

    def _make_summary_card(self, parent, title: str, value_var: tk.StringVar, row: int, col: int):
        colors = {
            "Brutos": ("#EAF2F8", "#1F4E79"),
            "Selecionados": ("#FFF4CC", "#7F6000"),
            "Gerados OK": ("#E8F5E9", "#2E7D32"),
            "Erros": ("#FDECEA", "#C62828"),
        }
        bg, fg = colors.get(title, ("#FFFFFF", "#1F4E79"))
        frm = tk.Frame(parent, bg=bg, bd=1, relief="solid")
        frm.grid(row=row, column=col, sticky="nsew", padx=4, pady=4)
        tk.Label(frm, text=title.upper(), bg=bg, fg=fg, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        tk.Label(frm, textvariable=value_var, bg=bg, fg=fg, font=("Segoe UI", 20, "bold")).pack(anchor="w", padx=10, pady=(0, 10))

    # ------------------------------------------------------------------
    def _set_all(self, value: bool):
        for var in self.indicator_vars.values():
            var.set(value)
        self.quick_selected_var.set(str(sum(1 for v in self.indicator_vars.values() if v.get())))

    def _choose_input(self):
        d = filedialog.askdirectory(initialdir=self.input_dir.get() or str(ROOT_DESKTOP),
                                    title="Escolha a pasta com os arquivos brutos")
        if d:
            self.input_dir.set(d)
            self._refresh_input_files()

    def _choose_output(self):
        d = filedialog.askdirectory(initialdir=self.output_dir.get() or str(OUT_DIR),
                                    title="Escolha a pasta de saida")
        if d:
            self.output_dir.set(d)

    def _refresh_input_files(self):
        for item in self.tree_files.get_children():
            self.tree_files.delete(item)
        try:
            files = desktop_files(Path(self.input_dir.get()))
        except Exception as exc:
            messagebox.showerror("Erro", f"Nao foi possivel listar os arquivos.\n\n{exc}")
            return
        for f in sorted(files, key=lambda x: x.name.lower()):
            size_mb = f.stat().st_size / (1024 * 1024)
            erro = _validate_input_file(f)
            tag = "erro" if erro else "ok"
            self.tree_files.insert("", "end",
                                   values=(f.name, f.suffix.lower(),
                                           f"{size_mb:.2f}", erro if erro else "OK"),
                                   tags=(tag,))
        self.summary_var.set(
            f"Arquivos brutos detectados: {len(files)}\n"
            f"Entrada: {self.input_dir.get()}\nSaida: {self.output_dir.get()}")
        self.status_var.set("Lista de arquivos atualizada.")

    def _open_output(self):
        path = Path(self.output_dir.get())
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("Pasta de saida", str(path))

    def _open_last_result(self):
        if self._last_result_path and self._last_result_path.exists():
            try:
                os.startfile(self._last_result_path)
            except Exception:
                messagebox.showinfo("Ultimo resultado", str(self._last_result_path))
        else:
            messagebox.showinfo("Ultimo resultado", "Nenhum resultado gerado nesta sessao.")

    def _open_dashboard(self):
        try:
            dashboard_file = Path(__file__).with_name("aps_dashboard.py")
            subprocess.Popen(
                [sys.executable, str(dashboard_file), str(Path(self.output_dir.get()))],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                cwd=str(Path(__file__).parent),
            )
            self.status_var.set("Dashboard aberto em janela separada.")
        except Exception as exc:
            messagebox.showerror("Erro ao abrir dashboard", str(exc))

    def _open_cloner(self):
        try:
            win = aps_clonador_interativo.launch_clonador(self)
            if win:
                win.lift(); win.focus_force()
        except Exception as exc:
            messagebox.showerror("Erro ao abrir clonador", str(exc))

    def _open_editor_planilha(self):
        try:
            win = aps_clonador_interativo.launch_editor(self)
            if win:
                win.lift(); win.focus_force()
        except Exception as exc:
            messagebox.showerror("Erro ao abrir editor", str(exc))

    def _open_comparador_paciente(self):
        try:
            win = aps_comparador_paciente.launch_comparador(
                master=self, out_dir=Path(self.output_dir.get()))
            if win:
                win.lift(); win.focus_force()
        except Exception as exc:
            messagebox.showerror("Erro ao abrir comparador", str(exc))

    def _open_aprazamento(self):
        try:
            win = aps_aprazamento.launch_aprazamento(
                master=self,
                base_dir=Path(self.output_dir.get()),
                auto_import=True,
            )
            if win:
                win.lift(); win.focus_force()
        except Exception as exc:
            messagebox.showerror("Erro ao abrir controle de aprazamento", str(exc))

    def _open_historico(self):
        try:
            import aps_historico
            win = aps_historico.launch_historico(
                master=self, out_dir=Path(self.output_dir.get()))
            if win:
                win.lift(); win.focus_force()
        except Exception as exc:
            messagebox.showerror("Erro ao abrir historico", str(exc))

    def _append_log(self, msg: str):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.update_idletasks()

    def _run_selected(self):
        if self.processing:
            return
        selecionados = [c for c, v in self.indicator_vars.items() if v.get()]
        if not selecionados:
            messagebox.showwarning("Nada selecionado", "Selecione ao menos um indicador.")
            return

        try:
            files = desktop_files(Path(self.input_dir.get()))
        except Exception as exc:
            messagebox.showerror("Erro", f"Nao foi possivel listar arquivos.\n\n{exc}")
            return
        if not files:
            messagebox.showerror("Sem arquivos",
                                 "Nenhum arquivo bruto encontrado na pasta de entrada.\n"
                                 "Verifique a pasta e clique em 'Atualizar arquivos'.")
            return

        erros_val = [_validate_input_file(f) for f in files]
        erros_val = [e for e in erros_val if e]
        if erros_val:
            msg = "Alguns arquivos apresentaram problemas:\n\n" + \
                  "\n".join(f"  - {e}" for e in erros_val) + \
                  "\n\nDeseja continuar mesmo assim?"
            if not messagebox.askyesno("Atencao - arquivos com problemas", msg):
                return

        total = len(selecionados)
        self.quick_selected_var.set(str(total))
        self._progressbar["maximum"] = total
        self._progressbar["value"] = 0
        self._prog_label.config(text=f"0 / {total} indicadores")
        self.processing = True
        self.status_var.set("Processando...")
        self._append_log("Inicio do processamento...")

        out_dir = Path(self.output_dir.get())
        aps_log.log_session_start(out_dir, selecionados)

        processed = [0]

        def _log_wrap(msg: str):
            self.after(0, lambda m=msg: self._append_log(m))
            msg_norm = _norm_ascii(msg)
            if ("concluido" in msg_norm) or ("erro" in msg_norm) or ("nao encontrado" in msg_norm):
                processed[0] += 1
                n = processed[0]
                self.after(0, lambda c=n: (
                    self._progressbar.__setitem__("value", c),
                    self._prog_label.config(text=f"{c} / {total} indicadores"),
                    self.update_idletasks()
                ))

        def worker():
            try:
                from sistema_aps import process_selected
                results = process_selected(
                    selected_codes=selecionados,
                    in_dir=Path(self.input_dir.get()),
                    out_dir=out_dir,
                    log=_log_wrap,
                )
                aps_log.log_result(out_dir, results)
                self.after(0, lambda: self._finish_ok(results))
            except Exception as exc:
                tb = traceback.format_exc()
                self.after(0, lambda: self._finish_error(exc, tb))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_ok(self, results):
        self.processing = False
        ok = [r for r in results if r.get("status") == "ok"]
        erros = [r for r in results if r.get("status") == "erro"]
        nao_enc = [r for r in results if "nao encontrado" in _norm_ascii(r.get("status", ""))]

        self._progressbar["value"] = self._progressbar["maximum"]
        self._prog_label.config(text=f"{len(results)} / {len(results)} - concluido")

        ultimo_ok = next((r for r in reversed(ok) if r.get("saida")), None)
        if ultimo_ok:
            self._last_result_path = Path(ultimo_ok["saida"])
            self.quick_last_var.set(self._last_result_path.name)

        self.quick_ok_var.set(str(len(ok)))
        self.quick_error_var.set(str(len(erros)))
        self.summary_var.set(
            f"Ultimo processamento:\n"
            f"Gerados OK: {len(ok)}\n"
            f"Erros: {len(erros)}\n"
            f"Nao encontrados: {len(nao_enc)}\n"
            f"Pasta de saida: {self.output_dir.get()}\n"
            f"Ultimo resultado: {self._last_result_path.name if self._last_result_path else 'Nenhum'}")

        self.status_var.set(
            f"Concluido - OK {len(ok)} gerado(s)  ERRO {len(erros)} erro(s)  "
            f"ALERTA {len(nao_enc)} nao encontrado(s)")
        self._append_log("=" * 60)
        self._append_log(f"RESUMO: {len(ok)} OK | {len(erros)} ERRO | {len(nao_enc)} NAO ENCONTRADO")

        if erros:
            self._append_log("\nINDICADORES COM ERRO:")
            for r in erros:
                self._append_log(f"  {r['code']}: {(r.get('erro') or '').splitlines()[0]}")
            messagebox.showwarning("Concluido com erros",
                                   f"Processamento com {len(erros)} erro(s).\n"
                                   f"Gerados OK: {len(ok)}\n"
                                   f"Com erro: {', '.join(r['code'] for r in erros)}\n"
                                   f"Nao encontrados: {len(nao_enc)}")
        else:
            messagebox.showinfo("Concluido",
                                f"Processamento finalizado.\nArquivos gerados: {len(ok)}")

    def _finish_error(self, exc, tb):
        self.processing = False
        self._prog_label.config(text="Erro no processamento")
        self.quick_error_var.set("1+")
        self.status_var.set("Erro no processamento.")
        self._append_log(tb)
        messagebox.showerror("Erro no processamento", str(exc))


def main():
    app = APSInterface()
    app.mainloop()


if __name__ == "__main__":
    main()



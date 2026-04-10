"""
aps_cli.py — Interface de linha de comando para a APS Suite.

Permite processar indicadores sem abrir a interface gráfica.
Útil para automação via Task Scheduler, scripts ou servidores sem display.

Uso:
    python aps_cli.py --help
    python aps_cli.py --indicadores C1 C4 C5
    python aps_cli.py --indicadores todos --entrada D:\\Dados --saida D:\\Resultados
    python aps_cli.py --indicadores C4 --sem-cache
    python aps_cli.py --historico
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="aps",
        description="APS Suite — processamento de indicadores de Atenção Primária à Saúde",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python aps_cli.py --indicadores todos
  python aps_cli.py --indicadores C4 C5 --entrada C:\\Brutos --saida C:\\Resultados
  python aps_cli.py --indicadores C1 --sem-cache
  python aps_cli.py --historico
  python aps_cli.py --listar-indicadores
""",
    )
    p.add_argument(
        "--indicadores", nargs="+", metavar="CX",
        help='Códigos dos indicadores a processar (ex: C1 C4 C5) ou "todos".')
    p.add_argument(
        "--entrada", metavar="PASTA",
        help="Pasta com os arquivos brutos do e-SUS. Padrão: Desktop.")
    p.add_argument(
        "--saida", metavar="PASTA",
        help="Pasta de saída para os xlsx gerados. Padrão: Desktop/APS_RESULTADOS.")
    p.add_argument(
        "--sem-cache", action="store_true",
        help="Ignora o cache MD5 e reprocessa mesmo sem alterações.")
    p.add_argument(
        "--historico", action="store_true",
        help="Lista os últimos resultados gerados na pasta de saída.")
    p.add_argument(
        "--listar-indicadores", action="store_true",
        help="Mostra todos os indicadores disponíveis (incluindo plugins).")
    return p


def _print_historico(out_dir: Path) -> None:
    from datetime import datetime
    arquivos = sorted(
        [f for f in out_dir.glob("*.xlsx") if f.is_file()],
        key=lambda f: f.stat().st_mtime, reverse=True
    )
    if not arquivos:
        print(f"Nenhum resultado encontrado em: {out_dir}")
        return
    print(f"\nHistórico — {out_dir}\n{'─'*60}")
    for f in arquivos[:20]:
        ts = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
        size_kb = f.stat().st_size / 1024
        print(f"  {ts}  {f.name:<45}  {size_kb:6.0f} KB")
    if len(arquivos) > 20:
        print(f"  ... e mais {len(arquivos) - 20} arquivo(s).")


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    # Importa dependências do sistema
    from sistema_aps import (
        ROOT_INPUT, OUT_DIR, get_indicators, process_selected, desktop_files
    )

    out_dir = Path(args.saida) if args.saida else OUT_DIR
    in_dir = Path(args.entrada) if args.entrada else ROOT_INPUT

    # --historico
    if args.historico:
        _print_historico(out_dir)
        return

    # --listar-indicadores
    if args.listar_indicadores:
        print("\nIndicadores disponíveis:")
        for cfg, _ in get_indicators():
            origem = "plugin" if not cfg.official_like else "oficial"
            print(f"  {cfg.code:<6} {cfg.titulo.split('|')[0].strip()[:55]}  [{origem}]")
        return

    # --indicadores
    if not args.indicadores:
        parser.print_help()
        sys.exit(0)

    todos = [cfg.code for cfg, _ in get_indicators()]
    if args.indicadores == ["todos"]:
        selecionados = todos
    else:
        selecionados = [c.upper() for c in args.indicadores]
        invalidos = [c for c in selecionados if c not in todos]
        if invalidos:
            print(f"Indicadores inválidos: {', '.join(invalidos)}")
            print(f"Disponíveis: {', '.join(todos)}")
            sys.exit(1)

    print(f"\nAPS Suite — CLI")
    print(f"  Entrada:      {in_dir}")
    print(f"  Saída:        {out_dir}")
    print(f"  Indicadores:  {', '.join(selecionados)}")
    print(f"  Cache MD5:    {'não' if args.sem_cache else 'sim'}")
    print()

    # Valida que há arquivos
    try:
        files = desktop_files(in_dir)
    except Exception as exc:
        print(f"Erro ao listar arquivos: {exc}")
        sys.exit(1)

    if not files:
        print(f"Nenhum arquivo bruto encontrado em: {in_dir}")
        sys.exit(1)

    print(f"Arquivos encontrados: {len(files)}")
    for f in files:
        print(f"  • {f.name}")
    print()

    results = process_selected(
        selected_codes=selecionados,
        in_dir=in_dir,
        out_dir=out_dir,
        log=print,
        use_cache=not args.sem_cache,
    )

    # Sumário final
    ok = [r for r in results if r["status"] == "ok"]
    erros = [r for r in results if r["status"] == "erro"]
    nao_enc = [r for r in results if r["status"] == "não encontrado"]
    cache_hits = [r for r in ok if r.get("cache_hit")]

    print(f"\n{'='*60}")
    print(f"  ✔ Gerados:         {len(ok) - len(cache_hits)}")
    if cache_hits:
        print(f"  ↩ Cache (sem alt): {len(cache_hits)}")
    if erros:
        print(f"  ✘ Erros:           {len(erros)}")
        for r in erros:
            print(f"    {r['code']}: {(r.get('erro') or '').splitlines()[-1]}")
    if nao_enc:
        print(f"  ⚠ Não encontrados: {len(nao_enc)}")
        for r in nao_enc:
            print(f"    {r['code']}")

    sys.exit(1 if erros else 0)


if __name__ == "__main__":
    main()

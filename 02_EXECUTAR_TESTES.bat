@echo off
cd /d "%~dp0"
echo ============================================================
echo  APS Suite — Executar testes automatizados
echo ============================================================
py -m pytest tests\ -v
pause

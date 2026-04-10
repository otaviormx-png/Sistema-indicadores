@echo off
cd /d "%~dp0"
if exist dist\Central_APS.exe (
    start "" "dist\Central_APS.exe"
) else (
    echo EXE nao encontrado. Execute primeiro: 04_GERAR_EXE.bat
    pause
)

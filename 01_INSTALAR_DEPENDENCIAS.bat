@echo off
cd /d "%~dp0"
echo ============================================================
echo  APS Suite v3 — Instalacao de dependencias
echo ============================================================

where uv >nul 2>&1
if %errorlevel% == 0 (
    echo Usando uv...
    uv pip install -e ".[dev]"
) else (
    echo uv nao encontrado — usando pip padrao...
    py -m pip install --upgrade pip
    py -m pip install pandas openpyxl matplotlib pyinstaller pillow "tomli; python_version < '3.11'"
)

echo.
echo === Opcional: PDF ===
echo Para exportar PDF:  py -m pip install reportlab
echo.
echo === Opcional: Interface moderna ===
echo Para dark mode:     py -m pip install customtkinter
echo.
echo Instalacao base concluida!
pause

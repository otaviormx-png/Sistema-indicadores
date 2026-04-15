@echo off
setlocal
cd /d "%~dp0\src"

where pyinstaller >nul 2>&1
if errorlevel 1 (
  echo Instalando pyinstaller...
  pip install pyinstaller
)

echo Gerando EXE unico APS_LITE_3em1...
pyinstaller --noconfirm --onefile --windowed --name APS_LITE_3em1 --icon APS_Suite.ico aps_lite_3em1.py

if not exist dist\APS_LITE_3em1.exe (
  echo Falha ao gerar EXE.
  pause
  exit /b 1
)

cd /d "%~dp0"
copy /Y "src\dist\APS_LITE_3em1.exe" "dist\APS_LITE_3em1.exe" >nul

echo.
echo Concluido. EXE em:
echo %~dp0dist\APS_LITE_3em1.exe
pause

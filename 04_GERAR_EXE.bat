@echo off
cd /d "%~dp0"
echo ============================================================
echo  APS Suite v3 — Gerar executavel unico
echo ============================================================
rmdir /s /q build 2>nul
rmdir /s /q dist  2>nul

pyinstaller ^
  --onefile ^
  --windowed ^
  --clean ^
  --icon=APS_Suite.ico ^
  --name Central_APS ^
  --hidden-import=tkinter ^
  --hidden-import=tkinter.filedialog ^
  --hidden-import=tkinter.messagebox ^
  --hidden-import=matplotlib.backends.backend_tkagg ^
  --hidden-import=PIL._tkinter_finder ^
  --collect-submodules pandas ^
  --collect-submodules openpyxl ^
  --collect-submodules matplotlib ^
  --add-data "config.toml;." ^
  --add-data "plugins;plugins" ^
  aps_interface.py

echo.
echo EXE gerado em dist\Central_APS.exe
pause

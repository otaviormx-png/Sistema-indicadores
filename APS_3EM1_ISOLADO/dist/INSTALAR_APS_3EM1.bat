@echo off
setlocal
set "TARGET=%ProgramFiles%\APS_LITE_3EM1"
if not exist "%TARGET%" mkdir "%TARGET%"

copy /Y "%~dp0APS_LITE_3em1.exe" "%TARGET%\APS_LITE_3em1.exe" >nul
copy /Y "%~dp0config.toml" "%TARGET%\config.toml" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\APS LITE 3 em 1.lnk'); $s.TargetPath='%TARGET%\APS_LITE_3em1.exe'; $s.WorkingDirectory='%TARGET%'; $s.Save()"

echo.
echo Instalacao concluida em: %TARGET%
echo Atalho criado na area de trabalho.
pause

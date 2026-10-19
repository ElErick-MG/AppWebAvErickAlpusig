@echo off
REM ===========================================================================
REM  PoliMusic - refresca la copia publicada en IIS sin tocar la configuracion.
REM  Uso: doble clic (pide elevacion). Copia los archivos a
REM  C:\inetpub\wwwroot\webapppolimusicv1_asp
REM ===========================================================================
setlocal

net session >nul 2>&1
if not "%errorlevel%"=="0" (
    echo Solicitando privilegios de administrador...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1" ^
    -Destination "C:\inetpub\wwwroot\webapppolimusicv1_asp" ^
    -SkipIisConfig

echo.
pause

@echo off
REM ===========================================================================
REM  PoliMusic - double-click launcher that requests elevation and runs
REM  deploy.ps1 against C:\inetpub\wwwroot
REM ===========================================================================
setlocal

net session >nul 2>&1
if not "%errorlevel%"=="0" (
    echo Requesting administrator privileges...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1" %*

echo.
pause

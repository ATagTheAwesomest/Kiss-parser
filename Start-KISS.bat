@echo off
rem KISS launcher. Double-click this, or drag a job folder onto it.
rem Runs the script for this window only; nothing on the PC is changed.
setlocal
set "KISSROOT=%~dp0"
if not "%~1"=="" (
    set "WORK=%~1"
) else (
    set "WORK=%KISSROOT%jobs"
)
if not exist "%WORK%" mkdir "%WORK%"
cd /d "%WORK%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%KISSROOT%tools\Parse-KSS.ps1"
echo.
pause

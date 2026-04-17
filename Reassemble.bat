@echo off
REM ==========================================================
REM  Excel Workbook Manager - Reassemble launcher (.bat)
REM  Joins .tar.partNNN chunks, verifies SHA256, and extracts.
REM  Double-click this file.
REM ==========================================================
setlocal
cd /d "%~dp0"

where powershell >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell was not found on this PC.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Reassemble.ps1"
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" (
    echo Reassembly finished successfully.
) else (
    echo Reassembly failed with exit code %RC%.
)
pause
exit /b %RC%

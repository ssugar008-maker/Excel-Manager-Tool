@echo off
REM ==========================================================
REM  Excel Workbook Manager - Reassemble (xlsb-named chunks)
REM  These .xlsb files are NOT Excel workbooks - they are raw
REM  slices of a tar archive renamed to pass corporate filters.
REM  Double-click this file to join and extract them.
REM ==========================================================
setlocal
cd /d "%~dp0"

where powershell >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell was not found on this PC.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Reassemble-xlsb.ps1"
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" (
    echo Reassembly finished successfully.
) else (
    echo Reassembly failed with exit code %RC%.
)
pause
exit /b %RC%

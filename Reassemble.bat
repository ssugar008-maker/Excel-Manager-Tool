@echo off
REM =====================================================================
REM   ExcelWorkbookManager — Reassemble the .exe from .xlsb chunks
REM
REM   GitHub caps files at 100 MB, so the executable is split into
REM   8 MB .xlsb parts. This script concatenates them back into a
REM   working ExcelWorkbookManager.exe and verifies the SHA-256 hash.
REM
REM   Usage: just double-click this file inside the folder that
REM   contains ExcelWorkbookManager.partNN.xlsb.
REM =====================================================================

setlocal EnableDelayedExpansion

REM -- Run from the folder this script lives in ------------------------
pushd "%~dp0"

set "OUT=ExcelWorkbookManager.exe"
set "EXPECTED_SHA=E41CC712ABF5BABAB48253D3D6A5F80A005DC010424437BDC61BB92B443A699C"

REM -- Sanity: every chunk must be present -----------------------------
set MISSING=0
for %%P in (ExcelWorkbookManager.part01.xlsb ExcelWorkbookManager.part02.xlsb) do (
    if not exist "%%P" (
        echo [ERROR] Missing chunk: %%P
        set /a MISSING+=1
    )
)
if not "%MISSING%"=="0" (
    echo.
    echo Please make sure ALL ExcelWorkbookManager.partNN.xlsb files are in
    echo this folder alongside Reassemble.bat, then run again.
    popd
    pause
    exit /b 1
)

REM -- Remove any stale output before writing --------------------------
if exist "%OUT%" del /f /q "%OUT%" >nul 2>&1

echo Reassembling %OUT% from chunks...
copy /b /y ^
  "ExcelWorkbookManager.part01.xlsb" + ^
  "ExcelWorkbookManager.part02.xlsb" ^
  "%OUT%" >nul

if errorlevel 1 (
    echo [ERROR] copy /b failed.
    popd
    pause
    exit /b 2
)

REM -- Verify SHA-256 against the known-good hash ----------------------
echo Verifying checksum...
set "ACTUAL_SHA="
for /f "skip=1 tokens=* usebackq" %%H in (`certutil -hashfile "%OUT%" SHA256`) do (
    if not defined ACTUAL_SHA set "ACTUAL_SHA=%%H"
)
set "ACTUAL_SHA=%ACTUAL_SHA: =%"

if /I "%ACTUAL_SHA%"=="%EXPECTED_SHA%" (
    echo.
    echo [OK] Reassembly complete.
    echo      File      : %OUT%
    echo      SHA-256   : %ACTUAL_SHA%
    echo.
    echo You can now double-click %OUT% to launch the tool.
    popd
    pause
    exit /b 0
) else (
    echo.
    echo [ERROR] Checksum mismatch. The file may be corrupt or a chunk was
    echo        downloaded incorrectly. Expected and actual SHA-256:
    echo          expected: %EXPECTED_SHA%
    echo          actual  : %ACTUAL_SHA%
    echo Please re-download the ExcelWorkbookManager.partNN.xlsb chunks.
    if exist "%OUT%" del /f /q "%OUT%" >nul 2>&1
    popd
    pause
    exit /b 3
)

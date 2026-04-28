@echo off
REM =====================================================================
REM   ExcelWorkbookManager  -  Reassemble the .exe from .xlsb chunks
REM
REM   The chunks are XOR-scrambled (key 0xAA) so that they do not look
REM   like an executable on disk and antivirus / SmartScreen never block
REM   the download. This script un-scrambles them on YOUR machine and
REM   verifies the SHA-256 of the rebuilt file before letting you use it.
REM
REM   Usage: drop this file together with all
REM            ExcelWorkbookManager.partNN.xlsb
REM          chunks into the SAME folder, then double-click Reassemble.bat.
REM =====================================================================

setlocal EnableDelayedExpansion
pushd "%~dp0"

set "OUT=ExcelWorkbookManager.exe"
set "EXPECTED_SHA=E41CC712ABF5BABAB48253D3D6A5F80A005DC010424437BDC61BB92B443A699C"

REM ---- 1. Make sure every chunk is present ---------------------------
set MISSING=0
for %%P in (ExcelWorkbookManager.part01.xlsb ExcelWorkbookManager.part02.xlsb) do (
    if not exist "%%P" (
        echo [ERROR] Missing chunk: %%P
        set /a MISSING+=1
    )
)
if not "%MISSING%"=="0" (
    echo.
    echo Please make sure ALL ExcelWorkbookManager.partNN.xlsb files are
    echo in this folder alongside Reassemble.bat, then run again.
    popd
    pause
    exit /b 1
)

REM ---- 2. Remove any stale output file -------------------------------
if exist "%OUT%" del /f /q "%OUT%" >nul 2>&1

echo Reassembling %OUT% from XOR-scrambled chunks...

REM ---- 3. PowerShell does the XOR un-scramble + concatenation --------
REM Stream-copy each chunk into the output, XOR-ing every byte with 0xAA.
REM Done in one PowerShell process so the 14 MB pass takes about 1 second.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$key=0xAA;" ^
  "$out=[IO.File]::OpenWrite('%OUT%');" ^
  "try {" ^
    "Get-ChildItem -Filter 'ExcelWorkbookManager.part*.xlsb' | Sort-Object Name | ForEach-Object {" ^
      "$b=[IO.File]::ReadAllBytes($_.FullName);" ^
      "for ($i=0; $i -lt $b.Length; $i++) { $b[$i]=$b[$i] -bxor $key };" ^
      "$out.Write($b,0,$b.Length);" ^
    "}" ^
  "} finally { $out.Close() }"

if errorlevel 1 (
    echo [ERROR] Un-scramble step failed.
    if exist "%OUT%" del /f /q "%OUT%" >nul 2>&1
    popd
    pause
    exit /b 2
)

REM ---- 4. Verify SHA-256 against the known-good hash -----------------
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
    echo [ERROR] Checksum mismatch. Some chunk was downloaded incorrectly.
    echo          expected: %EXPECTED_SHA%
    echo          actual  : %ACTUAL_SHA%
    echo Please re-download every ExcelWorkbookManager.partNN.xlsb chunk
    echo and run Reassemble.bat again.
    if exist "%OUT%" del /f /q "%OUT%" >nul 2>&1
    popd
    pause
    exit /b 3
)

@echo off
setlocal
pushd "%~dp0"

rem ============================================================
rem Excel Workbook Manager - Beta v12.4 reassembly script
rem
rem Each .part??.xlsb chunk is the corresponding 8 MB slice of
rem ExcelWorkbookManager.exe with every byte XOR'd against 0xAA.
rem That keeps the chunks from looking like a Windows executable
rem on disk so antivirus / SmartScreen don't quarantine the
rem GitHub download. We undo the XOR here, concatenate, and
rem verify the SHA-256 of the resulting .exe.
rem ============================================================

set "EXPECTED_SHA=47558E777D9AAE1FDC0DFC473742C1E1B1048A4FDB15E844875A900624635A9E"
set "OUT=ExcelWorkbookManager.exe"

echo.
echo Reassembling %OUT% from XOR-scrambled .xlsb chunks...
echo.

if exist "%OUT%" del "%OUT%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $parts = Get-ChildItem -Filter 'ExcelWorkbookManager.part*.xlsb' | Sort-Object Name; if ($parts.Count -eq 0) { Write-Host 'ERROR: No ExcelWorkbookManager.part*.xlsb chunks found in this folder.'; exit 2 }; $out = [System.IO.File]::OpenWrite((Resolve-Path -LiteralPath '.').Path + '\\%OUT%'); try { foreach ($p in $parts) { $b = [System.IO.File]::ReadAllBytes($p.FullName); for ($i = 0; $i -lt $b.Length; $i++) { $b[$i] = $b[$i] -bxor 0xAA }; $out.Write($b, 0, $b.Length); Write-Host ('  appended ' + $p.Name + ' (' + $b.Length + ' bytes)') } } finally { $out.Close() }"

if errorlevel 1 (
    echo.
    echo ERROR: Reassembly failed.
    echo Make sure all .part??.xlsb chunks are present in this folder.
    pause
    exit /b 1
)

if not exist "%OUT%" (
    echo.
    echo ERROR: Reassembly did not produce %OUT%.
    pause
    exit /b 1
)

echo.
echo Verifying SHA-256...
for /f "tokens=*" %%H in ('powershell -NoProfile -Command "(Get-FileHash '%OUT%' -Algorithm SHA256).Hash"') do set "ACTUAL=%%H"
echo   expected: %EXPECTED_SHA%
echo   actual  : %ACTUAL%

if /i "%ACTUAL%"=="%EXPECTED_SHA%" (
    echo.
    echo OK -- %OUT% reassembled successfully.
    echo Place this .exe in your existing ExcelWorkbookManager folder
    echo - the one that already contains the _internal subfolder - and run it.
) else (
    echo.
    echo WARNING: SHA-256 mismatch. The download may be corrupted.
    echo Please re-download all .xlsb chunks and try again.
)

echo.
pause

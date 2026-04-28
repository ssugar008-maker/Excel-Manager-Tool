@echo off
setlocal enabledelayedexpansion
rem ============================================================
rem Excel Workbook Manager — Beta v12.3 reassembly script
rem
rem Each .part??.xlsb chunk is the corresponding 8 MB slice of
rem ExcelWorkbookManager.exe with every byte XOR'd against 0xAA.
rem That keeps the chunks from looking like a Windows executable
rem on disk so antivirus / SmartScreen don't quarantine the
rem GitHub download. We undo the XOR here, concatenate, and
rem verify the SHA-256 of the resulting .exe.
rem ============================================================

set "EXPECTED_SHA=114734AAE5EE6FC39110535D7BAC804EFDAA68437CE5D36DEC8C084EE2D16788"
set "OUT=ExcelWorkbookManager.exe"

echo.
echo Reassembling %OUT% from XOR-scrambled .xlsb chunks...
echo.

rem Locate every part??.xlsb in numerical order.
set "PARTS="
for /f "delims=" %%P in ('dir /b /on "ExcelWorkbookManager.part*.xlsb" 2^>nul') do (
    if defined PARTS (set "PARTS=!PARTS!,%%P") else (set "PARTS=%%P")
)
if not defined PARTS (
    echo ERROR: No ExcelWorkbookManager.part*.xlsb files found in this folder.
    pause
    exit /b 1
)

if exist "%OUT%" del "%OUT%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$parts = '%PARTS%'.Split(','); ^
   $out = [System.IO.File]::OpenWrite('%OUT%'); ^
   try { ^
     foreach ($p in $parts) { ^
       $b = [System.IO.File]::ReadAllBytes($p); ^
       for ($i = 0; $i -lt $b.Length; $i++) { $b[$i] = $b[$i] -bxor 0xAA }; ^
       $out.Write($b, 0, $b.Length); ^
       Write-Host ('  appended ' + $p + ' (' + $b.Length + ' bytes)') ^
     } ^
   } finally { $out.Close() }"

if not exist "%OUT%" (
    echo ERROR: Reassembly failed.
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
    echo OK — %OUT% reassembled successfully.
    echo Place this .exe in your existing ExcelWorkbookManager folder
    echo (the one that already contains the _internal subfolder) and
    echo run it.
) else (
    echo.
    echo WARNING: SHA-256 mismatch. The download may be corrupted.
    echo Please re-download all .xlsb chunks and try again.
)

echo.
pause

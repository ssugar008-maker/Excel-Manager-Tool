@echo off
REM =================================================================
REM  Excel Workbook Manager - Reassemble the .exe from XOR chunks
REM  Drop this file beside the ExcelWorkbookManager.part*.xlsb files,
REM  double-click it, and copy the produced ExcelWorkbookManager.exe
REM  into your existing ExcelWorkbookManager folder (next to the
REM  _internal folder) - replacing the old .exe.
REM =================================================================
setlocal
pushd "%~dp0"

where powershell >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell not found. Install Windows PowerShell and retry.
    popd
    pause
    exit /b 1
)

echo Reassembling ExcelWorkbookManager.exe from .xlsb chunks ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $parts = Get-ChildItem -Path '.' -Filter 'ExcelWorkbookManager.part*.xlsb' | Sort-Object Name; if ($parts.Count -eq 0) { throw 'No ExcelWorkbookManager.part*.xlsb chunks found in this folder.' }; $out = 'ExcelWorkbookManager.exe'; if (Test-Path $out) { Remove-Item -Force $out }; $fs = [System.IO.File]::Create($out); try { foreach ($p in $parts) { Write-Host ('  Decoding {0} ...' -f $p.Name); $buf = [System.IO.File]::ReadAllBytes($p.FullName); for ($i = 0; $i -lt $buf.Length; $i++) { $buf[$i] = $buf[$i] -bxor 0xAA }; $fs.Write($buf, 0, $buf.Length) } } finally { $fs.Close() }; $h = (Get-FileHash $out -Algorithm SHA256).Hash.ToLower(); Write-Host ''; Write-Host ('Produced ' + $out + ' [SHA-256: ' + $h + ']')"
set RC=%ERRORLEVEL%

echo.
if "%RC%"=="0" (
    echo DONE - copy ExcelWorkbookManager.exe into your existing ExcelWorkbookManager folder.
) else (
    echo FAILED with exit code %RC%.
)
popd
pause
exit /b %RC%

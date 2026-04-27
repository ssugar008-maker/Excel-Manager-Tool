@echo off
REM reassemble.bat — Reassemble ExcelWorkbookManager from .xlsb chunks
REM Run from the folder that contains this script AND all .xlsb chunk files.

setlocal EnableDelayedExpansion
set OUTTAR=ExcelWorkbookManager.tar
set OUTDIR=ExcelWorkbookManager

echo Reassembling chunks into %OUTTAR% ...

if exist "%OUTTAR%" del "%OUTTAR%"

for /f "tokens=*" %%f in ('dir /b /on "*.xlsb" 2^>nul') do (
    echo   Appending %%f ...
    copy /b "%OUTTAR%"+"%%f" "%OUTTAR%.tmp" >nul 2>&1 || copy /b "%%f" "%OUTTAR%.tmp" >nul
    move /y "%OUTTAR%.tmp" "%OUTTAR%" >nul
)

echo Extracting %OUTTAR% ...
if exist "%OUTDIR%" rmdir /s /q "%OUTDIR%"
tar -xf "%OUTTAR%"

echo.
echo Done!  Run:  %OUTDIR%\ExcelWorkbookManager.exe
pause

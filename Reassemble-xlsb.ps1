<#
.SYNOPSIS
    Joins ExcelWorkbookManager.partNNN.xlsb chunks and extracts the app.

.DESCRIPTION
    IMPORTANT: The .xlsb files in this folder are NOT Excel workbooks.
    They are raw binary slices of a tar archive, renamed to .xlsb so they
    slip past corporate attachment/AV filters that block unusual types.
    Do NOT open them in Excel - run this script instead.

    This script:
        1. Concatenates ExcelWorkbookManager.part001.xlsb .. .partNNN.xlsb
           into a single ExcelWorkbookManager.tar (in this folder).
        2. Verifies SHA256 against SHA256SUMS-xlsb.txt (if present).
        3. Extracts the tar using Windows' built-in tar.exe
           (Windows 10 1803+ / Windows 11).
    Result: an ExcelWorkbookManager\ folder next to this script.

.USAGE
    Right-click Reassemble-xlsb.ps1 -> "Run with PowerShell"
        OR
    powershell -ExecutionPolicy Bypass -File .\Reassemble-xlsb.ps1
#>

[CmdletBinding()]
param(
    [string]$PartsPattern = "ExcelWorkbookManager.part*.xlsb",
    [string]$OutputTar    = "ExcelWorkbookManager.tar",
    [string]$ExtractTo    = ".",
    [switch]$KeepTar,
    [switch]$SkipVerify
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Excel Workbook Manager - Reassemble (xlsb chunks) " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

$partFiles = Get-ChildItem -File -Filter $PartsPattern |
    Where-Object { $_.Name -match '\.part\d{3}\.xlsb$' } |
    Sort-Object Name

if ($partFiles.Count -eq 0) {
    Write-Error "No chunks found matching '$PartsPattern' in $PSScriptRoot"
    exit 1
}

Write-Host ("Found {0} chunk(s):" -f $partFiles.Count)
$partFiles | ForEach-Object {
    Write-Host ("  {0}  ({1:N2} MB)" -f $_.Name, ($_.Length / 1MB))
}
Write-Host ""

# --- Per-chunk SHA256 verify ---------------------------------------------
$sumsFile = "SHA256SUMS-xlsb.txt"
if (-not $SkipVerify -and (Test-Path $sumsFile)) {
    Write-Host "Verifying chunk SHA256 values against $sumsFile ..."
    $expectedMap = @{}
    Get-Content $sumsFile | ForEach-Object {
        if ($_ -match '^\s*([0-9A-Fa-f]{64})\s+(\S+)\s*$') {
            $expectedMap[$Matches[2]] = $Matches[1].ToUpper()
        }
    }
    $anyFailed = $false
    foreach ($p in $partFiles) {
        if ($expectedMap.ContainsKey($p.Name)) {
            $actual = (Get-FileHash $p.FullName -Algorithm SHA256).Hash.ToUpper()
            if ($actual -eq $expectedMap[$p.Name]) {
                Write-Host ("  OK    {0}" -f $p.Name) -ForegroundColor Green
            } else {
                Write-Host ("  FAIL  {0}" -f $p.Name) -ForegroundColor Red
                Write-Host ("        expected {0}" -f $expectedMap[$p.Name])
                Write-Host ("        actual   {0}" -f $actual)
                $anyFailed = $true
            }
        }
    }
    if ($anyFailed) {
        Write-Error "One or more chunks failed SHA256 verification. Re-download them."
        exit 2
    }
    Write-Host ""
}

# --- Join ----------------------------------------------------------------
if (Test-Path $OutputTar) {
    Write-Host "Removing existing $OutputTar ..."
    Remove-Item $OutputTar -Force
}

Write-Host "Joining chunks into $OutputTar ..."
$out = [System.IO.File]::Create((Join-Path $PSScriptRoot $OutputTar))
try {
    $buffer = New-Object byte[] 1048576
    foreach ($p in $partFiles) {
        Write-Host ("  + {0}" -f $p.Name)
        $in = [System.IO.File]::OpenRead($p.FullName)
        try {
            while (($read = $in.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $out.Write($buffer, 0, $read)
            }
        } finally { $in.Close() }
    }
} finally { $out.Close() }

$joinedSize = (Get-Item $OutputTar).Length
Write-Host ("Joined archive: {0:N2} MB" -f ($joinedSize / 1MB))
Write-Host ""

# --- Verify joined tar ---------------------------------------------------
if (-not $SkipVerify -and (Test-Path $sumsFile)) {
    $expected = $null
    Get-Content $sumsFile | ForEach-Object {
        if ($_ -match '^\s*([0-9A-Fa-f]{64})\s+ExcelWorkbookManager\.tar\s*$') {
            $expected = $Matches[1].ToUpper()
        }
    }
    if ($expected) {
        $actual = (Get-FileHash $OutputTar -Algorithm SHA256).Hash.ToUpper()
        if ($actual -eq $expected) {
            Write-Host "Joined tar SHA256 OK." -ForegroundColor Green
        } else {
            Write-Host "Joined tar SHA256 MISMATCH!" -ForegroundColor Red
            Write-Host "  expected $expected"
            Write-Host "  actual   $actual"
            Write-Error "Checksum failed."
            exit 3
        }
        Write-Host ""
    }
}

# --- Extract -------------------------------------------------------------
$tarExe = Get-Command tar.exe -ErrorAction SilentlyContinue
if (-not $tarExe) {
    Write-Host "tar.exe not found. Joined archive at: $OutputTar" -ForegroundColor Yellow
    Write-Host "Update Windows (1803+) or use Git for Windows' tar, then run: tar -xf $OutputTar"
    exit 0
}

Write-Host "Extracting $OutputTar with tar.exe ..."
& tar.exe -xf $OutputTar -C $ExtractTo
if ($LASTEXITCODE -ne 0) {
    Write-Error "tar extraction failed with exit code $LASTEXITCODE"
    exit 4
}

if (-not $KeepTar) {
    Remove-Item $OutputTar -Force
    Write-Host "Removed intermediate $OutputTar (use -KeepTar to keep it)."
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " DONE" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
$exePath = Join-Path $PSScriptRoot "ExcelWorkbookManager\ExcelWorkbookManager.exe"
if (Test-Path $exePath) {
    Write-Host "Launch the app:"
    Write-Host "  $exePath"
} else {
    Write-Host "Extracted under: $PSScriptRoot\ExcelWorkbookManager\"
    Write-Host "Double-click ExcelWorkbookManager.exe inside that folder."
}
Write-Host ""
Write-Host "NOTE: Microsoft Excel must be installed on this PC for the tool to work."

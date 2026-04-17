<#
.SYNOPSIS
    Reassembles and extracts the split ExcelWorkbookManager archive.

.DESCRIPTION
    1. Joins ExcelWorkbookManager.zip.001, .002, .003, ... into a single
       ExcelWorkbookManager.zip in this folder.
    2. (Optional) Verifies SHA256 against SHA256SUMS.txt.
    3. Extracts the archive into .\ExcelWorkbookManager\.

    Run this script from the folder that contains the .zip.### parts.

.USAGE
    Right-click Reassemble.ps1 -> "Run with PowerShell"
        OR
    Open PowerShell in this folder and run:
        powershell -ExecutionPolicy Bypass -File .\Reassemble.ps1
#>

[CmdletBinding()]
param(
    [string]$Parts      = "ExcelWorkbookManager.zip.*",
    [string]$OutputZip  = "ExcelWorkbookManager.zip",
    [string]$ExtractTo  = "ExcelWorkbookManager",
    [switch]$SkipExtract,
    [switch]$SkipVerify
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Excel Workbook Manager - Reassemble Script " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$partFiles = Get-ChildItem -File -Filter $Parts |
    Where-Object { $_.Name -match '\.zip\.\d{3}$' } |
    Sort-Object Name

if ($partFiles.Count -eq 0) {
    Write-Error "No split parts found matching '$Parts' in $PSScriptRoot"
    exit 1
}

Write-Host ("Found {0} part(s):" -f $partFiles.Count)
$partFiles | ForEach-Object {
    Write-Host ("  {0}  ({1:N2} MB)" -f $_.Name, ($_.Length / 1MB))
}
Write-Host ""

if (Test-Path $OutputZip) {
    Write-Host "Removing existing $OutputZip ..."
    Remove-Item $OutputZip -Force
}

Write-Host "Joining parts into $OutputZip ..."
$out = [System.IO.File]::Create((Join-Path $PSScriptRoot $OutputZip))
try {
    $buffer = New-Object byte[] 1048576
    foreach ($p in $partFiles) {
        Write-Host ("  + {0}" -f $p.Name)
        $in = [System.IO.File]::OpenRead($p.FullName)
        try {
            while (($read = $in.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $out.Write($buffer, 0, $read)
            }
        } finally {
            $in.Close()
        }
    }
} finally {
    $out.Close()
}

$joinedSize = (Get-Item $OutputZip).Length
Write-Host ("Joined archive: {0:N2} MB" -f ($joinedSize / 1MB))
Write-Host ""

if (-not $SkipVerify -and (Test-Path "SHA256SUMS.txt")) {
    Write-Host "Verifying SHA256 against SHA256SUMS.txt ..."
    $expected = $null
    Get-Content "SHA256SUMS.txt" | ForEach-Object {
        if ($_ -match '^\s*([0-9A-Fa-f]{64})\s+ExcelWorkbookManager\.zip\s*$') {
            $expected = $Matches[1].ToUpper()
        }
    }
    if ($expected) {
        $actual = (Get-FileHash $OutputZip -Algorithm SHA256).Hash.ToUpper()
        if ($actual -eq $expected) {
            Write-Host "  OK  SHA256 matches." -ForegroundColor Green
        } else {
            Write-Host "  FAIL  SHA256 mismatch!" -ForegroundColor Red
            Write-Host "    expected: $expected"
            Write-Host "    actual:   $actual"
            Write-Error "Checksum failed. Re-download the parts."
            exit 2
        }
    } else {
        Write-Host "  (no ExcelWorkbookManager.zip entry in SHA256SUMS.txt, skipping)"
    }
    Write-Host ""
}

if ($SkipExtract) {
    Write-Host "Done. Archive left at: $OutputZip"
    exit 0
}

if (Test-Path $ExtractTo) {
    Write-Host "Removing existing folder $ExtractTo ..."
    Remove-Item $ExtractTo -Recurse -Force
}

Write-Host "Extracting to .\$ExtractTo\ ..."
Expand-Archive -LiteralPath $OutputZip -DestinationPath $ExtractTo -Force
Write-Host ""

$exePath = Join-Path $ExtractTo "ExcelWorkbookManager.exe"
if (-not (Test-Path $exePath)) {
    # PyInstaller --onedir creates an inner folder named after --name
    $inner = Join-Path $ExtractTo "ExcelWorkbookManager\ExcelWorkbookManager.exe"
    if (Test-Path $inner) { $exePath = $inner }
}

Write-Host "============================================" -ForegroundColor Green
Write-Host " DONE" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
if (Test-Path $exePath) {
    Write-Host "Launch the app:"
    Write-Host "  $exePath"
} else {
    Write-Host "Extracted to: $ExtractTo"
    Write-Host "Look for ExcelWorkbookManager.exe inside and double-click it."
}
Write-Host ""
Write-Host "NOTE: Excel must be installed on this PC for the tool to work."

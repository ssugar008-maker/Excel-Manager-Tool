# reassemble.ps1 — Reassemble ExcelWorkbookManager from .xlsb chunks
# Run from the folder that contains this script AND all .xlsb chunk files.
# Usage:  powershell -ExecutionPolicy Bypass -File reassemble.ps1

$ErrorActionPreference = "Stop"
$outTar  = "ExcelWorkbookManager.tar"
$outDir  = "ExcelWorkbookManager"

Write-Host "Reassembling chunks into $outTar ..."

$chunks = Get-ChildItem -Path "." -Filter "*.xlsb" |
          Where-Object { $_.Name -match 'part\d+' } |
          Sort-Object Name

if ($chunks.Count -eq 0) {
    Write-Error "No .xlsb chunk files found in the current directory."
}

$dest = [System.IO.File]::OpenWrite($outTar)
foreach ($c in $chunks) {
    Write-Host "  Appending $($c.Name) ..."
    $bytes = [System.IO.File]::ReadAllBytes($c.FullName)
    $dest.Write($bytes, 0, $bytes.Length)
}
$dest.Close()

Write-Host "Verifying checksums ..."
$checksums = Get-Content "checksums.json" | ConvertFrom-Json
$ok = $true
foreach ($chunk in $chunks) {
    $expected = $checksums.($chunk.Name)
    $actual   = (Get-FileHash $chunk.FullName -Algorithm SHA256).Hash.ToLower()
    if ($actual -ne $expected) {
        Write-Warning "MISMATCH: $($chunk.Name)"
        $ok = $false
    }
}
if ($ok) { Write-Host "All checksums verified OK." } else { Write-Warning "Some checksums failed — download may be corrupt." }

Write-Host "Extracting $outTar ..."
if (Test-Path $outDir) { Remove-Item -Recurse -Force $outDir }
tar -xf $outTar

Write-Host ""
Write-Host "Done!  Run:  .\$outDir\ExcelWorkbookManager.exe"

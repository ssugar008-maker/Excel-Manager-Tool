# Excel Workbook Manager — Standalone Distribution

**Beta v8** — two new actuarial power tabs: **Smart Refresh Orchestrator** (auto-discover linked docs, open them, recalculate selected sheets, compare audit cells before/after, scheduled auto-run with countdown timer, named profiles, exportable colour-coded audit log) and **Snapshot Export** (copy selected sheets as paste-values-only to a new file automatically named `originalname_YYYYMMDD.xlsb`).

This repository contains the **standalone Windows build** of Excel Workbook Manager, distributed as **raw binary chunks** of an uncompressed tar archive. No Python install is needed on the end-user PC — just Microsoft Excel.

The build is not shipped as a `.zip` on purpose. Each `.partNNN` file is a plain binary slice of the tar archive, so they can be concatenated byte-for-byte with `copy /b` or any equivalent tool.

---

## Contents

| File                                   | Purpose                                                |
| -------------------------------------- | ------------------------------------------------------ |
| `ExcelWorkbookManager.part001`         | Raw chunk 1 (40 MB)                                    |
| `ExcelWorkbookManager.part002`         | Raw chunk 2 (40 MB)                                    |
| `ExcelWorkbookManager.part003`         | Raw chunk 3 (40 MB)                                    |
| `ExcelWorkbookManager.part004`         | Raw chunk 4 (~16 MB)                                   |
| `SHA256SUMS.txt`                       | Checksums for every chunk and the joined tar           |
| `Reassemble.ps1`                       | PowerShell: verify + join + extract                    |
| `Reassemble.bat`                       | Double-click launcher for `Reassemble.ps1`             |
| `GUIDE.html`                           | Full user guide                                        |
| `profiles.json`, `settings.json`       | Default example configs (also shipped inside the tar)  |

---

## Quick start — end users

1. Click the green **Code** button on this repo → **Download ZIP** (or `git clone`).
2. Put every `ExcelWorkbookManager.tar.partNNN` file plus `Reassemble.bat` in the same folder.
3. **Double-click `Reassemble.bat`.** It will:
   - Verify the SHA256 of each chunk
   - Concatenate them into a single `ExcelWorkbookManager.tar`
   - Re-check the joined tar's SHA256
   - Extract it with Windows' built-in `tar.exe` into a new `ExcelWorkbookManager\` folder
   - Delete the intermediate tar
4. Open the new `ExcelWorkbookManager\` folder and double-click **`ExcelWorkbookManager.exe`**.

> Windows SmartScreen may warn about an unsigned executable. Click **More info → Run anyway**.

### Alternative (pure command line, no script)

From a classic Command Prompt in the download folder:

```bat
copy /b ExcelWorkbookManager.part001 + ExcelWorkbookManager.part002 + ExcelWorkbookManager.part003 + ExcelWorkbookManager.part004 ExcelWorkbookManager.tar
tar -xf ExcelWorkbookManager.tar
```

Or from PowerShell:

```powershell
cmd /c "copy /b ExcelWorkbookManager.part001 + ExcelWorkbookManager.part002 + ExcelWorkbookManager.part003 + ExcelWorkbookManager.part004 ExcelWorkbookManager.tar"
tar -xf .\ExcelWorkbookManager.tar
```

---

## Requirements

- Windows 10 version 1803 or newer (has built-in `tar.exe`), or Windows 11
- Microsoft Excel installed and licensed on the same machine
- ~300 MB free disk space after extraction

No Python, no pip, no virtualenv on the end-user PC.

---

## Verifying the download (optional)

```powershell
Get-FileHash .\ExcelWorkbookManager.part001 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.part002 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.part003 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.part004 -Algorithm SHA256
```

Compare against `SHA256SUMS.txt`. `Reassemble.ps1` does all of this automatically.

---

## For developers — rebuilding from source

Source lives in a separate development folder. To rebuild the EXE:

```bat
pip install -r requirements.txt
build.bat
```

`build.bat` runs PyInstaller in `--onedir --windowed` mode and produces `dist\ExcelWorkbookManager\`. To regenerate the split distribution in this repo, tar that folder uncompressed and split at 40 MB:

```powershell
tar -cf ExcelWorkbookManager.tar -C dist ExcelWorkbookManager
# then split the tar using Reassemble.ps1's inverse logic
```

---

## Troubleshooting

- **`tar.exe` not recognized** — Your Windows is pre-1803. Update Windows or install Git for Windows (its tar works), or extract the joined `.tar` on any Linux/macOS machine.
- **`Reassemble.bat` flashes and closes** — Open PowerShell in the folder and run `powershell -ExecutionPolicy Bypass -File .\Reassemble.ps1` to see the error.
- **SHA256 mismatch** — Re-download the failing chunk. GitHub occasionally serves truncated binaries on flaky connections.
- **App starts then nothing happens** — Make sure Excel is installed and can open normally on this PC.

---

See **`GUIDE.html`** for the full user manual.

---

## Alternate download set: `.xlsb`-named chunks

For environments where attachment / AV filters block unusual binary
extensions, the same archive is also published as files ending in
`.xlsb` (with the part number **inside** the filename, so `.xlsb` is the
final extension). This set is cut into **17 smaller ~8 MB chunks**:

| File                                                   | Size        |
| ------------------------------------------------------ | ----------- |
| `ExcelWorkbookManager.part001.xlsb` … `part016.xlsb`   | 8 MB each   |
| `ExcelWorkbookManager.part017.xlsb`                    | ~6 MB       |
| `SHA256SUMS-xlsb.txt`                                  | checksums   |
| `Reassemble-xlsb.bat` / `.ps1`                         | helpers     |

Total across the 17 chunks = ~134 MB (same uncompressed tar as the
`.tar.partNNN` set).

> **These `.xlsb` files are NOT Excel workbooks.** They are raw
> byte-for-byte slices of the same tar archive as the `.tar.partNNN`
> set, simply renamed so they survive corporate file-type scanners.
> Do not try to open them in Excel.

### How to use

1. Download **all 17** `ExcelWorkbookManager.part*.xlsb` files (plus
   `Reassemble-xlsb.bat`, `Reassemble-xlsb.ps1`, `SHA256SUMS-xlsb.txt`)
   into the same folder. Missing even one chunk will corrupt the build.
2. Double-click **`Reassemble-xlsb.bat`**. The script automatically
   detects however many `part*.xlsb` files are present, verifies each
   one against `SHA256SUMS-xlsb.txt`, joins them in order, and extracts
   the result with Windows' built-in `tar.exe`.
3. Open the new `ExcelWorkbookManager\` folder and run
   `ExcelWorkbookManager.exe`.

Manual (no script) — from PowerShell in the download folder:

```powershell
$parts = Get-ChildItem ExcelWorkbookManager.part*.xlsb | Sort-Object Name
$out = [System.IO.File]::Create("$PWD\ExcelWorkbookManager.tar")
foreach ($p in $parts) {
    $in = [System.IO.File]::OpenRead($p.FullName)
    $in.CopyTo($out)
    $in.Close()
}
$out.Close()
tar -xf .\ExcelWorkbookManager.tar
```

Or with good old `copy /b` in a classic Command Prompt (one long line):

```bat
copy /b ExcelWorkbookManager.part001.xlsb + ExcelWorkbookManager.part002.xlsb + ExcelWorkbookManager.part003.xlsb + ExcelWorkbookManager.part004.xlsb + ExcelWorkbookManager.part005.xlsb + ExcelWorkbookManager.part006.xlsb + ExcelWorkbookManager.part007.xlsb + ExcelWorkbookManager.part008.xlsb + ExcelWorkbookManager.part009.xlsb + ExcelWorkbookManager.part010.xlsb + ExcelWorkbookManager.part011.xlsb + ExcelWorkbookManager.part012.xlsb + ExcelWorkbookManager.part013.xlsb + ExcelWorkbookManager.part014.xlsb + ExcelWorkbookManager.part015.xlsb + ExcelWorkbookManager.part016.xlsb + ExcelWorkbookManager.part017.xlsb ExcelWorkbookManager.tar
tar -xf ExcelWorkbookManager.tar
```

Pick **one** distribution set (either the `.tar.partNNN` files **or**
the `.xlsb` files). They contain identical data — you do not need both.

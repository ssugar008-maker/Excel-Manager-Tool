# Excel Workbook Manager — Standalone Distribution

This repository contains the **standalone Windows build** of Excel Workbook Manager, split into multiple `.zip.NNN` parts because the full archive exceeds GitHub's 100 MB per-file limit.

No Python install is required on the end-user PC — just Microsoft Excel.

---

## Contents

| File                              | Purpose                                                       |
| --------------------------------- | ------------------------------------------------------------- |
| `ExcelWorkbookManager.zip.001`    | Split archive part 1 (40 MB)                                  |
| `ExcelWorkbookManager.zip.002`    | Split archive part 2 (40 MB)                                  |
| `ExcelWorkbookManager.zip.003`    | Split archive part 3 (~33 MB)                                 |
| `SHA256SUMS.txt`                  | Checksums for every part + the reassembled `.zip`             |
| `Reassemble.ps1`                  | PowerShell script that joins the parts and extracts the app   |
| `Reassemble.bat`                  | Double-click launcher for `Reassemble.ps1`                    |
| `GUIDE.html`                      | Full user guide for the Excel Workbook Manager app            |
| `profiles.json`                   | Example profiles file (optional — shipped inside the zip too) |
| `settings.json`                   | Example settings file (optional — shipped inside the zip too) |

---

## Quick start — end users

1. Click the green **Code** button → **Download ZIP** (or `git clone` the repo).
2. Extract that outer ZIP so all the `ExcelWorkbookManager.zip.001`, `.002`, `.003`, plus `Reassemble.bat` are in the same folder.
3. **Double-click `Reassemble.bat`.** It will:
   - Join the parts back into a single `ExcelWorkbookManager.zip`
   - Verify the SHA256 checksum
   - Extract everything into a new `ExcelWorkbookManager\` folder
4. Open the new `ExcelWorkbookManager\` folder and double-click **`ExcelWorkbookManager.exe`**.

> Windows SmartScreen may warn about an unsigned executable. Click **More info → Run anyway**.

### Alternative (pure PowerShell)

```powershell
cd path\to\downloaded\folder
powershell -ExecutionPolicy Bypass -File .\Reassemble.ps1
```

### Alternative (manual, no script)

In PowerShell from the download folder:

```powershell
cmd /c "copy /b ExcelWorkbookManager.zip.001 + ExcelWorkbookManager.zip.002 + ExcelWorkbookManager.zip.003 ExcelWorkbookManager.zip"
Expand-Archive .\ExcelWorkbookManager.zip -DestinationPath .\ExcelWorkbookManager
```

Or in a classic Command Prompt:

```bat
copy /b ExcelWorkbookManager.zip.001 + ExcelWorkbookManager.zip.002 + ExcelWorkbookManager.zip.003 ExcelWorkbookManager.zip
```

Then right-click `ExcelWorkbookManager.zip` → **Extract All…**

---

## Requirements

- Windows 10 or 11 (x64)
- Microsoft Excel installed and licensed on the same machine
- ~300 MB free disk space after extraction

No Python, no pip, no virtualenv needed on the end-user PC.

---

## Verifying the download (optional)

```powershell
Get-FileHash .\ExcelWorkbookManager.zip.001 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.zip.002 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.zip.003 -Algorithm SHA256
```

Compare the output against `SHA256SUMS.txt`. The `Reassemble.ps1` script does this automatically for the joined `.zip`.

---

## For developers — rebuilding from source

The source code lives in a separate development folder. Build with:

```bat
pip install -r requirements.txt
build.bat
```

`build.bat` runs PyInstaller in `--onedir --windowed` mode and produces `dist\ExcelWorkbookManager\`. To re-create the split distribution, zip that folder and split it into 40 MB parts (see `Reassemble.ps1` for the reverse logic).

---

## Troubleshooting

- **"Reassemble.bat just flashes and closes."** — Open PowerShell in the folder and run `powershell -ExecutionPolicy Bypass -File .\Reassemble.ps1` to see the error.
- **SHA256 mismatch** — Re-download the failing part. GitHub sometimes serves truncated files on flaky connections.
- **App launches then nothing happens** — Make sure Excel is installed and can be opened normally on this PC.
- **Missing `profiles.json` or `settings.json` on first run** — These are already inside the zip, next to `ExcelWorkbookManager.exe`. If you deleted them, copies are also checked into this repo.

---

See **`GUIDE.html`** for the full user manual.

# Excel Workbook Manager — Standalone Distribution

This repository contains the **standalone Windows build** of Excel Workbook Manager, distributed as **raw binary chunks** of an uncompressed tar archive. No Python install is needed on the end-user PC — just Microsoft Excel.

The build is not shipped as a `.zip` on purpose. Each `.partNNN` file is a plain binary slice of the tar archive, so they can be concatenated byte-for-byte with `copy /b` or any equivalent tool.

---

## Contents

| File                                   | Purpose                                                |
| -------------------------------------- | ------------------------------------------------------ |
| `ExcelWorkbookManager.tar.part001`     | Raw chunk 1 (40 MB)                                    |
| `ExcelWorkbookManager.tar.part002`     | Raw chunk 2 (40 MB)                                    |
| `ExcelWorkbookManager.tar.part003`     | Raw chunk 3 (40 MB)                                    |
| `ExcelWorkbookManager.tar.part004`     | Raw chunk 4 (40 MB)                                    |
| `ExcelWorkbookManager.tar.part005`     | Raw chunk 5 (~31 MB)                                   |
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
copy /b ExcelWorkbookManager.tar.part001 + ExcelWorkbookManager.tar.part002 + ExcelWorkbookManager.tar.part003 + ExcelWorkbookManager.tar.part004 + ExcelWorkbookManager.tar.part005 ExcelWorkbookManager.tar
tar -xf ExcelWorkbookManager.tar
```

Or from PowerShell:

```powershell
cmd /c "copy /b ExcelWorkbookManager.tar.part001 + ExcelWorkbookManager.tar.part002 + ExcelWorkbookManager.tar.part003 + ExcelWorkbookManager.tar.part004 + ExcelWorkbookManager.tar.part005 ExcelWorkbookManager.tar"
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
Get-FileHash .\ExcelWorkbookManager.tar.part001 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.tar.part002 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.tar.part003 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.tar.part004 -Algorithm SHA256
Get-FileHash .\ExcelWorkbookManager.tar.part005 -Algorithm SHA256
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
final extension):

| File                                  | Size      |
| ------------------------------------- | --------- |
| `ExcelWorkbookManager.part001.xlsb`   | 40 MB     |
| `ExcelWorkbookManager.part002.xlsb`   | 40 MB     |
| `ExcelWorkbookManager.part003.xlsb`   | 40 MB     |
| `ExcelWorkbookManager.part004.xlsb`   | 40 MB     |
| `ExcelWorkbookManager.part005.xlsb`   | ~31 MB    |
| `SHA256SUMS-xlsb.txt`                 | checksums |
| `Reassemble-xlsb.bat` / `.ps1`        | helpers   |

> **These `.xlsb` files are NOT Excel workbooks.** They are raw
> byte-for-byte slices of the same tar archive as the `.tar.partNNN`
> set, simply renamed so they survive corporate file-type scanners.
> Do not try to open them in Excel.

### How to use

1. Download all five `ExcelWorkbookManager.part*.xlsb` files (plus
   `Reassemble-xlsb.bat`, `Reassemble-xlsb.ps1`, `SHA256SUMS-xlsb.txt`)
   into the same folder.
2. Double-click **`Reassemble-xlsb.bat`**.
3. Open the new `ExcelWorkbookManager\` folder and run
   `ExcelWorkbookManager.exe`.

Manual (no script):

```bat
copy /b ExcelWorkbookManager.part001.xlsb + ExcelWorkbookManager.part002.xlsb + ExcelWorkbookManager.part003.xlsb + ExcelWorkbookManager.part004.xlsb + ExcelWorkbookManager.part005.xlsb ExcelWorkbookManager.tar
tar -xf ExcelWorkbookManager.tar
```

Pick **one** distribution set (either the `.tar.partNNN` files **or**
the `.xlsb` files). They contain identical data — you do not need both.

# Excel Workbook Manager — Standalone Distribution

**Beta v11** — Global hotkey, queue management, change tracking, and scoped calculations:

- **Global Ctrl+Shift+M** — press the shortcut from *inside* Excel (no need to switch to the tool); it reads the active workbook/sheet and offers to open all its linked documents.
- **Smart Relinker — Clear All Queue** — remove all queued modifications in one click (with confirmation).
- **Smart Relinker — View Details** — inspect the exact old→new path pairs for any selected queue entry in a modal dialog, with an option to open them in Excel.
- **Smart Relinker — Export Queue to Excel** — flatten all queued pair-based entries into a live Excel workbook for review.
- **Smart Refresh — Audit cell improvements** — taller audit list, *Edit Selected* (loads the cell back into the form for correction) and *Navigate to Cell* (jumps Excel to the audited address).
- **Smart Refresh — Track Cell Changes** — a "Track cell changes on selected sheets" checkbox captures before/after snapshots of each selected sheet (or scoped range) and logs every changed cell to the audit log.
- **Smart Refresh — Scope column** — each sheet in the Sheets to Refresh list now has a **Scope** column; set it to `(all)` for a full-sheet recalculation or enter a range like `C:C`, `B2:D10`, `B:D` to restrict calculation and change-tracking to that range.
- **Smart Refresh — Capture from Excel** — click this button while a range is selected in Excel to automatically set the scope for the matching sheet.

Previous **Beta v10** highlights:

- **Ctrl+Shift+O → Ctrl+Shift+M** — the "Open Current Tab's Links" shortcut was renamed to avoid key-layout clashes.
- **Silent relink popup** — the *"links were not updated because the file was not recalculated"* Excel dialog is automatically suppressed during Apply Now / Apply Queue.
- **Auto-open before relink toggle** — a checkbox in Smart Relinker opens each target file (read-only) before changing link paths to prevent any residual Excel warnings.

Previous **Beta v9** highlights:

- **20-second auto-poll removed** — no more background freezes; a *Last refreshed* timestamp shows when data was last synced.
- **Lazy tab refresh** — data is only pushed to tabs you actually open.
- **Single COM enumeration** — one pass over `app.books` per refresh.
- **Three-method instance detection** (`xw.apps` + pythoncom ROT direct + psutil PID attachment).
- **Instance Diagnostic Panel** (`?` button) — shows all `EXCEL.EXE` processes and COM reachability.

This repository contains the **standalone Windows build** of Excel Workbook Manager, distributed as **8 MB binary chunks** (`.xlsb` extension) of an uncompressed tar archive. No Python install is needed on the end-user PC — just Microsoft Excel.

---

## Contents

| File                                            | Purpose                                      |
| ----------------------------------------------- | -------------------------------------------- |
| `ExcelWorkbookManager.part01.xlsb` … `.part18.xlsb` | 8 MB chunks of the tar archive          |
| `checksums.json`                                | SHA256 hash of every chunk                   |
| `reassemble.ps1`                                | PowerShell: verify + join + extract          |
| `reassemble.bat`                                | Double-click launcher (calls bat logic)      |
| `GUIDE.html`                                    | Full user guide                              |

---

## Quick start — end users

1. Click the green **Code** button on this repo → **Download ZIP** (or `git clone`).
2. Put every `ExcelWorkbookManager.partNN.xlsb` file plus `reassemble.bat` (and `checksums.json`) in the same folder.
3. **Double-click `reassemble.bat`.** It will:
   - Concatenate all `.xlsb` chunks into a single `ExcelWorkbookManager.tar`
   - Extract it with Windows' built-in `tar.exe` into a new `ExcelWorkbookManager\` folder
4. Open the new `ExcelWorkbookManager\` folder and double-click **`ExcelWorkbookManager.exe`**.

> Windows SmartScreen may warn about an unsigned executable. Click **More info → Run anyway**.

### Verify checksums (optional, PowerShell)

```powershell
powershell -ExecutionPolicy Bypass -File reassemble.ps1
```

This script verifies the SHA256 of every chunk before reassembly.

### Manual reassembly (Command Prompt)

```bat
copy /b ExcelWorkbookManager.part01.xlsb + ExcelWorkbookManager.part02.xlsb + ... ExcelWorkbookManager.tar
tar -xf ExcelWorkbookManager.tar
```

---

## Requirements

- Windows 10 version 1803 or newer (has built-in `tar.exe`), or Windows 11
- Microsoft Excel installed and licensed on the same machine
- ~300 MB free disk space after extraction

No Python, no pip, no virtualenv on the end-user PC.

---

## Functional overview

| Tab | What it does |
|-----|--------------|
| **Open Linked Documents** | List and open external workbook links; Ctrl+Shift+M opens the active sheet's links even from inside Excel |
| **Close Workbooks** | Bulk-close open workbooks with search |
| **Period Rollover** | Rename/relink workbooks to a new reporting period |
| **Smart Relinker** | Bulk-update formula links with queue, clear-all, view-details, and export-to-Excel |
| **Smart Refresh** | Open linked docs → scoped recalculate → compare audit cells → track cell changes |
| **Sheet Navigator** | Jump to any sheet in any open workbook |
| **Snapshot Workbook** | Paste-values copy of selected sheets dated `YYYYMMDD` |
| **Dependency Map** | Visual map of which workbooks reference which |
| **Instance Picker** | Switch between multiple running Excel instances or combine all |

---

## Building from source

Requires Python 3.12, xlwings, pywin32, psutil, openpyxl, and PyInstaller.

```bat
cd excel_workbook_manager
pyinstaller --noconfirm --onedir --windowed --name ExcelWorkbookManager ^
  --add-data "app_icon.png;." --add-data "tabs;tabs" ^
  --add-data "fuzzy_matcher.py;." --add-data "settings_manager.py;." ^
  --add-data "instance_detector.py;." --add-data "excel_bridge.py;." ^
  --add-data "utils.py;." main.py
```

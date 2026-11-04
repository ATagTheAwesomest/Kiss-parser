# KISS: Getting started (first-time users)

KISS reads the `.kss` files Tekla exports and builds the shop paperwork from them: the BOM workbook, the folder labels, the work orders, the quality reports and the ProNest helper scripts. This page walks through one run from nothing to finished.

## 1. What you need

- A Windows PC.
- **Microsoft Excel (the desktop app).** KISS builds every workbook through Excel, so it must be installed and signed in. Excel on the web will not work.
- The job's `.kss` files. Optional: the job's `... REQUIRED ASSM.xlsx` workbook (turns on folder labels and work orders).

Nothing else. No Python, no extra modules, no settings to change. KISS checks the PC itself every time it starts and tells you in plain words if something is missing.

## 2. Install

1. Get the KISS folder onto the PC (clone the `Kiss-parser` repo, or copy the folder). Anywhere is fine, for example `C:\Users\<you>\Documents\Projects\Kiss-parser`.
2. That's it. The first time you run it, KISS unblocks its own script files if Windows marked them as downloaded.

## 3. Launch

Double-click **`Start-KISS.bat`** in the KISS folder.

- It opens in the `jobs` folder next to it (created if missing). Put the job's `.kss` files (and the Required Assm workbook) in `jobs` first.
- To work in another folder, **drag that folder onto `Start-KISS.bat`**.
- You can also run it by hand: open PowerShell in the job folder and run `powershell -ExecutionPolicy Bypass -File <KISS folder>\tools\Parse-KSS.ps1`.

Everything KISS makes is saved in the folder you ran it in.

## 4. The screens, in order

Press **Enter** to take the answer shown as `(default)`. Some screens only appear when they apply.

| Screen | What it asks | Normal answer |
| --- | --- | --- |
| Checking this computer | Nothing. Shows green `[OK]` lines. A red `[X]` line says what to fix. | Just read it |
| Data source | `1` = KSS files, `2` = soundwall CSV | Enter (KSS) |
| Select KSS files | Up/Down to move, Space to tick, `A` all, `N` none, Enter to confirm | Tick the TLs for this run |
| TL for a file | Only if the file name has no TL. Type `TL012` or `12`, or Enter to leave blank | The TL number |
| Folder labels | Only if a Required Assm workbook is in the folder. Build labels and work orders from it? | Yes |
| Our job number / initials | Printed on the labels and work orders. Initials = who is making the folders | Job number, your initials |
| Assembly filter | Limit the run to a list of assemblies? | No for a full job. Yes for a REV or partial run, then paste the list |
| REV assemblies | Only after a filter with `R` marks. Should REV assemblies replace the same part in earlier TLs? | Yes |
| Duplicate parts | Only with more than one file. Merge identical parts into one row, or keep them separate per TL | Merge |
| Shop scripts | Only with more than one file. One combined set of scripts, or one set per TL | Combined |
| Quality Process Report | Make QPR inspection sheets? If yes: job name (sheet prefix) and one sheet per TL or per Sequence | No, unless QC needs it |

Then KISS works for a minute or two (Excel runs hidden; don't touch it) and shows a **Done** screen.

## 5. What you get

All in the folder you ran it in:

| Output | What it is |
| --- | --- |
| `<project> - Job <job> - <TLs> BOM.xlsx` | The master workbook: Statistics, Settings, Master BOM, plates, profiles, hardware, assemblies, labor, plus Folder Labels, TERRA Work Order and QPR tabs when you asked for them |
| `... - Filtered Assemblies.kss` | The filtered KISS file to import into Tekla. Finish stays as received; the TL goes in the cost code slot |
| `... QPR <date-time>.xlsx` | The QPR as its own workbook (only if you said Yes) |
| `Sort-PlatesByThickness - <TL>.bat` | Sorts the plate DXFs into thickness folders. Keep it next to the DXF folder |
| `Type-Quantities - <TL>.bat`, `Type-Grades - <TL>.bat` | Type values into ProNest. Click into the ProNest field first, as the script instructs |

At the end, say Yes to open the folder.

## 6. If something goes wrong

- **Red `[X]` on the checking screen:** it says what to do. Most common: Excel isn't installed or hasn't been opened and signed in once.
- **"Something went wrong and the run stopped":** take a screenshot of the window and send it to Andrew. Your `.kss` files are never changed.
- **A workbook name is locked:** close it in Excel and run again, or KISS saves a new copy with a time stamp.
- **A stray Excel left running in the background:** KISS closes its own, but if it was killed mid-run, end `EXCEL.EXE` in Task Manager.

More detail on how the numbers are produced: [PARSING.md](PARSING.md). The file format itself: [KISS-STANDARD.md](KISS-STANDARD.md).

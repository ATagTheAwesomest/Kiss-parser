# KISS: Getting started (first-time users)

KISS reads the `.kss` files Tekla exports and builds the shop paperwork from them: the BOM workbook, the folder labels, the work orders, the quality reports and the ProNest helper scripts. This page walks through one run from nothing to finished.

## 1. What you need

- A Windows PC.
- **Microsoft Excel (the desktop app).** KISS builds every workbook through Excel, so it must be installed and signed in. Excel on the web will not work.
- The job's `.kss` files. Optional: the job's `... REQUIRED ASSM.xlsx` workbook (turns on folder labels and work orders).

Nothing else. No Python, no extra modules, no settings to change. KISS checks the PC itself every time it starts and tells you in plain words if something is missing.

## 2. Install

1. Get the KISS folder onto the PC (clone the `Kiss-parser` repo, or copy the folder). Anywhere is fine, for example `C:\Tools\Kiss-parser`.
2. That's it. The first time you run it, KISS unblocks its own script files if Windows marked them as downloaded.

## 3. Launch

Double-click **`Start-KISS.bat`** in the KISS folder.

- It opens in the `jobs` folder next to it (created if missing). Put the job's `.kss` files (and the Required Assm workbook) in `jobs` first.
- To work in another folder, **drag that folder onto `Start-KISS.bat`**.
- You can also run it by hand: open PowerShell in the job folder and run `powershell -ExecutionPolicy Bypass -File <KISS folder>\tools\Parse-KSS.ps1`.

Everything KISS makes is saved in the folder you ran it in.

### Starting a new job: the blank Required Assm workbook

In the main menu choose **3**. KISS asks for the job number, saves `<job> REQUIRED ASSM.xlsx` in the folder you are working in (it never overwrites an existing file) and offers to open it. Fill it in:

1. **Sheet1**: replace an orange `Column#` header with the TL name (`TL1296`; a REV is `TLR1301`), then paste that TL's supplier assembly numbers underneath.
2. **TL Notes**: the sequences appear on the left by themselves. In the yellow table on the right type each sequence's Paint / Details, Explanation and Due Date.
3. **Exceptions**: paste only the ASSM numbers whose paint, explanation or due date differ from their sequence, and fill the yellow Override columns. KISS applies them to the Master BOM, work orders, labels and statistics.

Keep the file in the job folder next to the `.kss` files; KISS finds it by its name. It needs the Microsoft 365 version of Excel.

## 4. The screens, in order

Press **Enter** to take the answer shown as `(default)`. Some screens only appear when they apply.

| Screen | What it asks | Normal answer |
| --- | --- | --- |
| Checking this computer | Nothing. Shows green `[OK]` lines. A red `[X]` line says what to fix. | Just read it |
| Main menu | Centered list: `1` build the job paperwork (default), `2` soundwall CSV, `3` new blank Required Assm workbook, `H` this guide, `Q` quit. Each entry has a short explanation under it | Enter (build the paperwork) |
| Select KSS files | Up/Down to move, Space to tick, `A` all, `N` none, Enter to confirm | Tick the TLs for this run |
| TL for a file | Only if the file name has no TL. Type `TL012` or `12`, or Enter to leave blank | The TL number |
| Folder labels | Only if a Required Assm workbook is in the folder. Build labels, work orders, paint / due dates and exceptions from it? | Yes |
| Our job number / initials | Printed on the labels and work orders. Initials = who is making the folders | Job number, your initials |
| Assembly filter | Limit the run to a list of assemblies? | No for a full job. Yes for a REV or partial run, then paste the list |
| REV assemblies | Only after a filter with `R` marks. Should REV assemblies replace the same part in earlier TLs? | Yes |
| Duplicate parts | Only with more than one file. Merge identical parts into one row, or keep them separate per TL | Merge |
| Quality Process Report | Make QPR inspection sheets? If yes: job name (sheet prefix) and one sheet per TL or per Sequence | No, unless QC needs it |

Then KISS works for a minute or two (Excel runs hidden; don't touch it) and shows a **Done** screen.

## 5. What you get

All in the folder you ran it in:

| Output | What it is |
| --- | --- |
| `<project> - Job <job> - <TLs> BOM.xlsx` | The master workbook: Statistics, Settings, Master BOM, plates, profiles, hardware, assemblies, labor, plus Folder Labels, TERRA Work Order and QPR tabs when you asked for them |
| `... - Filtered Assemblies.kss` | The filtered KISS file to import into Tekla. Finish stays as received; the TL goes in the cost code slot |
| `... QPR <date-time>.xlsx` | The QPR as its own workbook (only if you said Yes) |
| `Sort-PlatesByThickness - <TLs>.bat` | Sorts the plate DXFs into thickness folders. Keep it next to the DXF folder |
| `Type-Plates - <TLs>.bat` | Types plate quantities or grades into ProNest (you choose Q or G after the preview, and can do both one after the other). Click into the ProNest field first, as the script instructs |

There is one set of these two scripts for the whole run, however many TLs you parsed. **When you start one, it asks which transmittals / sequences you are nesting now** (see section 5a).

## 5a. Using the shop scripts: pick what you are nesting

Parse every TL at once, then tell the script which ones are being nested today. Start either `.bat` file and you see a list like:

```
[1] TL1282    seq 5203     14 plate mark(s)   qty 40
[2] TL1296    seq 5131     22 plate mark(s)   qty 100
[3] TL1296    seq 5132      9 plate mark(s)   qty 31
[A] All of them
```

Type what you are nesting now, separated by spaces or commas: numbers (`1 3`), a TL name (`TL1296` = every sequence of it) or a sequence (`5131` = that sequence in every TL). `A` picks everything. The script then shows how many plate marks and how much plate quantity that is, then a **preview** to check against your other documents: the first 5, the last 5 and the 5 highest-quantity plates of the selection (mark, thickness, quantity, grade; the Sort script shows the `.dxf` name and whether the file was found). Enter continues, `R` picks again.

- **Sort** moves only the DXFs of plates in your selection. Everything else stays where it is.
- **Type-Plates** asks `[Q]` quantities or `[G]` grades. Quantities: for each selected plate, the quantity from the selected TLs / sequences only. A plate that appears in several TLs gets the sum of the ones you picked, and each line on screen shows where the number came from (`TL1282/5203:50 TL1296/5131:100`).
  Grades: the grade of each selected plate.
- Plates that belong only to TLs you did not pick are skipped, so the rows you type line up with only the DXFs you imported into ProNest.
- The last answer is remembered (in `Shop-Selection - <TLs>.txt` next to the scripts), so for the second script press Enter to reuse it.

At the end, say Yes to open the folder.

## 6. If something goes wrong

- **Red `[X]` on the checking screen:** it says what to do. Most common: Excel isn't installed or hasn't been opened and signed in once.
- **"Something went wrong and the run stopped":** take a screenshot of the window and send it to whoever maintains KISS. Your `.kss` files are never changed.
- **A workbook name is locked:** close it in Excel and run again, or KISS saves a new copy with a time stamp.
- **A stray Excel left running in the background:** KISS closes its own, but if it was killed mid-run, end `EXCEL.EXE` in Task Manager.

More detail on how the numbers are produced: [PARSING.md](PARSING.md). The file format itself: [KISS-STANDARD.md](KISS-STANDARD.md).

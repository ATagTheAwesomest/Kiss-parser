# KISS: How `Parse-KSS.ps1` Reads Its Input Files

The official field-by-field format is in [KISS-STANDARD.md](KISS-STANDARD.md); note that its column names differ from the labels used below (e.g. D field 1 is the standard's "Drawing no", field 11 is "Notes", and weights live in fields 14/15). Draft README for the KISS tool (source: `KISS/Parse-KSS.ps1`, about 3,390 lines of PowerShell). Written from reading the script and the sample `.kss` files in the project folder. Meant to later feed an SOP. The script was not modified or run.

## 1. What it does

`Parse-KSS.ps1` reads KISS export files (`.kss`) that come from other companies (Tekla-based detailers), and produces:

1. An Excel workbook, `<Project> - Job <Job> - <TLs> BOM.xlsx`, with BOM, plate, profile, hardware, assembly, labor and statistics tabs.
2. A "Filtered Assemblies" `.kss` file, a cleaned copy of what was kept.
3. Windows `.bat` shop scripts that move DXFs into thickness folders and type quantities and grades into ProNest.
4. Optionally, a Quality Process Report (QPR) workbook.

It also has a second input mode that reads a Soundwall post-schedule CSV (section 7).

**Requirements:** Windows, PowerShell, desktop Excel (the script drives Excel through COM, `New-Object -ComObject Excel.Application`). It writes all output into the current folder, so run it from the folder holding the `.kss` files.

## 2. Input: the KSS file format

A `.kss` file is plain text, one record per line, comma separated. The first field is the record type. Sample (structure taken from a real file, header details replaced):

```
KISS,1.1,TeklaStructures
H,00-000,PROJECT NAME,CLIENT/COMPANY,10/14/26,16:36:51,F,
*
D,52031000,0,52031000,52031000,2,HSS,6X6X3/8,A500C,7035.80,3-1,COLUMN(TS-C),,8-22,,
S,5203,2
A,520301-02
D,52031000,0,52031000,3bp43,2,PL,3/4X7,A572-50,304.80,,BASE_Plate,,23,,0.0
```

| Type | Meaning | Fields the script uses (0-based index after splitting on commas) |
|---|---|---|
| `KISS` | File banner (`KISS,1.1,<software>`) | Ignored by the parser |
| `H` | Header | 1 = job number, 2 = project name. Other fields (client, date, time) are ignored |
| `*` | Separator between assemblies; any line starting with `*` is skipped | none |
| `D` | Detail: one part inside an assembly | see next table |
| `S` | Assembly summary (`S,<4-digit sequence>,<assembly qty>`) | Ignored when reading. Written and checked when the script writes a KSS |
| `A` | Assembly drawing number, attached to the most recent `D` assembly | 1 = drawing |
| `L` | Labor line (holes, welds, cuts, copes) attached to the most recent `D` | 1 = type, 2 = qty, 3 = size 1, 4 = size 2, 5 = notes |

The sample files in the folder contain no `L` lines, so the labor tab is probably empty for these jobs. This hasn't been tested with a file that has them.

### `D` record columns

| Index | Field | Notes |
|---|---|---|
| 1 | Assembly ID (standard: Drawing no) | e.g. `52031000`. First 4 digits are the sequence (e.g. 5203) |
| 3 | Main mark | Mark of the assembly's main part |
| 4 | Part mark | e.g. `3bp43`. **Blank for loose hardware**; the script then builds a mark (section 4) |
| 5 | Qty | Must be an integer, or the whole line is silently skipped |
| 6 | Shape | `PL`, `HSS`, `W`, `HS`, ... (upper-cased) |
| 7 | Size | e.g. `3/4X7` for plates (thickness X width), `6X6X3/8` for HSS |
| 8 | Grade | e.g. `A572-50` |
| 9 | Length | Looks like millimeters in the samples (`7035.80`) but the script stores it as text and does not convert |
| 10 | Finish / coating | Optional free text from the detailer, e.g. `3-1`, `0-0`, `SEE NOTE`, or blank. Shown as COATING on the QPR. Kept as received in filtered files (gotcha 12) |
| 11 | Description (standard: Notes) | Optional, e.g. `SHEAR_PLATE` |
| last numeric | Unit weight | See gotcha 1 (`-StandardWeight` reads fields 14/15). Field 13 is the cost code |

A `D` line must have at least 10 comma-separated fields to be accepted. Fields are split on commas with no quote handling, so a comma inside a description would shift every later column. The script's own generated KSS files replace commas in descriptions with `;` for this reason.

## 3. Run flow (what the operator is asked)

1. **Data source:** `1` = KSS files, `2` = Soundwall CSV.
2. **File picker:** lists every `.kss` in the current folder (arrow keys and Space to check, `A` all, `N` none, Enter to confirm). It hides files whose names contain `- Filtered Assemblies` or `(Generated)`, so the script's own outputs aren't re-read.
3. **TL (transmittal) detection:** the TL tag comes from the **file name**, matching `TL` followed by digits (`TL1282`, `TL 12`, `TL_012`, `TL-12` all become `TL<digits>`). If none is found, the operator is asked to type one or leave it blank. This is why file names matter.
4. **Parse:** each file is read line by line and `H`, `A`, `D`, `L` lines are collected into in-memory lists (Master data and Labor data). The first file's `H` line sets the job number and project name. If a later file's header differs, a yellow warning is shown and the first file's values are kept.
4b. **Folder labels (optional):** if the folder holds a `... REQUIRED ASSM.xlsx`, the script offers to add a `Folder Labels` tab to the BOM workbook (section 6a). It asks for our job number (default: the leading digits of the file name, e.g. `26102`).
5. **Assembly whitelist (optional):** limits the run to listed assemblies (section 5).
6. **REV override (only if the whitelist has `R` entries):** replace older TL versions of an assembly.
7. **Duplicate handling (only if more than one file):** `M`erge identical part marks across files into one row (default), or `S`eparate to keep each file's rows apart by TL. Merging warns when the same mark has a different shape/size/grade/length in different files.
8. The filtered KSS is written, Excel is launched and the tabs are built, and the workbook is saved.
9. Shop scripts are generated: one select-aware set for the whole run (no questions at parse time).
10. Optional: Quality Process Report, grouped by TL (default) or by sequence (first 4 digits of the assembly), with a job name prefix for the sheet names. Sheet names (and the Job: box on each sheet) read `{job} - {TL}`, or `{job} - {TL} - {sequence}` when grouped by sequence. The QPR questions are asked before the BOM workbook is built: answering Yes adds the QPR sheets as tabs at the end of the master BOM workbook (same sheet layout and names) and also saves the separate QPR workbook as before. In the master workbook the Job: box follows `Set_CustomerJob` on the Settings tab (when the job name was left at the default), and on sequence-grouped sheets the line under the title shows the Due date and Paint / REVISES note read live from the Sequence Notes tab, so edit them there.

## 4. Parsing logic details

- **Material category** (decides which tab a row lands on), checked in this order:
  1. **Hardware** if shape is one of `HS, MB, AB, WA, WS, NUT, WASHER, BOLT`, or the description contains `NUT, WASHER, BOLT, ANCHOR, ROD, STUD, FIELD, SHOP`, or the mark starts with `nut/washer/bolt/anchor`, or the size contains `NUT, WASHER, HEAVY_HEX`.
  2. **Plate** if shape is `PL, SHT, BPL, PLATE, SHEET, CP, AR, UM, T1`.
  3. Otherwise **Profile**.
  The description match is a substring match, so a plate with `FIELD` or `SHOP` in its description is classed as Hardware.
- **Blank part marks:** the effective mark becomes `<SHAPE>-<description>`, or `<SHAPE>-<size>` if there is no description. The raw (blank) mark is kept separately for the integrity check.
- **Plate thickness:** the text before the first `X` in the size. It can be a decimal (`0.375`), a fraction (`3/8`), a mixed number (`1-1/2` or `1 1/2`), or a gauge (`10GA`, via a lookup table from 3GA to 26GA). Anything unparseable returns 0, which becomes a folder named `0`.
- **Weight:** see gotcha 1.
- **Merge rules:** the Plates, Profiles and Hardware tabs group rows by part mark (or by source file and part mark in Separate mode), sum the quantities, and take the first row's size, grade, length and weight.
- **Master BOM extended weight** is a live Excel formula (`Qty x Unit Wt`). The `Status` column is filled with `PENDING`.

## 5. Assembly whitelist and REV logic

The operator can paste a list, type `@path.txt` to load a file, or press Enter to use whatever is on the clipboard. Tokens split on commas, semicolons, pipes and whitespace, so a column copied from Excel works.

| Syntax | Meaning |
|---|---|
| `51328511` | Keep this assembly |
| `TL1296 51328511 51318512` | A TL token at the start of a line (or anywhere) tags the assemblies that follow with that TL. Overrides the TL taken from the file name |
| `1_TL1301` | Number prefix sets manual TL ordering (used when Separate mode removes duplicates) |
| `R51328511` | `R` prefix marks a **REV**: this version replaces older TL versions of the assembly |
| `51328511---4` | Wanted quantity. Part quantities are scaled by wanted qty / qty of the main-part row (the `D` row whose part mark equals the main mark) |

Matching is case-insensitive. Names not found in the parsed data are logged as ignored; if nothing matches, the run stops.

**Gotcha:** an assembly name that genuinely starts with the letter `R` would be read as a REV marker, and the `R` is stripped.

REV override keeps the newest TL (highest number in the file name) for each assembly, and for REV-marked assemblies replaces the earlier version. If REV overrides are declined, duplicates show for every TL.

## 6. Outputs

**Workbook tabs:** Master BOM (16 columns: TL, Assembly, Drawing, MainMark, PartMark, Qty, Category, Shape, Size, Length, Grade, Description, Unit Wt, Total Wt, Status, SourceFile), Plates & Sheets, Profiles, Hardware & Fasteners, Shop Assemblies (one block per assembly, with a VARIANCE note when the same assembly is defined differently in different files), Labor & Weld Operations, and Statistics (dashboard of formulas), and, when a Required Assm workbook is used, Folder Labels (section 6a). Plate/profile/hardware tabs print "No ... Data Found" when empty.

**File naming:** `<Project> - Job <Job> - <TL1>_<TL2>... BOM.xlsx`. If the old file is locked, a timestamp is appended instead. With many TLs the name is shortened to `<first>-<last> (n TLs)`.

**Filtered KSS:** `... - Filtered Assemblies.kss`, rebuilt from the kept data. Sorted by assembly. If an assembly appears in several files, only the first selected file's version is used. The script then re-reads it with its integrity check (`Test-KissAssemblyIntegrity`): every `D` block must have 12+ fields and a numeric qty, each assembly must end with a matching `S` line, the `S` qty must match the main-part qty, and only `KISS/H/D/S/A/*/END` record types are allowed. A failed check deletes the file and stops the script.

**Shop scripts (only generated if the selection has plates):**
One set of two scripts for the whole run (no combined / separate question; the old per-file mode and the unused master launcher are gone). `<label>` is the TL list. The data of every plate is embedded with its transmittal + sequence, and each script **asks at runtime which transmittals / sequences are being nested now** before it does anything.

- **Selection key:** transmittal + sequence. The transmittal is the one the part is placed under in the Master BOM, labels and work orders (a REV assembly counts under its own TLR); the sequence is the first 4 characters of the assembly. Each key is listed with its plate-mark count and plate quantity.
- **Answering:** numbers from the list, a TL name (all its sequences), a sequence (in all TLs) or `A` (all), separated by spaces or commas. The answer is saved in `Shop-Selection - <label>.txt` next to the scripts; Enter at the prompt reuses it. After a summary (plate marks and plate quantity selected) Enter continues and `R` picks again.
- A plate mark is **in** the selection if it occurs in at least one selected key. Its quantity is the sum over the selected keys only, so a mark with 50 in TL1282 and 100 in TL1296 types 50 when only TL1282 is selected, 100 for TL1296, 150 for both. Marks in no selected key are skipped (not moved, not typed).
- `Sort-PlatesByThickness - <label>.bat`: moves each selected `<partmark>.dxf` into a folder named by decimal thickness. It looks in a `dxf\` subfolder if one exists, otherwise the batch file's own folder. Reports moved, missing and left-alone counts.
- `Type-Plates - <label>.bat`: after the selection and preview you choose `Q` (quantities) or `G` (grades), then after a 5-second countdown it types each selected plate's quantity (selected keys only) or grade into the active window (ProNest) with SendKeys, pressing Down after each. Rows are sorted by part mark (A-Z). The operator must click the first cell first; afterwards the Q / G menu returns so the other value can be typed. Each typed line also prints where the quantity came from, e.g. `TL1282/5203:50 TL1296/5131:100`.
- **Preview** (before anything runs, built from the selection): the first 5 plates (A-Z), the last 5 and the 5 highest quantities, each with mark, thickness folder, quantity and grade; the Sort script also shows `<mark>.dxf` and whether the file is found. Implemented by collecting the selected plates into `P_<n>_*` variables.
- If one mark has different sizes or grades in different transmittals, the generator warns and uses the first for thickness and grade.
- Menus use ANSI colors (Windows 10 or newer console).

## 6a. Folder labels (Required Assm workbook)

Needs a `*REQUIRED ASSM*.xlsx` in the run folder (git-ignored like every other job file). Not used in Soundwall CSV mode.

**Reads `Sheet1`** (the first sheet with a `TL Name` cell in column A). One column per transmittal:

| Row | Used for |
|---|---|
| `Paint / TL Details` (above the header) | Fallback coating for the whole transmittal, used only when the `TL Notes` tab has none for that sequence. |
| `Due Date` (above the header) | Fallback due date, same rule. |
| `TL Name` (header) | Transmittal, e.g. `TL1296`. A REV is `TLR1301` and keeps the `R` on the label. Columns that are not named `TL#### / TLR####`, or that hold no assembly numbers, are ignored. |
| Rows under the header | AFCO assembly numbers. The first 4 digits are the sequence (an `R` prefix is ignored). Each unique sequence in a column gets its own label (line 1). |

**Coating and due date come from the `TL Notes` tab** ("Sequence Details Register"), which is keyed by 4-digit sequence: columns `Sequence`, `Paint / Details`, `Explanation`, `Due Date`. Each label takes its own sequence's `Paint / Details` (line 3, left) and `Due Date` (line 3, right, as `DUE <date>`); real dates are shown as `M/d/yy`, text as typed, blank = left out. `Explanation` is read too (it feeds the Master BOM `Explanation` column and the Exceptions defaults below). The tab holds the yellow input table plus a formula-spilled copy; the first non-blank value per sequence is used.

The Sheet1 `Explanation` row is not used. The other tabs in that workbook (except `TL Notes` and `Exceptions`) are formulas over Sheet1 and are not read.

**One label per TL and sequence** (one per folder), sorted by TL number, then sequence. If KSS files were selected for the run, only transmittals matching a selected file's TL number get a label (`TL1301.kss` matches `TLR1301`); skipped ones are reported in yellow.

| Line | Content |
|---|---|
| 1 (big) | `<our job> - <TL> - <sequence>`, e.g. `12345 - TL1296 - 5131`. On a REV (`TLR1301`) only the `R` is italic. Our job is typed at the prompt (default: leading digits of the Required Assm file name). |
| 2 (small) | Left: `<project name> - <customer job number>` (the KSS header job number as written, e.g. `SAMPLE PROJECT - 12-345`). Right-aligned: `IS: <initials>`, the initials typed at the prompt "Initials of who is making the folders". Left blank it prints just `IS:` to hand-write (the text is `$IssuedLabel` in `Add-FolderLabelSheet`). |
| 3 | Left (11 pt, comments / paint info): the sequence's `Paint / Details` from `TL Notes`, blank if none. On a REV (`TLR####`) label it starts with `REVISES TL1296` (several: `REVISES TL1296, TL1313`): the transmittals in the Required Assm workbook that hold any of the same assembly numbers (digits only, so an `R` prefix is ignored) within that sequence. Other REV transmittals are listed only when no non-REV one matches. Paint text follows after ` - `. Right-aligned (18 pt bold): `DUE <m/d/yy>` from the `TL Notes` due date, or the placeholder `DUE    /    /` to fill in later. |

Every label cell has a thin black border on all four edges (merged cells are outlined as a whole), on both label tabs. 

**Print setup (tab `Folder Labels`):** each label is a 3.5 in x 1.125 in block (252 x 81 pt) of three rows (40.5 + 15 + 24.75 pt) plus a 0.75 pt spacer row so the bottom border is not clipped, with a manual page break between labels, margins 0, landscape, centered, print area covering the labels only, scaling set to fit all columns on one page. Select the Dymo as the printer with the 3.5 x 1.125 in address label size (e.g. 30252), then print the tab. Text shrinks to fit. If a label clips or spills onto an extra page, check the Dymo paper size in Page Setup. Excel cannot create a custom paper size by itself, so the paper size comes from the Dymo driver.

**`TL Labels` tab:** built automatically right after `Folder Labels` from the same data (no extra prompt), one 3.5 x 1.125 in label per transmittal, ordered by TL number, with the same print setup (margins 0, fit all columns on one page, a page break between labels). Rows: (1) the transmittal big and centered, e.g. `TL1296`; (2) that transmittal's sequences big, side by side and evenly spread (a 12-column grid, so 1 to 6 sequences share the width; more than 6 shrink to fit); (3) under each sequence its `DUE <m/d/yy>` from `TL Notes` in small type, or `DUE    /    /` when none. A REV (`TLR1301`) is different: the `R` is italic and the right third of row 1 carries a small `REVISES TL1296`, found the same way as on the Folder Labels tab (union over the REV's sequences). Our job number and the customer job are not on this tab.
Assumptions: the due date text is used as typed in the sheet (real dates are shown as `M/d/yy`); coating comes only from the `Paint / TL Details` row, not from the KSS Finish field.

## 6b. ASSM exceptions (Required Assm `Exceptions` tab)

The Required Assm workbook's `Exceptions` tab lists individual assemblies (ASSM numbers) whose **Paint / Details**, **Explanation** or **Due Date** differ from the default of their sequence on the `TL Notes` tab. The tab itself computes the matched transmittal(s), the sequence (first 4 digits), the defaults and the effective values; KISS reads only what is typed: the `ASSM Number` column and the three yellow **Override** columns. An ASSM with no override filled in "uses default" and is ignored. If the same ASSM is listed twice the last row wins (yellow warning). Quantities, sequences, REV placement, weights and the filtered `.kss` are not affected: an exception changes paint / due / explanation only.

Effective value for an assembly = its override if one is set, else the sequence default. What uses it:

| Where | Effect |
|---|---|
| `Assembly Exceptions` tab (new, built after `Sequence Notes`) | One row per exception ASSM that is in the run, with the transmittal it is placed on (a REV'd assembly shows `TLR####`), its sequence, the default (formula into `Sequence Notes`), the override (yellow, editable) and the effective paint / explanation / due. ASSMs on the Exceptions tab that are not in the run are listed at the bottom as `Not in this run: <reason>` (their transmittal's `.kss` was not selected, or the assembly filter left them out, or they are not in Sheet1) and also reported in yellow on screen. |
| Master BOM | `Due Date` and `Paint / Notes` take the effective value for exception assemblies; new columns `Explanation` (AD) and `Exception` (AE, `EXCEPTION` on the flagged rows) appear when the run has exceptions. Other assemblies follow their folder, as before. |
| TERRA Work Order | The item's NOTES cell shows its effective paint (e.g. `NO PAINT` inside a `GREY PAINT` sequence), plus ` - <explanation>` when the exception has its own, plus ` - DUE <date>` when its due date differs from the work order's date. The work order header (one per transmittal and sequence) is unchanged. |
| `Sequence Notes`, Folder Labels, QPR (sequence sheets) | New columns `Exception ASSMs` (count in that folder) and `Exception Paint` (effective paint of the first one). The Label Notes text appends them: `GREY PAINT - 12 ASSM NO PAINT`. Folders and labels are not split; the shop sees from the label that some pieces differ and from the work order which ones. |
| Statistics | Headlines `Assemblies with exceptions` and `Exception ASSMs not in this run`. |

Interpretation to confirm: an exception changes how an assembly is treated, not which folder it is in. If an exception is meant to move an assembly to a different folder or work order, that is not done today.

## 6c. Making a blank Required Assm workbook

In the main menu choose **3** (`New blank Required Assm workbook`). KISS asks for our job number, writes `<job> REQUIRED ASSM.xlsx` (or `TEMPLATE REQUIRED ASSM.xlsx`) in the working folder (never overwriting an existing file) and offers to open it. The workbook has every tab (`Sheet1`, `Table1 Transposed`, `Sequences in Each TL`, `Sequences by TL`, `TL Notes`, `Exceptions`) with all formulas, tables and formatting and no data; `Sheet1` starts with empty `Column1`..`Column20` headers. The workbook is embedded in `Parse-KSS.ps1` itself (`$script:RequiredAssmTemplateBase64`), so there is still only one file and no Excel is needed to write it, only to fill it in. It needs Microsoft 365 Excel (the tabs use LET / LAMBDA / BYCOL / FILTER). To change the template, open a written copy, edit it, and re-embed the saved `.xlsx` as base64 in that variable.

## 7. Soundwall CSV mode

Reads a post schedule with `Import-Csv`. Required column: `Post ID` (rows without it are skipped). Optional columns include `Post Size`, `Post Type`, `Post Length`, `Base Plate Type`, `Anchor Rod Type`, `Flat Bar Left/Right Size` and `Length`, `Seat Plate Left/Right Size` and `Height`, and a `Drawing`/`Dwg` column. Each post becomes an assembly with parts: the post (Profile, weight from the `X<lb/ft>` in names like `W8X24` times length), base plate, anchor rods (qty = the "anchors per post" answer, default 4), flat bars (a leading `(2)` in the size sets the qty), and seat plates. Flat bars and seat plates get generated marks (`FB1`, `FB2`, `SPL1`...), and plate weight is thickness x width x length x 0.2836 lb/in³. Lengths over 100 with no feet/inches format are treated as millimeters (divided by 304.8). The script also writes a companion `... (Generated).kss` and asks whether to continue to the workbook. Grade is left blank for everything in this mode.

## 8. Gotchas and risks (things the SOP should cover)

1. **Weight can be wrong (default behavior).** The script takes the unit weight from the *last field that parses as a number*, scanning right to left. On Tekla files, which carry no weight, that is the cost code on main marks (e.g. 23) and `0.0` on parts. Run with `-StandardWeight` to read gross weight (field 14), else net (field 15), else 0 instead. Spot-check weights either way.
2. **Length units are not converted** for KSS files (mm in the samples). Do not read the Length column as inches or feet.
3. **Rows can vanish silently:** a `D` line with a non-integer qty or fewer than 10 fields is skipped without a warning.
4. **No quote handling** in `D` lines (commas inside text break the columns).
5. **TL comes from the file name**, not from inside the file. Rename files carefully (`TL1282` style).
6. **Merge mode can hide conflicts:** the same part mark with different sizes across TLs is merged using the first row's data. The script only warns.
7. **Quantity scaling** with `---N` only works when the assembly has a main-part row; otherwise a yellow message is shown and the quantities are unscaled.
8. **Excel dependency:** a crash mid-run can leave a hidden `EXCEL.EXE` process. Close it in Task Manager if the next run misbehaves.
9. **Typing macros are timing-based** (SendKeys with fixed delays) and type into whatever window has focus. The operator must not touch the keyboard during a run.
10. **Output goes to the current folder** and old outputs are not cleaned up. That is why the folder holds many dated BOM and QPR workbooks.
11. The script file is about 150 KB with Excel layout code making up most of it; the actual parsing is roughly lines 1,380–1,480 (KSS) and 966–1,115 (Soundwall).
12. **Transmittal placement.** The filtered `.kss` keeps Finish/coating (field 10) as received and carries the TL in field 13 (cost code / pay category), replacing the cost code value from the detailer. The drawing revision (field 2) and fields 12+ are passed through. `-TransmittalInFinish` restores the old behavior (TL written over Finish, revision forced to `0`). Filtered files made by older versions still have the TL in field 10; re-run them from the original `.kss`.

## 9. Open questions

- Which companies' KSS files does the shop receive, and do they always follow the layout above? Do any ever contain `L` lines?
- Is the weight behavior in gotcha 1 known and accepted (e.g. weights come from elsewhere), or a bug to fix?
- Should the SOP follow the team's workflow end to end (receive KSS, run script, review BOM, sort DXFs, type quantities and grades into ProNest)? A short breakdown from you would let me write it.

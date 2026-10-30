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
| 10 | Finish / coating | Optional free text from the detailer, e.g. `3-1`, `0-0`, `SEE NOTE`, or blank. Shown as COATING on the QPR. Never overwritten by the TL (see section 8, gotcha 12) |
| 11 | Description (standard: Notes) | Optional, e.g. `SHEAR_PLATE` |
| last numeric | Unit weight | See gotcha 1. The standard puts gross part weight at 14 and net weight at 15; field 13 is the cost code |

A `D` line must have at least 10 comma-separated fields to be accepted. Fields are split on commas with no quote handling, so a comma inside a description would shift every later column. The script's own generated KSS files replace commas in descriptions with `;` for this reason.

## 3. Run flow (what the operator is asked)

1. **Data source:** `1` = KSS files, `2` = Soundwall CSV.
2. **File picker:** lists every `.kss` in the current folder (arrow keys and Space to check, `A` all, `N` none, Enter to confirm). It hides files whose names contain `- Filtered Assemblies` or `(Generated)`, so the script's own outputs aren't re-read.
3. **TL (transmittal) detection:** the TL tag comes from the **file name**, matching `TL` followed by digits (`TL1282`, `TL 12`, `TL_012`, `TL-12` all become `TL<digits>`). If none is found, the operator is asked to type one or leave it blank. This is why file names matter.
4. **Parse:** each file is read line by line and `H`, `A`, `D`, `L` lines are collected into in-memory lists (Master data and Labor data). The first file's `H` line sets the job number and project name. If a later file's header differs, a yellow warning is shown and the first file's values are kept.
5. **Assembly whitelist (optional):** limits the run to listed assemblies (section 5).
6. **REV override (only if the whitelist has `R` entries):** replace older TL versions of an assembly.
7. **Duplicate handling (only if more than one file):** `M`erge identical part marks across files into one row (default), or `S`eparate to keep each file's rows apart by TL. Merging warns when the same mark has a different shape/size/grade/length in different files.
8. **Shop script style (only if more than one file):** `C`ombined (one set for all TLs, default) or `S`eparate per TL.
9. The filtered KSS is written, Excel is launched and the tabs are built, and the workbook is saved.
10. Shop scripts are generated.
11. Optional: Quality Process Report, grouped by TL (default) or by sequence (first 4 digits of the assembly), with a job number prefix for the sheet names.

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

**Workbook tabs:** Master BOM (16 columns: TL, Assembly, Drawing, MainMark, PartMark, Qty, Category, Shape, Size, Length, Grade, Description, Unit Wt, Total Wt, Status, SourceFile), Plates & Sheets, Profiles, Hardware & Fasteners, Shop Assemblies (one block per assembly, with a VARIANCE note when the same assembly is defined differently in different files), Labor & Weld Operations, and Statistics (dashboard of formulas). Plate/profile/hardware tabs print "No ... Data Found" when empty.

**File naming:** `<Project> - Job <Job> - <TL1>_<TL2>... BOM.xlsx`. If the old file is locked, a timestamp is appended instead. With many TLs the name is shortened to `<first>-<last> (n TLs)`.

**Filtered KSS:** `... - Filtered Assemblies.kss`, rebuilt from the kept data. Sorted by assembly. If an assembly appears in several files, only the first selected file's version is used. The script then re-reads it with its integrity check (`Test-KissAssemblyIntegrity`): every `D` block must have 12+ fields and a numeric qty, each assembly must end with a matching `S` line, the `S` qty must match the main-part qty, and only `KISS/H/D/S/A/*/END` record types are allowed. A failed check deletes the file and stops the script.

**Shop scripts (only generated if the selection has plates):**
- `Sort-PlatesByThickness - <label>.bat`: moves each `<partmark>.dxf` into a folder named by decimal thickness. It looks in a `dxf\` subfolder if one exists, otherwise the batch file's own folder. Reports moved and missing counts.
- `Type-Quantities - <label>.bat` and `Type-Grades - <label>.bat`: after a 5-second countdown, type each plate's total qty or grade into the active window (ProNest) with SendKeys, pressing Down after each. Rows are sorted by part mark (A-Z). The operator must click the first cell first.
- `Master-ShopScripts - <TLs>.bat`: `Write-MasterScript` exists but is **never called** (it appears only as a definition), and the script ends by logging "Generated shop scripts without a master launcher." So no master menu is generated today.

## 7. Soundwall CSV mode

Reads a post schedule with `Import-Csv`. Required column: `Post ID` (rows without it are skipped). Optional columns include `Post Size`, `Post Type`, `Post Length`, `Base Plate Type`, `Anchor Rod Type`, `Flat Bar Left/Right Size` and `Length`, `Seat Plate Left/Right Size` and `Height`, and a `Drawing`/`Dwg` column. Each post becomes an assembly with parts: the post (Profile, weight from the `X<lb/ft>` in names like `W8X24` times length), base plate, anchor rods (qty = the "anchors per post" answer, default 4), flat bars (a leading `(2)` in the size sets the qty), and seat plates. Flat bars and seat plates get generated marks (`FB1`, `FB2`, `SPL1`...), and plate weight is thickness x width x length x 0.2836 lb/in³. Lengths over 100 with no feet/inches format are treated as millimeters (divided by 304.8). The script also writes a companion `... (Generated).kss` and asks whether to continue to the workbook. Grade is left blank for everything in this mode.

## 8. Gotchas and risks (things the SOP should cover)

1. **Weight comes from the file, and Tekla files have none.** The parser now reads gross part weight (field 14), else net part weight (field 15), else 0. Tekla exports (Kiewit, Crusoe) leave both blank or `0.0`, so their BOM weights show 0; Advance Steel files carry real weights. (Before, it used the last numeric field, which was the cost code on main marks.) Weights in the BOM should still be spot-checked.
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
12. **Coating/finish was replaced by the transmittal (fixed).** When a whitelist assigned TLs, the parser copied the TL into the Finish field (field 10), so Filtered Assemblies `.kss` files carried `TL1296` where the original had `3-1`, `0-0`, `2-9`, `SEE NOTE` or blank. Finish is now never touched. The TL is carried in the standard's *lot name* slot instead: the 4th field of the `S` line (`S,5131,1,TL1296`). Fields 12 and up of each `D` line (part name, cost code, weights, area) and the drawing revision (field 2) are now passed through unchanged; before, the rewrite hard-coded revision `0` and put the weight in the cost-code slot. Files filtered before the fix still carry the TL in field 10; re-run them from the original `.kss`.

## 9. Open questions for Andrew

- Which companies' KSS files does the shop receive, and do they always follow the layout above? Do any ever contain `L` lines?
- Is the weight behavior in gotcha 1 known and accepted (e.g. weights come from elsewhere), or a bug to fix?
- Should the SOP follow the team's workflow end to end (receive KSS, run script, review BOM, sort DXFs, type quantities and grades into ProNest)? A short breakdown from you would let me write it.

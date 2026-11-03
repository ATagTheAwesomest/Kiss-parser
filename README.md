# KISS

Tools for turning KISS files (`.kss`, the steel detailing export from Tekla Structures and other detailing
software) into the shop's working documents.

## Layout

| Folder | What lives there |
| --- | --- |
| `tools/` | The programs. Nothing job-specific. |
| `docs/` | How the engine works. Start with [PARSING.md](docs/PARSING.md). |
| `jobs/` | Local only, git-ignored. The junk drawer for every job's `.kss`, BOM and QPR workbooks, zips, generated shop `.bat` scripts and `assemblies.txt`. Safe to dump files here. |

## Tools

| File | What it does |
| --- | --- |
| `tools/Parse-KSS.ps1` | The KISS engine. Reads the `.kss` files (or a soundwall post-schedule CSV) in the current folder and writes the BOM workbook, filtered `.kss` files, an optional QPR workbook, a `Folder Labels` tab (Dymo 3.5 x 1.125 in, one label per transmittal and sequence) when a Required Assm workbook is present, a `TERRA Work Order` tab (letter pages, one work order per sequence, with item list, weights and totals; paint notes come from the Required Assm TL Notes tab, else NO PAINT), and per-TL shop scripts: `Sort-PlatesByThickness` (sorts plate DXFs into thickness folders), `Type-Quantities` and `Type-Grades` (type values into ProNest). |
| `tools/verify-whitelist.ps1` | Tests an assembly whitelist on its own, using the same parser as the engine. |

## Needs

- Windows PowerShell 5.1 and Microsoft Excel (the engine drives Excel to build the workbooks).

## Options

No options are needed for a normal run. The script asks everything on screen, with colored menus: cyan frames, yellow titles, green `[X]` choices and prompts, gray examples.

By default the filtered `.kss` keeps Finish/coating exactly as received and writes the transmittal (TL) into the cost code / pay category slot (D field 13).

| Switch (optional) | Effect |
| --- | --- |
| `-TransmittalInFinish` | Old behavior: write the TL over Finish/coating (D field 10) instead. Only for reproducing old output. |
| `-StandardWeight` | BOM weight is read from KISS fields 14/15 (gross, else net, else 0) instead of the last numeric field on the line. Off by default. |

Quality Process Report: each sheet is named `{job} - {TL}`, or `{job} - {TL} - {sequence}` when sorted by sequence, and the "Job:" box on the sheet shows the same text. Excel cuts sheet names at 31 characters; the Job: box shows the full text.

## Run

1. Copy or drop the job's `.kss` files into `jobs/` (or any folder).
2. In that folder: `powershell -ExecutionPolicy Bypass -File <path to KISS>\tools\Parse-KSS.ps1`
3. Answer the prompts: data source, which files and TLs, optional folder labels (needs a `... REQUIRED ASSM.xlsx` in the
   folder, see [PARSING.md](docs/PARSING.md) section 6a), optional assembly whitelist, rev overrides,
   duplicates, shop script grouping, QPR. Details in [docs/PARSING.md](docs/PARSING.md).

Outputs are written to the folder the script runs in. The generated `.bat` files find their DXF folder
relative to themselves, so keep them with the job files.

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
| `tools/Parse-KSS.ps1` | The KISS engine. Reads the `.kss` files (or a soundwall post-schedule CSV) in the current folder and writes the BOM workbook, filtered `.kss` files, an optional QPR workbook, and per-TL shop scripts: `Sort-PlatesByThickness` (sorts plate DXFs into thickness folders), `Type-Quantities` and `Type-Grades` (type values into ProNest). |
| `tools/verify-whitelist.ps1` | Tests an assembly whitelist on its own, using the same parser as the engine. |

## Needs

- Windows PowerShell 5.1 and Microsoft Excel (the engine drives Excel to build the workbooks).

## Run

1. Copy or drop the job's `.kss` files into `jobs/` (or any folder).
2. In that folder: `powershell -ExecutionPolicy Bypass -File <path to KISS>\tools\Parse-KSS.ps1`
3. Answer the prompts: data source, which files and TLs, optional assembly whitelist, rev overrides,
   duplicates, shop script grouping, QPR. Details in [docs/PARSING.md](docs/PARSING.md).

Outputs are written to the folder the script runs in. The generated `.bat` files find their DXF folder
relative to themselves, so keep them with the job files.

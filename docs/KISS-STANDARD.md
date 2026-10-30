# KISS Standard v1.1: Every Record and Field

Source: "The Keep It Simple Steel (KISS) Standard", FabTrol Systems / SteelKiss.org, v1.1 (March 2005), a public-domain format. This page restates the spec and marks what `tools/Parse-KSS.ps1` uses today. For how the script behaves, see [PARSING.md](PARSING.md).

**Index** below is the 0-based position after splitting a line on commas, with the record letter at index 0. The spec says the record letter must be followed by a comma. Lines should be 254 characters or less, dates are `MM/DD/YY`, and any unknown record letter must be ignored by readers. Labor, Sequence, ABM and CNC lines belong to the **previous `D` line** (order among them doesn't matter).

Parser use: **Yes** = read and used, **Written** = only written to output files, **No** = ignored (an opportunity).

## Record order in a file

`KISS` id line, one `H`, optional `A` address line, then many `D` lines. Each `D` can be followed by `L`, `S`, `A` (ABM), `C` lines. In v1.1 an optional `W` (once per drawing) and `M` (once per assembly mark) follow the first `D` for that drawing/assembly. `*` comments can go anywhere except line 1.

## `KISS` identification line (mandatory, line 1)

| Idx | Field | Notes | Parser |
|---|---|---|---|
| 0 | `KISS` | | Ignored |
| 1 | Version | `1.1` in our samples | No |
| 2 | Generating software | e.g. `TeklaStructures` | No |

## `H` header (mandatory, one only)

| Idx | Field | Notes | Parser |
|---|---|---|---|
| 1 | Job # | up to 8 chars, DOS-safe | Yes |
| 2 | Job name | up to 50 | Yes |
| 3 | Customer name | up to 40, e.g. `KIEWIT/TIC` | No |
| 4 | Created date | mm/dd/yy | No |
| 5 | Created time | hh:mm:ss | No |
| 6 | Metric | `T` = metric sizes, `F` = imperial sizes | No (samples are `F`) |
| 7 | File type | Estimate, Advanced Bill, Final Bill, Revised Bill, Change Order | No (blank in samples) |

## `A` address (optional, one line; same letter as the ABM line, told apart by field count/content)

Idx 1 Description (Owner, Architect, Engineer, General, Erector, Job Site), 2 Name, 3 Mailing address, 4 city, 5 state, 6 country, 7 postal, 8 Shipping address, 9 city, 10 state, 11 country, 12 postal, 13 Phone, 14 Fax, 15 Pager, 16 E-mail, 17 Contact. **Parser: No.** None of our samples contain one.

## `D` detail (mandatory, unlimited)

| Idx | Field | Len | Req | Notes | Parser |
|---|---|---|---|---|---|
| 1 | Drawing no | 30 | Yes | In our Tekla files this is the assembly number, e.g. `52031000` | Yes, called "Assembly" |
| 2 | Drawing revision | 5 | No | `0` in samples | No |
| 3 | Assembly mark | 16 | Yes | Ship mark; the main part's mark | Yes, called "MainMark" |
| 4 | Part mark | 16 | Yes | Piece mark; blank on loose hardware | Yes |
| 5 | Quantity | 4 | Yes | **Total** quantity (parts x assemblies), not per assembly | Yes |
| 6 | Type of material | 3 | Yes | Shape code, see list below | Yes |
| 7 | Size | 24 | Yes | AISC designation; plates are `thickness X width` | Yes |
| 8 | Grade | 8 | No | If blank, A36 is assumed (A36 for steel, Grade 50 for pipe) | Yes |
| 9 | Length | 9.2 | Yes | **Millimeters** to 2 decimals | Yes, not converted |
| 10 | **Finish** | 24 | No | Paint/coating spec. Both sides must agree what the text means | Yes (this is the COATING value) |
| 11 | Notes | to 254 | No | "Beam", "Column", "Base Plate", "Curve to 24' radius", or "Shop"/"Field" on bolts | Yes, called "Description" |
| 12 | Part name | 30 | No | v1.1 | No |
| 13 | Cost code category | 6 | No | v1.1; overrides the `M` cost code for this part. Our samples show values like `8-22`, `23` | No |
| 14 | Part weight (gross, one part) | 7.2 | No | v1.1; lb for imperial, kg for metric | **No** (see below) |
| 15 | Part net weight | 7.2 | No | v1.1 | No |
| 16 | Part net area | 7.2 | No | v1.1; plates only; ft2 imperial, m2 metric | No |

**Weight (fixed):** the parser now reads field 14, else 15, else 0. It used to take the last numeric field, which could be the cost code (field 13) or the length.

### Shape codes (field 6)

AB anchor bolt, AR AR plate, C channel, CP checkered plate, CR cold rolled shafting, FB flat bar, H H-pile, HS high-strength bolt, L angle, M M-beam, MB machine bolt (also nuts and washers), MC misc channel, MI miscellaneous, PI pipe, PL plate, PV pressure vessel plate, RB round bar, S S-beam, SQ square bar, ST S-tee, T1 T-1 plate, HSS tube, UM UM plate, W wide flange, WA wedge anchor, WS weld stud, WT W-tee. Others may be defined by users (the script also treats `SHT, BPL, SHEET, CP` as plate).

### Size examples (field 7)

`L 3x3x1/4` or metric `L 75x75x6`; `AB 3/4x12` (dia x length, optional hook and thread); `C 6x8.2`; `FB 1/2x3-1/2`; `MB 3/4x2-3/4`; `HS 3/4x2-3/4` (grade A325); nuts `MB 3/4" NC NUTS`; washers `MB 3/4" FW` or `CW`; pipe `PI 3" Std/XS/Sch 40`; plate `PL 1/2 x 24`; `SQ 1/2`; tube `HSS 3x3x.250`; `W 12x40`.

## `W` drawing (v1.1, optional, once per drawing no)

Idx 1 Drawing no (30, matches the `D`), 2 Revision (5), 3 Drawing title (30), 4 Drawing date (mm/dd/yy; created date if no revision, else revised date), 5 Drawn by (8). **Parser: No.** Not present in our samples. If a detailer sends it, title and drawn-by could print on work orders.

## `M` assembly (v1.1, optional, once per assembly mark)

| Idx | Field | Len | Notes |
|---|---|---|---|
| 1 | Assembly mark | 20 | Matches the `D` |
| 2 | Assembly quantity | 6 | Total instances of the assembly |
| 3 | Assembly name | 30 | |
| 4 | Assembly type | 4 | Code |
| 5 | Assembly description | 50 | |
| 6 | Cost code category | 6 | Applies to parts without their own |
| 7 | Assembly weight (gross, single) | 7.2 | Sum of its parts |
| 8 | Assembly net weight | 7.2 | |
| 9 | Gross length | 7.2 | Bounding box, mm |
| 10 | Gross depth | 7.2 | Bounding box, mm |
| 11 | Gross width | 7.2 | Bounding box, mm |

**Parser: No.** Not in our samples. Useful for shipping weights and overall dimensions per assembly.

## `L` labor (optional, unlimited per `D`)

Idx 1 Labor type (15), 2 Quantity (6, per piece), 3 Size1 (9.2), 4 Size2 (9.2), 5 Notes (40). **Parser: Yes** (collected for the Labor tab and written back). Labor on a main mark is not summed from its minor marks. Recognized types:

| Type | Qty | Size1 | Size2 | Notes |
|---|---|---|---|---|
| Cuts | cuts per piece | unused | unused | |
| Holes | holes per piece | hole diameter, mm | material thickness | Round (default), Square, Slotted, Countersunk |
| Copes | copes per piece | web cut length | flange cut length | |
| Camber | usually 1 | camber depth | | |
| Burn | burns per piece | burn length, mm | thickness, mm | Web or Flange |
| Weld | welds per piece | weld length, mm | weld size, mm | Fillet, Flat, Full Pen |
| Other | operations per piece | thickness | length | |

Our samples contain none, but the scribe-vs-hole jobs (26102, 26123) would show up here as `Holes` if the detailer sends them.

## `S` sequence (optional, unlimited per `D`)

| Idx | Field | Notes | Parser |
|---|---|---|---|
| 1 | Sequence | up to 10 chars (FabTrol keeps 3) | Written; checked in the integrity test |
| 2 | Quantity | assembly quantity in this sequence | Written; checked |
| 3 | Lot name (v1.1, 16) | qty then means assemblies in this sequence and lot | No |

Apply to main members only. Sequence quantities should total the assembly quantity. The parser derives a 4-digit sequence from the first 4 digits of the assembly number instead of reading `S`.

## `A` ABM line (optional, unlimited per `D`)

Idx 1 ABM code (16): the Advanced Bill of Materials number tying an advance bill to the final BOM. A mark can have several. Our files hold values like `520301-02` here, and the parser reads it as the **drawing** (workbook "Drawing" column). Worth confirming with Andrew what these really are before relying on either meaning.

## `C` CNC data (optional, unlimited per `D`)

Idx 1 CNC data (254), an ASCII command string, multiple commands separated by CR. **Parser: No.** Not in our samples, but it could carry machine instructions if a detailer sends them.

## `*` comment

Free text up to 254 characters, anywhere but line 1. Skipped.

## Quick list: unused fields we could start using

- `D` 13 cost code, 14/15 weights, 16 net area (plate area could feed nesting or paint-area estimates), 12 part name.
- `M` assembly weight, name and bounding-box dimensions (shipping and loading).
- `W` drawing title, date and drawn-by.
- `H` customer, date, time, metric flag and file type (Revised Bill and Change Order identify revisions, which helps the REV logic).
- `S` lot name.
- `C` CNC strings.

## What the files we actually receive contain

Checked against `25-351` TL1282/1296/1301/1313, `26-418` TL068 and `KSS-298923` (Advance Steel) in the project folder.

| Item | Tekla (Kiewit/TIC, Crusoe/DPR) | Advance Steel (Hoffman) |
|---|---|---|
| Generator in line 1 | `TeklaStructures` / `Tekla Structures` | `AdvanceSteel` |
| `D` fields | always 16 | 17 |
| Record types present | `KISS H * D S A` only | `KISS H * D S` |
| Not present anywhere | `M W L C`, address `A` | same |
| Metric flag (H field 6) | `F` (imperial sizes) | `F` |
| Revision (D 2) | `0`, but `1` on TL1301 (revised parts) | `2` on most lines |
| Finish (D 10) | `3-1`, `0-0` (25-351); `2-9` (26-418); blank on all plates and hardware | `SEE NOTE` on every line |
| Notes (D 11) | `COLUMN(C8)`, `BASE_PLATE`, ... | blank |
| Part name (D 12) | always blank | `beam` / `plate` |
| Cost code (D 13) | `23`, `3`, `2`, `0`, `8`, `8-22`... | blank |
| Gross weight (D 14) | always blank | filled (e.g. `81.93`) |
| Net weight (D 15) | `0.0` on parts, blank on main marks | filled |
| Net area (D 16) | not present | filled |
| `S` line | 3 fields, one per assembly | 4 fields (trailing empty lot name) |
| `A` line | one per assembly after `S`, e.g. `A,520301-02` | none |

Consequences:

- **Tekla files carry no weight.** The parser now reads fields 14/15, so Tekla BOM weights show 0 (previously main marks showed the cost code, e.g. 23, 8, 17, as the weight). Advance Steel files carry real weights in fields 14 and 15.
- **Finish values are job codes**, not paint names (`3-1`, `2-9`), so what they mean depends on the customer. They are the COATING column on the QPR and must be passed through untouched.
- **Transmittal is not in the file.** It comes from the file name (`TL1313`). The filtered output now writes it to the `S` line lot name so it survives without touching Finish.

<#
.SYNOPSIS
    Professional KSS to Excel Nesting, Hardware Takeoff & Shop Management Engine.
    Includes Master BOM, PL/SHT/BPL Nesting, Profiles, Hardware & Fasteners, 
    Shop Assembly Views, Labor & Weld Aggregations, DXF Sorting Batch Generator,
    and A-Z Ordered Auto-Type Quantities Batch Generator.
#>

param(
    # Legacy behavior: write the transmittal over Finish/coating (D field 10) in the filtered .kss.
    # Default: Finish stays as received and the transmittal goes in the cost code / pay category slot (D field 13).
    [switch]$TransmittalInFinish,
    # Read part weight from KISS fields 14 (gross) / 15 (net) instead of the last numeric field on the line.
    [switch]$StandardWeight
)

$ErrorActionPreference = "Stop"

# --- COLORED CONSOLE OUTPUT ---
# Menus and prompts are colored so a new operator can scan them: cyan frames, yellow titles,
# green [X] choices and prompts, gray examples. Calls that pass their own color are left alone.
$script:FrameState = 0   # 0 = none, 1 = just printed an opening ==== line, 2 = just printed the title
function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]$Object,
        [switch]$NoNewline,
        $Separator = ' ',
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    if ($PSBoundParameters.ContainsKey('ForegroundColor') -or $PSBoundParameters.ContainsKey('BackgroundColor') -or $NoNewline) {
        Microsoft.PowerShell.Utility\Write-Host @PSBoundParameters
        return
    }
    $text = ($Object | ForEach-Object { "$_" }) -join $Separator
    $state = $script:FrameState
    $script:FrameState = 0
    if ($text -match '^=+$') {
        if ($state -eq 0) { $script:FrameState = 1 }
        Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor Cyan
    }
    elseif ($state -eq 1 -and $text.Trim()) {
        $script:FrameState = 2
        Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor Yellow
    }
    elseif ($text -match '^(\s*)(\[[^\]]+\]\w*)(.*)$') {
        Microsoft.PowerShell.Utility\Write-Host $Matches[1] -NoNewline
        Microsoft.PowerShell.Utility\Write-Host $Matches[2] -ForegroundColor Green -NoNewline
        $rest = $Matches[3]
        if ($rest -match '^(.*?)(\(default[^)]*\))(.*)$') {
            Microsoft.PowerShell.Utility\Write-Host $Matches[1] -NoNewline
            Microsoft.PowerShell.Utility\Write-Host $Matches[2] -ForegroundColor Yellow -NoNewline
            Microsoft.PowerShell.Utility\Write-Host $Matches[3]
        }
        else { Microsoft.PowerShell.Utility\Write-Host $rest }
    }
    elseif ($text -match '^\s*(Example|Files|Sheets|Note)\b') {
        Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor DarkGray
    }
    elseif ($text -match '^\s*(Use for|-\s)') {
        Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor Gray
    }
    elseif ($text -match '^\s*(WITH |Simple list)') {
        Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor Yellow
    }
    else {
        Microsoft.PowerShell.Utility\Write-Host $text
    }
}

function Read-Host {
    param([Parameter(Position = 0)][string]$Prompt)
    if ($Prompt) { Microsoft.PowerShell.Utility\Write-Host "${Prompt}: " -ForegroundColor Green -NoNewline }
    Microsoft.PowerShell.Utility\Read-Host
}

# --- HELPER FUNCTIONS ---
function Write-Log ($Message, $Color = "Cyan") {
    $timestamp = Get-Date -Format "HH:mm:ss.fff"
    Write-Host "[$timestamp] $Message" -ForegroundColor $Color
}

function Get-OleColor ($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

function Get-ExcelColumnLetter ($columnNumber) {
    $letters = ""
    $value = [int]$columnNumber

    while ($value -gt 0) {
        $value--
        $remainder = $value % 26
        $letters = [char][int](65 + $remainder) + $letters
        $value = [Math]::Floor($value / 26)
    }

    return $letters
}

function Get-DecimalThickness ($sizeStr) {
    if (-not $sizeStr) { return 0.0 }
    $cleanSize = $sizeStr.Trim().ToUpper()

    # Check for Gauge format (e.g., 10GA, 14GA)
    if ($cleanSize -match '^(\d+)\s*GA') {
        $gaNum = $Matches[1]
        switch ($gaNum) {
            '3'  { return 0.2391 }
            '4'  { return 0.2242 }
            '5'  { return 0.2092 }
            '6'  { return 0.1943 }
            '7'  { return 0.1793 }
            '8'  { return 0.1644 }
            '9'  { return 0.1495 }
            '10' { return 0.1345 }
            '11' { return 0.1196 }
            '12' { return 0.1046 }
            '13' { return 0.0897 }
            '14' { return 0.0747 }
            '15' { return 0.0673 }
            '16' { return 0.0598 }
            '18' { return 0.0478 }
            '20' { return 0.0359 }
            '22' { return 0.0299 }
            '24' { return 0.0239 }
            '26' { return 0.0179 }
            default { return 0.0 }
        }
    }

    $token = if ($cleanSize -match '^[xX]') { $cleanSize } else { ($cleanSize -split '[xX]')[0].Trim() }
    $token = $token.Trim()
    if (-not $token) { return 0.0 }

    $val = 0.0
    if ([decimal]::TryParse($token, [ref]$val)) {
        return $val
    }

    if ($token -match '^(\d+)\s*(?:-|\s+)\s*(\d+)/(\d+)$') {
        $whole = [decimal]$Matches[1]
        $num = [decimal]$Matches[2]
        $den = [decimal]$Matches[3]
        if ($den -ne 0) { return $whole + ($num / $den) }
    }
    elseif ($token -match '^(\d+)\s+(\d+)/(\d+)$') {
        $whole = [decimal]$Matches[1]
        $num = [decimal]$Matches[2]
        $den = [decimal]$Matches[3]
        if ($den -ne 0) { return $whole + ($num / $den) }
    }
    elseif ($token -match '^(\d+)/(\d+)$') {
        $num = [decimal]$Matches[1]
        $den = [decimal]$Matches[2]
        if ($den -ne 0) { return $num / $den }
    }

    return 0.0
}

function Format-DecimalThickness ($dec) {
    if ($dec -eq 0) { return "0" }
    $str = $dec.ToString("0.###")
    if ($str.StartsWith("0.")) {
        $str = $str.Substring(1)
    }
    return $str
}

function Format-ThicknessFolder ($dec) {
    if ($dec -eq 0) { return "0" }
    return $dec.ToString("0.###")
}

function Escape-BatchText ($text) {
    if ($null -eq $text) { return "" }
    $value = [string]$text
    return ($value -replace '([&|<>\^%])', '^$1')
}

function ConvertTo-SendKeysText ($text) {
    if ($null -eq $text) { return "" }
    $value = [string]$text
    $result = [System.Text.StringBuilder]::new()

    foreach ($ch in $value.ToCharArray()) {
        switch ($ch) {
            '+' { [void]$result.Append('{+}') }
            '^' { [void]$result.Append('{^}') }
            '%' { [void]$result.Append('{%}') }
            '~' { [void]$result.Append('{~}') }
            '(' { [void]$result.Append('{(}') }
            ')' { [void]$result.Append('{)}') }
            '{' { [void]$result.Append('{{}') }
            '}' { [void]$result.Append('{}}') }
            default { [void]$result.Append($ch) }
        }
    }

    return $result.ToString().Replace("'", "''")
}

function Get-MaterialCategory ($shape, $partMark, $description, $size) {
    $s = if ($shape) { $shape.Trim().ToUpper() } else { "" }
    $d = if ($description) { $description.Trim().ToUpper() } else { "" }
    $m = if ($partMark) { $partMark.Trim().ToLower() } else { "" }
    $sz = if ($size) { $size.Trim().ToUpper() } else { "" }

    # 1. Hardware & Fasteners Detection
    if ($s -match '^(HS|MB|AB|WA|WS|NUT|WASHER|BOLT)$' -or 
        $d -match '(NUT|WASHER|BOLT|ANCHOR|ROD|STUD|FIELD|SHOP)' -or
        $m -match '^(nut|washer|bolt|anchor)' -or
        $sz -match '(NUT|WASHER|HEAVY_HEX)') {
        return "Hardware"
    }

    # 2. Plate & Sheetmetal Detection
    if ($s -match '^(PL|SHT|BPL|PLATE|SHEET|CP|AR|UM|T1)$') {
        return "Plate"
    }

    # 3. Rolled Structural Profiles
    return "Profile"
}


# --- SOUNDWALL CSV IMPORT HELPERS ---
function ConvertTo-DecimalFeet ($str) {
    if (-not $str) { return 0.0 }
    $s = $str.Trim().Replace('"', '')
    if ($s -match "^(\d+)'-(\d+(?:\s+\d+/\d+)?)$") {
        $feet = [decimal]$Matches[1]
        $inchPart = $Matches[2]
        $inches = 0.0
        if ($inchPart -match '^(\d+)\s+(\d+)/(\d+)$') {
            $inches = [decimal]$Matches[1] + ([decimal]$Matches[2] / [decimal]$Matches[3])
        }
        else {
            [decimal]::TryParse($inchPart, [ref]$inches) | Out-Null
        }
        return $feet + ($inches / 12)
    }

    $val = 0.0
    if ([decimal]::TryParse($s, [ref]$val)) {
        if ($val -gt 100) {
            # KISS exports most raw structural lengths in millimeters, not feet.
            return $val / 304.8
        }
        return $val / 12.0
    }

    return 0.0
}

function Get-ShapePrefix ($sizeStr) {
    if ($sizeStr -match '^([A-Za-z]+)') { return $Matches[1].ToUpper() }
    return "PROFILE"
}

function Get-WShapeUnitWeight ($sizeStr, $lengthStr) {
    # AISC designations encode nominal weight/ft after the "X" (e.g. W8X24 = 24 lbs/ft)
    $wtPerFt = 0.0
    if ($sizeStr -match '[Xx](\d+(?:\.\d+)?)\s*$') {
        $wtPerFt = [decimal]$Matches[1]
    }
    $lengthFt = ConvertTo-DecimalFeet $lengthStr
    return [Math]::Round($wtPerFt * $lengthFt, 2)
}

function Get-QtyAndSize ($str) {
    $s = $str.Trim()
    $qty = 1
    if ($s -match '^\((\d+)\)\s*(.+)$') {
        $qty = [int]$Matches[1]
        $s = $Matches[2].Trim()
    }
    return [PSCustomObject]@{ Qty = $qty; Size = $s }
}

function Get-FlatPlateUnitWeight ($sizeStr, $lengthStr) {
    $clean = $sizeStr -replace '"', ''
    $tokens = $clean -split '[xX]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if ($tokens.Count -lt 2) { return 0.0 }
    $thick = Get-DecimalThickness $tokens[0]
    $width = Get-DecimalThickness $tokens[1]
    $lengthIn = if ($tokens.Count -ge 3) { Get-DecimalThickness $tokens[2] } else { (ConvertTo-DecimalFeet $lengthStr) * 12 }
    $steelDensity = 0.2836 # lbs per cubic inch
    return [Math]::Round($thick * $width * $lengthIn * $steelDensity, 2)
}

$script:PartMarkRegistry = @{}
function Get-OrCreatePartMark ($prefix, $key) {
    if (-not $script:PartMarkRegistry.ContainsKey($prefix)) {
        $script:PartMarkRegistry[$prefix] = @{ Counter = 0; Map = @{} }
    }
    $reg = $script:PartMarkRegistry[$prefix]
    if ($reg.Map.ContainsKey($key)) { return $reg.Map[$key] }
    $reg.Counter++
    $mark = "$prefix$($reg.Counter)"
    $reg.Map[$key] = $mark
    return $mark
}

function Select-CsvFile {
    $csvFiles = @(Get-ChildItem -Filter "*.csv" -File | Sort-Object Name)
    Write-Host ""
    Write-Host "========================================================"
    Write-Host "  Select a Soundwall CSV File"
    Write-Host "========================================================"
    $i = 1
    foreach ($f in $csvFiles) {
        Write-Host " [$i] $($f.Name)"
        $i++
    }
    $browseOption = $i
    Write-Host " [$browseOption] Browse for a CSV file..."
    Write-Host ""

    $choice = Read-Host "Enter selection"
    $choiceNum = 0
    if (-not [int]::TryParse($choice, [ref]$choiceNum)) { return $null }

    if ($choiceNum -eq $browseOption) {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $dialog.InitialDirectory = (Get-Location).Path
        $dialog.Title = "Select Soundwall CSV File"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.FileName
        }
        return $null
    }
    elseif ($choiceNum -ge 1 -and $choiceNum -le $csvFiles.Count) {
        return $csvFiles[$choiceNum - 1].FullName
    }
    return $null
}

# --- TRANSMITTAL (TL) HELPERS & KSS FILE PICKER ---
function Get-TransmittalFromName ($name) {
    if ($name -match '(?i)(?<![A-Za-z])TL[\s_-]?(\d+)') { return "TL$($Matches[1])" }
    return ""
}

function ConvertTo-TransmittalTag ($text) {
    if (-not $text) { return "" }
    $t = $text.Trim().ToUpper()
    if ($t -match '^TL[\s_-]?(\d+)$') { return "TL$($Matches[1])" }
    if ($t -match '^\d+$') { return "TL$t" }
    return ($t -replace '\s', '')
}

# --- FOLDER LABELS (from the "Required Assm" workbook) ---
# Looks for a "... REQUIRED ASSM.xlsx" in the current folder. Its Sheet1 has one column per transmittal:
# TL name in the "TL Name" row, the AFCO assembly numbers underneath, and optional "Paint / TL Details" and
# "Due Date" rows above. A REV transmittal is named TLR####. Each transmittal + sequence (folder) becomes one 3.5" x 1.125" label.
function Select-RequiredAssmFile {
    $files = @(Get-ChildItem -LiteralPath (Get-Location).Path -Filter *.xlsx -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '~$*' -and $_.Name -match '(?i)required[\s_-]*assm' } |
        Sort-Object Name)
    if ($files.Count -eq 0) { return $null }

    Write-Host ""
    Write-Host "========================================================"
    Write-Host "  FOLDER LABELS (Required Assm workbook)"
    Write-Host "========================================================"
    if ($files.Count -eq 1) {
        Write-Host "Files: $($files[0].Name)"
        Write-Host ""
        Write-Host "  [Y]es (default): add a 'Folder Labels' tab to the BOM, one 3.5x1.125 label per transmittal and sequence"
        Write-Host "  [N]o: skip the labels"
        $choice = Read-Host "Build folder labels from this file? [Y]es / [N]o (default: Yes)"
        if ($choice -match '^(n|no)$') { return $null }
        return $files[0].FullName
    }

    Write-Host "Found $($files.Count) Required Assm workbooks:"
    for ($i = 0; $i -lt $files.Count; $i++) { Write-Host "  [$($i + 1)] $($files[$i].Name)" }
    Write-Host "  [0] Skip the folder labels"
    $pick = Read-Host "Which file? (default: 1)"
    if ($pick -match '^\s*0\s*$|^(n|no)$') { return $null }
    $n = 1
    if ($pick -match '^\s*(\d+)\s*$') { $n = [int]$Matches[1] }
    if ($n -lt 1 -or $n -gt $files.Count) { $n = 1 }
    return $files[$n - 1].FullName
}

function Read-RequiredAssm {
    param($Excel, [string]$Path)

    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        foreach ($ws in $wb.Worksheets) {
            $last = $ws.Cells.SpecialCells(11)   # xlCellTypeLastCell
            $vals = $ws.Range($ws.Cells.Item(1, 1), $last).Value2
            if ($vals -isnot [System.Array] -or $vals.Rank -ne 2) { continue }
            $nRows = $vals.GetLength(0)
            $nCols = $vals.GetLength(1)

            # Header row: "TL Name" in column A
            $hdr = 0
            for ($r = 1; $r -le $nRows; $r++) {
                if (("$($vals[$r, 1])").Trim() -match '^(?i)TL\s*Name$') { $hdr = $r; break }
            }
            if ($hdr -eq 0) { continue }

            # Labeled rows above the header
            $paintRow = 0; $dueRow = 0
            for ($r = 1; $r -lt $hdr; $r++) {
                $label = ("$($vals[$r, 1])").Trim()
                if ($label -match '(?i)^paint|coating') { $paintRow = $r }
                elseif ($label -match '(?i)^due') { $dueRow = $r }
            }

            for ($c = 2; $c -le $nCols; $c++) {
                $name = ("$($vals[$hdr, $c])").Trim().ToUpper() -replace '\s', ''
                if ($name -notmatch '^TLR?\d+$') { continue }

                $seqs = [System.Collections.Generic.SortedSet[string]]::new()
                $seqAssm = @{}                                    # sequence -> assembly numbers (digits only) in this column
                $count = 0
                for ($r = $hdr + 1; $r -le $nRows; $r++) {
                    $v = $vals[$r, $c]
                    if ($null -eq $v) { continue }
                    $digits = ("$v").Trim() -replace '\D', ''     # drops an R (REV) prefix and anything else
                    if ($digits.Length -lt 4) { continue }
                    $count++
                    [void]$seqs.Add($digits.Substring(0, 4))
                    if (-not $seqAssm.ContainsKey($digits.Substring(0, 4))) { $seqAssm[$digits.Substring(0, 4)] = [System.Collections.Generic.HashSet[string]]::new() }
                    [void]$seqAssm[$digits.Substring(0, 4)].Add($digits)
                }
                if ($count -eq 0) { continue }                    # unused column

                $coating = ""
                if ($paintRow -gt 0 -and $null -ne $vals[$paintRow, $c]) { $coating = ("$($vals[$paintRow, $c])").Trim() }
                $due = ""
                if ($dueRow -gt 0 -and $null -ne $vals[$dueRow, $c]) {
                    $d = $vals[$dueRow, $c]
                    if ($d -is [double]) { $due = [DateTime]::FromOADate($d).ToString('M/d/yy') }
                    else { $due = ("$d").Trim() }
                }

                $result.Add([PSCustomObject]@{
                    TL        = $name
                    Number    = [int](($name -replace '\D', ''))
                    IsRev     = ($name -like 'TLR*')
                    Coating   = $coating
                    DueDate   = $due
                    Sequences = @($seqs)
                    SeqAssemblies = $seqAssm
                    Count     = $count
                })
            }
            break   # first sheet that has a "TL Name" row
        }
    }
    finally {
        $wb.Close($false)
    }
    # Per-sequence details from the "TL Notes" tab (register keyed by 4-digit sequence). Used first; the
    # Sheet1 rows above the TL header are only a fallback. The tab holds a formula-spilled copy of the same
    # table next to the yellow input table, so the first non-blank value per sequence wins.
    $seqNotes = @{}
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name -notmatch '(?i)^\s*TL\s*Notes\s*$') { continue }
            $last = $ws.Cells.SpecialCells(11)
            $vals = $ws.Range($ws.Cells.Item(1, 1), $last).Value2
            if ($vals -isnot [System.Array] -or $vals.Rank -ne 2) { break }
            $nRows = $vals.GetLength(0)
            $nCols = $vals.GetLength(1)
            for ($hr = 1; $hr -le $nRows; $hr++) {
                for ($hc = 1; $hc -le $nCols; $hc++) {
                    if (("$($vals[$hr, $hc])").Trim() -notmatch '^(?i)sequence$') { continue }
                    $paintCol = 0; $dueCol = 0
                    for ($k = $hc + 1; $k -le [Math]::Min($nCols, $hc + 6); $k++) {
                        $h = ("$($vals[$hr, $k])").Trim()
                        if ($h -match '(?i)^paint|coating') { $paintCol = $k }
                        elseif ($h -match '(?i)^due') { $dueCol = $k }
                    }
                    if ($paintCol -eq 0 -and $dueCol -eq 0) { continue }
                    for ($r = $hr + 1; $r -le $nRows; $r++) {
                        $seq = ("$($vals[$r, $hc])").Trim()
                        if ($seq -notmatch '^\d{4}$') { continue }
                        if (-not $seqNotes.ContainsKey($seq)) { $seqNotes[$seq] = @{ Coating = ""; DueDate = "" } }
                        if ($paintCol -gt 0 -and -not $seqNotes[$seq].Coating -and $null -ne $vals[$r, $paintCol]) {
                            $seqNotes[$seq].Coating = ("$($vals[$r, $paintCol])").Trim()
                        }
                        if ($dueCol -gt 0 -and -not $seqNotes[$seq].DueDate -and $null -ne $vals[$r, $dueCol]) {
                            $d = $vals[$r, $dueCol]
                            $seqNotes[$seq].DueDate = if ($d -is [double]) { [DateTime]::FromOADate($d).ToString('M/d/yy') } else { ("$d").Trim() }
                        }
                    }
                }
            }
            break
        }
    }
    finally {
        $wb.Close($false)
    }
    foreach ($tlEntry in $result) { $tlEntry | Add-Member -NotePropertyName SeqNotes -NotePropertyValue $seqNotes }
    return @($result | Sort-Object Number)
}

# Physical label: Dymo address label, 3.5" x 1.125" = 252 pt x 81 pt. Rows 40.5 + 15 + 25.5 = 81 pt.
# Margins are 0 and the page is set to fit all columns on one page, so the text stays inside the printable area.
function Add-FolderLabelSheet {
    param($Workbook, $Labels, [string]$OurJob, [string]$Title, [string]$CustomerJob, [string]$IssuedBy)

    $IssuedLabel = "IS:"   # text printed before the initials on line 2 (alone if the initials are blank)

    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "Folder Labels"
    $sheet.Cells.NumberFormat = "@"
    $sheet.Cells.Font.Name = "Arial"
    $sheet.Cells.Font.Bold = $true

    $setWidth = {
        param($col, [double]$pts)
        $w = $pts / 5.4
        for ($k = 0; $k -lt 8; $k++) {
            $col.ColumnWidth = $w
            $cur = [double]$col.Width
            if ([Math]::Abs($cur - $pts) -lt 0.8) { break }
            $w = $w * ($pts / $cur)
        }
    }
    & $setWidth $sheet.Columns.Item(1) 132
    & $setWidth $sheet.Columns.Item(2) 120

    # Alignment constants: xlLeft = -4131, xlRight = -4152, xlCenter = -4108
    $i = 0
    foreach ($label in $Labels) {
        $r1 = 3 * $i + 1; $r2 = $r1 + 1; $r3 = $r1 + 2

        $sheet.Rows.Item($r1).RowHeight = 40.5
        $sheet.Rows.Item($r2).RowHeight = 15
        $sheet.Rows.Item($r3).RowHeight = 25.5

        # Line 1: our job - transmittal - sequence (big)
        $jobParts = @($OurJob, $label.TL, $label.Sequence) | Where-Object { $_ }
        $cell1 = $sheet.Range($sheet.Cells.Item($r1, 1), $sheet.Cells.Item($r1, 2))
        $cell1.Merge()
        $line1 = ($jobParts -join " - ")
        $cell1.Value2 = $line1
        $cell1.Font.Size = 26
        if ($label.IsRev) {                      # REV transmittal: the R in TLR#### is italic
            $rPos = $line1.IndexOf($label.TL) + 3
            $cell1.Characters($rPos, 1).Font.Italic = $true
        }
        $cell1.HorizontalAlignment = -4108
        $cell1.VerticalAlignment = -4108
        $cell1.ShrinkToFit = $true

        # Line 2 (small): project title - customer job on the left, "ISU'D: initials" on the right
        $cellL = $sheet.Cells.Item($r2, 1)
        $cellL.Value2 = (@($Title, $CustomerJob) | Where-Object { $_ }) -join " - "
        $cellL.Font.Size = 10
        $cellL.HorizontalAlignment = -4131
        $cellL.VerticalAlignment = -4108
        $cellL.ShrinkToFit = $true
        $cellR = $sheet.Cells.Item($r2, 2)
        $cellR.Value2 = (@($IssuedLabel, $IssuedBy) | Where-Object { $_ }) -join " "
        $cellR.Font.Size = 10
        $cellR.HorizontalAlignment = -4152
        $cellR.VerticalAlignment = -4108
        $cellR.ShrinkToFit = $true

        # Line 3: comments / paint info on the left, due date on the right (blank date = slashes to fill in later)
        $cellC = $sheet.Cells.Item($r3, 1)
        $cellC.Value2 = $label.Notes
        $cellC.Font.Size = 11
        $cellC.HorizontalAlignment = -4131
        $cellC.VerticalAlignment = -4108
        $cellC.ShrinkToFit = $true
        $cellD = $sheet.Cells.Item($r3, 2)
        $cellD.Value2 = if ($label.DueDate) { "DUE $($label.DueDate)" } else { "DUE    /    /" }
        $cellD.Font.Size = 18
        $cellD.HorizontalAlignment = -4152
        $cellD.VerticalAlignment = -4108
        $cellD.ShrinkToFit = $true

        $i++
    }

    $lastRow = 3 * $Labels.Count
    $sheet.Range("D1").Value2 = "Print this tab on the Dymo (3.5 x 1.125 in address label). One label per page."
    $sheet.Range("D1").Font.Bold = $false
    $sheet.Range("D1").Font.Italic = $true

    try {
        $ps = $sheet.PageSetup
        $ps.PrintArea = "`$A`$1:`$B`$$lastRow"
        $ps.Orientation = 2                      # landscape: the long side (3.5 in) is the width
        $ps.LeftMargin = 0; $ps.RightMargin = 0
        $ps.TopMargin = 0; $ps.BottomMargin = 0
        $ps.HeaderMargin = 0; $ps.FooterMargin = 0
        $ps.CenterHorizontally = $true
        $ps.CenterVertically = $true
        $sheet.ResetAllPageBreaks()
        for ($b = 1; $b -lt $Labels.Count; $b++) {
            [void]$sheet.HPageBreaks.Add($sheet.Cells.Item(3 * $b + 1, 1))
        }
        $ps.Zoom = $false                        # after ResetAllPageBreaks (it resets the scaling): fit all columns on one page
        $ps.FitToPagesWide = 1
        $ps.FitToPagesTall = $false
    }
    catch {
        Write-Log "Folder Labels tab built, but page setup failed (is a printer installed? select the Dymo and set the page to 3.5 x 1.125 in): $($_.Exception.Message)" "Yellow"
    }
    return $sheet
}

function Read-KssSelectionInteractive ($Entries) {
    $cursor = 0
    $offset = 0
    $message = ""
    $pageSize = [Math]::Max(5, [Console]::WindowHeight - 9)
    $cursorWasVisible = $true
    try { $cursorWasVisible = [Console]::CursorVisible; [Console]::CursorVisible = $false } catch { }

    try {
        while ($true) {
            if ($cursor -lt $offset) { $offset = $cursor }
            if ($cursor -ge $offset + $pageSize) { $offset = $cursor - $pageSize + 1 }
            $last = [Math]::Min($Entries.Count, $offset + $pageSize) - 1
            $checkedCount = @($Entries | Where-Object { $_.Checked }).Count

            Clear-Host
            Write-Host "========================================================"
            Write-Host "  Select KSS Files (this folder)"
            Write-Host "========================================================"
            Write-Host "  Up/Down: move   Space: check   A: all   N: none   Enter: confirm   Esc: cancel" -ForegroundColor DarkGray
            Write-Host ""
            for ($i = $offset; $i -le $last; $i++) {
                $e = $Entries[$i]
                $box = if ($e.Checked) { "[x]" } else { "[ ]" }
                $pointer = if ($i -eq $cursor) { ">" } else { " " }
                $tlText = if ($e.TL) { $e.TL } else { "(no TL found)" }
                $text = " $pointer $box $($e.Name)   ->  $tlText"
                if ($i -eq $cursor) { Write-Host $text -ForegroundColor Black -BackgroundColor Gray }
                elseif ($e.Checked) { Write-Host $text -ForegroundColor Green }
                else { Write-Host $text }
            }
            Write-Host ""
            Write-Host "  Checked: $checkedCount of $($Entries.Count)   (showing $($offset + 1)-$($last + 1))"
            if ($message) { Write-Host "  $message" -ForegroundColor Yellow; $message = "" }

            $key = [Console]::ReadKey($true)
            switch ($key.Key) {
                'UpArrow'   { $cursor = if ($cursor -gt 0) { $cursor - 1 } else { $Entries.Count - 1 } }
                'DownArrow' { $cursor = if ($cursor -lt $Entries.Count - 1) { $cursor + 1 } else { 0 } }
                'Home'      { $cursor = 0 }
                'End'       { $cursor = $Entries.Count - 1 }
                'Spacebar'  { $Entries[$cursor].Checked = -not $Entries[$cursor].Checked }
                'A'         { foreach ($e in $Entries) { $e.Checked = $true } }
                'N'         { foreach ($e in $Entries) { $e.Checked = $false } }
                'Escape'    { return @() }
                'Enter'     {
                    if ($checkedCount -gt 0) { return @($Entries | Where-Object { $_.Checked }) }
                    $message = "Check at least one file (Space), then press Enter."
                }
            }
        }
    }
    finally {
        try { [Console]::CursorVisible = $cursorWasVisible } catch { }
        Clear-Host
    }
}

function Read-KssSelectionNumbered ($Entries) {
    Write-Host ""
    Write-Host "========================================================"
    Write-Host "  Select KSS Files (this folder)"
    Write-Host "========================================================"
    for ($i = 0; $i -lt $Entries.Count; $i++) {
        $tlText = if ($Entries[$i].TL) { $Entries[$i].TL } else { "(no TL found)" }
        Write-Host (" [{0}] {1}   ->  {2}" -f ($i + 1), $Entries[$i].Name, $tlText)
    }
    Write-Host ""
    $choice = Read-Host "Numbers separated by commas (e.g. 1,3), or A for all"
    if (-not $choice) { return @() }
    $trimmed = $choice.Trim()
    if ($trimmed -match '^(A|ALL)$') { return @($Entries) }

    $picked = [System.Collections.Generic.List[object]]::new()
    foreach ($tok in ($trimmed -split '[,\s]+')) {
        $num = 0
        if ([int]::TryParse($tok, [ref]$num) -and $num -ge 1 -and $num -le $Entries.Count -and -not $picked.Contains($Entries[$num - 1])) {
            $picked.Add($Entries[$num - 1])
        }
    }
    return @($picked)
}

function Select-KssFiles {
    $files = @(Get-ChildItem -Filter "*.kss" -File |
        Where-Object {
            $_.Name -notmatch ' - Filtered Assemblies' -and
            $_.Name -notmatch ' \(Generated\)' -and
            $_.Name -notmatch ' - Filtered Assemblies - VALID'
        } |
        Sort-Object Name)
    if ($files.Count -eq 0) { return @() }

    $entries = @(foreach ($f in $files) {
        [PSCustomObject]@{ FullName = $f.FullName; Name = $f.Name; TL = (Get-TransmittalFromName $f.Name); Checked = $false }
    })

    $interactive = $false
    try { $interactive = -not ([Console]::IsInputRedirected -or [Console]::IsOutputRedirected) -and [Console]::WindowHeight -gt 12 } catch { $interactive = $false }

    $picked = if ($interactive) { @(Read-KssSelectionInteractive $entries) } else { @(Read-KssSelectionNumbered $entries) }
    if ($picked.Count -eq 0) { return @() }

    foreach ($e in $picked) {
        if (-not $e.TL) {
            $typed = Read-Host "No TL found in '$($e.Name)'. Enter TL (e.g. TL012 or 12), or press Enter to leave blank"
            $e.TL = ConvertTo-TransmittalTag $typed
        }
    }

    Write-Host ""
    foreach ($e in $picked) { Write-Log "  $("$($e.TL)".PadRight(8)) $($e.Name)" "Green" }
    return $picked
}

function Parse-AssemblyWhitelist {
    param([string]$InputText)

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        return [PSCustomObject]@{
            Names   = @()
            TlMap   = @{}
            QtyMap  = @{}
            RevMap  = @{}
            HasTlMap = $false
        }
    }

    $raw = $InputText.Trim()
    if ($raw.StartsWith('@')) {
        $path = $raw.Substring(1).Trim()
        if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path (Get-Location) $path }
        if (Test-Path $path) {
            $raw = [System.IO.File]::ReadAllText($path)
        }
        else {
            Write-Log "Assembly whitelist file not found: $path" "Red"
            return [PSCustomObject]@{
                Names   = @()
                TlMap   = @{}
                QtyMap  = @{}
                RevMap  = @{}
                HasTlMap = $false
            }
        }
    }

    $names = [System.Collections.Generic.List[string]]::new()
    $tlMap = @{}
    $qtyMap = @{}
    $revMap = @{}
    $hasTlMap = $false

    # Split into lines to inspect whether TL prefixes exist per line
    $lines = $raw -split "(?:\r?\n)+"
    $currentLineTl = ""

    foreach ($line in $lines) {
        # Glue "NAME --- QTY" so the NAME---QTY pair survives the whitespace split below
        $trimmedLine = ($line -replace '\s*-{3}\s*', '---').Trim()
        if (-not $trimmedLine) { continue }

        # Split line by tab/comma/semicolon/pipe/whitespace
        $tokens = @($trimmedLine -split '(?:[,;|]|\s+)' | Where-Object { $_ -and $_.Trim() })
        if ($tokens.Count -eq 0) { continue }

        $lineTl = ""
        $firstToken = ($tokens[0].Trim() -replace '^[\s"''\[\]\(\)\{\}]+|[\s"''\[\]\(\)\{\}]+$', '').Trim()
        $firstTokenTl = ConvertTo-TransmittalTag $firstToken

        # If the first token on this line looks like a TL tag (e.g. TL012, TL17, TL-033) or prefixed order (e.g. 1_TL012, 2_TL3)
        if ($firstToken -match '^(?i)(?:\d+_)?TL[\s_-]?\d+$') {
            $lineTl = $firstTokenTl
            $tokens = @($tokens | Select-Object -Skip 1)
            $hasTlMap = $true
        }

        foreach ($token in $tokens) {
            $trimmed = ($token.Trim() -replace '^[\s"''\[\]\(\)\{\}]+|[\s"''\[\]\(\)\{\}]+$', '').Trim()
            if (-not $trimmed) { continue }

            # Check if this token itself is a TL marker starting a new section
            if ($trimmed -match '^(?i)(?:\d+_)?TL[\s_-]?\d+$') {
                $lineTl = ConvertTo-TransmittalTag ($trimmed -replace '^\d+_', '')
                $hasTlMap = $true
                continue
            }

            # Check for REV marker (R prefix on assembly name)
            $isRev = $false
            if ($trimmed -match '^R(.+)$') {
                $isRev = $true
                $trimmed = $Matches[1].Trim()
            }

            # NAME---QTY: quantity of this assembly wanted
            if ($trimmed -match '^(.+?)---(\d+)$') {
                $trimmed = $Matches[1].Trim()
                $qtyMap[$trimmed] = [int]$Matches[2]
            }

            if (-not $names.Contains($trimmed)) {
                $names.Add($trimmed)
            }

            if ($lineTl) {
                $tlMap[$trimmed] = $lineTl
            }

            if ($isRev) {
                $revMap[$trimmed] = $true
            }
        }
    }

    return [PSCustomObject]@{
        Names   = @($names)
        TlMap   = $tlMap
        QtyMap  = $qtyMap
        RevMap  = $revMap
        HasTlMap = $hasTlMap
    }
}

function Apply-AssemblyWhitelist {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$MasterData,
        [System.Collections.Generic.List[PSCustomObject]]$LaborData,
        [hashtable]$AssemblyDrawings,
        [string]$AssemblyListText
    )

    $parsed = Parse-AssemblyWhitelist -InputText $AssemblyListText
    $names = @($parsed.Names)
    $tlMap = $parsed.TlMap
    $qtyMap = $parsed.QtyMap
    $revMap = $parsed.RevMap
    $hasTlMap = $parsed.HasTlMap

    if ($names.Count -eq 0) {
        return [PSCustomObject]@{
            Applied = $false
            MasterData = $MasterData
            LaborData = $LaborData
            AssemblyDrawings = $AssemblyDrawings
            Included = @()
            Missing = @()
            TlMap = @{}
            RevMap = @{}
        }
    }

    $allowed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $names) { $allowed.Add($name.Trim()) | Out-Null }

    $available = @($MasterData | ForEach-Object { $_.Assembly } | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
    $missing = @()
    foreach ($name in $names) {
        $match = @($available | Where-Object { $_ -ieq $name })
        if ($match.Count -eq 0) { $missing += $name }
    }

    # Scale factor per assembly+file = requested qty / qty of the assembly's main part (the D row whose part mark is the main mark)
    $qtyFactor = @{}
    if ($qtyMap.Count -gt 0) {
        $reqQty = @{}
        foreach ($k in $qtyMap.Keys) { $reqQty[$k.Trim().ToLowerInvariant()] = $qtyMap[$k] }
        foreach ($g in ($MasterData | Group-Object { "$($_.Assembly.Trim())|$($_.SourceFile)".ToLowerInvariant() })) {
            $assmKey = $g.Group[0].Assembly.Trim().ToLowerInvariant()
            if (-not $reqQty.ContainsKey($assmKey)) { continue }
            $mainQty = ($g.Group | Where-Object { $_.RawPartMark -and $_.RawPartMark -eq $_.MainMark } | Select-Object -First 1).Qty
            if ($mainQty -gt 0) {
                $qtyFactor[$g.Name] = $reqQty[$assmKey] / $mainQty
            }
            else {
                Write-Log "Assembly '$($g.Group[0].Assembly)': no main part row found, requested qty $($reqQty[$assmKey]) not applied." "Yellow"
            }
        }
    }

    $filteredMaster = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $MasterData) {
        $assmKey = $item.Assembly.Trim()
        if ($allowed.Contains($assmKey)) {
            $fKey = "$assmKey|$($item.SourceFile)".ToLowerInvariant()
            if ($qtyFactor.ContainsKey($fKey)) { $item.Qty = [int][Math]::Round($item.Qty * $qtyFactor[$fKey]) }
            if ($hasTlMap) {
                if ($tlMap.ContainsKey($assmKey)) {
                    $item.TL = $tlMap[$assmKey]
                    if ($TransmittalInFinish) { $item.Finish = $tlMap[$assmKey] }
                } else {
                    $item.TL = ""
                    if ($TransmittalInFinish) { $item.Finish = "" }
                }
            }
            $filteredMaster.Add($item)
        }
    }

    $filteredLabor = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $LaborData) {
        $assmKey = $item.Assembly.Trim()
        if ($allowed.Contains($assmKey)) {
            $fKey = "$assmKey|$($item.SourceFile)".ToLowerInvariant()
            if ($qtyFactor.ContainsKey($fKey)) { $item.Qty = [int][Math]::Round($item.Qty * $qtyFactor[$fKey]) }
            if ($hasTlMap) {
                if ($tlMap.ContainsKey($assmKey)) {
                    $item.TL = $tlMap[$assmKey]
                } else {
                    $item.TL = ""
                }
            }
            $filteredLabor.Add($item)
        }
    }

    $filteredDrawings = @{}
    foreach ($key in $AssemblyDrawings.Keys) {
        if ($allowed.Contains($key.Trim())) { $filteredDrawings[$key] = $AssemblyDrawings[$key] }
    }

    Write-Log "Assembly whitelist active: $($filteredMaster.Count) detail records across $($filteredDrawings.Count) assemblies kept from $($available.Count) available." "Green"
    if ($missing.Count -gt 0) {
        Write-Log "Assembly whitelist ignored $($missing.Count) name(s) not found in the parsed data: $($missing -join ', ')" "Yellow"
    }

    if ($filteredMaster.Count -eq 0) {
        Write-Log "No records matched the assembly whitelist; the BOM will be empty." "Yellow"
    }
    elseif ($missing.Count -gt 0) {
        Write-Log "Assembly whitelist matched some names but ignored $($missing.Count) unknown entry(s); review the filter before importing." "Yellow"
    }

    return [PSCustomObject]@{
        Applied = $true
        MasterData = $filteredMaster
        LaborData = $filteredLabor
        AssemblyDrawings = $filteredDrawings
        Included = @($names)
        Missing = @($missing)
        TlMap = $tlMap
        RevMap = $revMap
    }
}

function Test-KissAssemblyIntegrity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "KISS validation failed: file does not exist: $Path"
    }

    $errors = [System.Collections.Generic.List[string]]::new()
    $lines = @(Get-Content -LiteralPath $Path)
    $currentAssembly = $null
    $currentAssemblyQty = 0
    $currentDetailCount = 0

    for ($idx = 0; $idx -lt $lines.Count; $idx++) {
        $line = $lines[$idx]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $parts = @($line.Split(','))
        $kind = $parts[0].Trim().ToUpper()

        switch ($kind) {
            'KISS' {
                if ($parts.Length -lt 3) { $errors.Add("Line $($idx + 1): KISS header is incomplete.") }
            }
            'H' {
                if ($parts.Length -lt 3) { $errors.Add("Line $($idx + 1): H header is incomplete.") }
            }
            'D' {
                if ($parts.Length -lt 12) { $errors.Add("Line $($idx + 1): D record is missing required fields."); continue }

                $assembly = $parts[1].Trim()
                if (-not $assembly) { $errors.Add("Line $($idx + 1): D record has a blank assembly ID."); continue }

                if ($null -ne $currentAssembly -and $assembly -ne $currentAssembly -and $currentDetailCount -gt 0) {
                    $errors.Add("Line $($idx + 1): assembly changed from '$currentAssembly' to '$assembly' before the summary line was written.")
                }

                $qty = 0
                if (-not [int]::TryParse($parts[5].Trim(), [ref]$qty)) {
                    $errors.Add("Line $($idx + 1): D quantity is not numeric for assembly '$assembly'.")
                }

                $mainMark = if ($parts.Length -ge 4) { $parts[3].Trim() } else { '' }
                $rawPartMark = if ($parts.Length -ge 5) { $parts[4].Trim() } else { '' }
                if ($rawPartMark -and $rawPartMark -eq $assembly) {
                    $currentAssemblyQty += $qty
                }

                $currentAssembly = $assembly
                $currentDetailCount++
            }
            'S' {
                if ($parts.Length -lt 3) { $errors.Add("Line $($idx + 1): S summary is incomplete."); continue }
                if (-not $currentAssembly) {
                    $errors.Add("Line $($idx + 1): S summary appears without a preceding D block.")
                    continue
                }

                $summaryQty = 0
                if (-not [int]::TryParse($parts[2].Trim(), [ref]$summaryQty)) {
                    $errors.Add("Line $($idx + 1): summary quantity is not numeric for assembly '$currentAssembly'.")
                    continue
                }

                if ($currentAssemblyQty -gt 0 -and $summaryQty -ne $currentAssemblyQty) {
                    $errors.Add("Line $($idx + 1): assembly '$currentAssembly' S quantity is $summaryQty but the assembly main quantity is $currentAssemblyQty.")
                }

                $currentAssembly = $null
                $currentAssemblyQty = 0
                $currentDetailCount = 0
            }
            'A' {
                if ($parts.Length -lt 2) { $errors.Add("Line $($idx + 1): A record is missing a drawing value.") }
            }
            '*' {
                if ($currentDetailCount -gt 0 -and $currentAssembly) {
                    $errors.Add("Line $($idx + 1): assembly '$currentAssembly' ended with D rows but no matching S summary line.")
                }
                $currentAssembly = $null
                $currentAssemblyQty = 0
                $currentDetailCount = 0
            }
            default {
                if ($kind -ne 'END') {
                    $errors.Add("Line $($idx + 1): unsupported record type '$kind'.")
                }
            }
        }
    }

    if ($currentDetailCount -gt 0 -and $currentAssembly) {
        $errors.Add("File ended while assembly '$currentAssembly' was still open without a closing S line.")
    }

    if ($errors.Count -gt 0) {
        throw "KISS validation failed for ${Path}:`n - $($errors -join "`n - ")"
    }

    return $true
}

# --- PER-FILE SHOP SCRIPTS & MASTER LAUNCHER ---
function Save-BatFile ($Name, $Lines) {
    $path = Join-Path (Get-Location).Path $Name
    [System.IO.File]::WriteAllLines($path, $Lines, [System.Text.Encoding]::ASCII)
    Write-Log "Successfully generated: $Name" "Green"
}

function New-TypeMacroLines ($Title, $Noun, $ValueLabel, $Rows, $ProjectName, $JobNumber, $TL) {
    $tlNote = if ($TL) { " | Transmittal $TL" } else { "" }
    $safeProject = Escape-BatchText ($ProjectName -replace '[&|<>^%]', '')
    $safeProjectLine = Escape-BatchText $ProjectName
    $safeTlNote = Escape-BatchText $tlNote
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("@echo off")
    $lines.Add(":: ========================================================")
    $lines.Add(":: Auto-Type $Title Macro Utility (A-Z Ordered)")
    $lines.Add(":: Project: $safeProjectLine (Job $JobNumber)$safeTlNote")
    $lines.Add(":: Instructions:")
    $lines.Add("::   1. Open your target software window.")
    $lines.Add("::   2. Click and select the FIRST $Noun cell at the top.")
    $lines.Add("::   3. Press any key here to start the 5-second countdown.")
    $lines.Add(":: ========================================================")
    $lines.Add("setlocal enabledelayedexpansion")
    $lines.Add("")
    $lines.Add("echo ========================================================")
    $lines.Add("echo   Auto-Type $Title Macro Initialized (A-Z Order)")
    $lines.Add("echo   Project: $safeProject (Job $JobNumber)$safeTlNote")
    $lines.Add("echo   Total items to type: $(@($Rows).Count)")
    $lines.Add("echo ========================================================")
    $lines.Add("echo.")
    $lines.Add("echo WARNING: Position your cursor on the first $Noun cell NOW.")
    $lines.Add("pause")
    $lines.Add("")
    $lines.Add("echo Starting countdown...")
    $lines.Add("for /l %%i in (5,-1,1) do (")
    $lines.Add("    echo   %%i seconds remaining...")
    $lines.Add("    timeout /t 1 >nul")
    $lines.Add(")")
    $lines.Add("echo GO! Typing $($Title.ToLower())...")
    $lines.Add("")
    $lines.Add("powershell -NoProfile -Command `"`$wshell = New-Object -ComObject WScript.Shell; Start-Sleep -Milliseconds 300`"")

    foreach ($row in $Rows) {
        $sendValue = ConvertTo-SendKeysText $row.Value
        $lines.Add("echo [TYPING] Part $($row.Mark) -^> $ValueLabel $($row.Value)")
        $lines.Add("powershell -NoProfile -Command `"`$wshell = New-Object -ComObject WScript.Shell; `$wshell.SendKeys('$sendValue'); Start-Sleep -Milliseconds 250; `$wshell.SendKeys('{DOWN}'); Start-Sleep -Milliseconds 200`"")
    }

    $lines.Add("")
    $lines.Add("echo.")
    $lines.Add("echo ========================================================")
    $lines.Add("echo   Auto-Type Complete!")
    $lines.Add("echo ========================================================")
    $lines.Add("echo.")
    $lines.Add("pause")
    return $lines
}

function Write-ShopScripts ($Items, $ProjectName, $JobNumber, $TL, $Label) {
    $plateItems = @($Items | Where-Object { $_.Category -eq "Plate" } | Group-Object PartMark)
    if ($plateItems.Count -eq 0) { return $null }

    $tlNote = if ($TL) { " - Transmittal $TL" } else { "" }
    $safeProject = Escape-BatchText ($ProjectName -replace '[&|<>^%]', '')
    $safeProjectHeader = Escape-BatchText $ProjectName
    $safeTlNote = Escape-BatchText $tlNote
    $sortName  = "Sort-PlatesByThickness - $Label.bat"
    $qtyName   = "Type-Quantities - $Label.bat"
    $gradeName = "Type-Grades - $Label.bat"

    $batLines = [System.Collections.Generic.List[string]]::new()
    $batLines.Add("@echo off")
    $batLines.Add(":: ========================================================")
    $batLines.Add(":: DXF Plate & Sheet Sorting Utility")
    $batLines.Add(":: Project: $safeProjectHeader (Job $JobNumber)$safeTlNote")
    $batLines.Add(":: Run from anywhere - auto-detects DXF folder location.")
    $batLines.Add(":: ========================================================")
    $batLines.Add("setlocal enabledelayedexpansion")
    $batLines.Add("")
    $batLines.Add(":: --- Navigate to the script's own directory, then auto-detect DXF folder ---")
    $batLines.Add("cd /d ""%~dp0""")
    $batLines.Add("set ""WORK_DIR=%~dp0""")
    $batLines.Add("if exist ""%~dp0dxf\"" (")
    $batLines.Add("    set ""WORK_DIR=%~dp0dxf\""")
    $batLines.Add("    echo [INFO] Found 'dxf' subfolder - sorting files inside it.")
    $batLines.Add(")")
    $batLines.Add("cd /d ""!WORK_DIR!""")
    $batLines.Add("")
    $batLines.Add("echo ========================================================")
    $batLines.Add("echo   DXF Plate Sorting Utility Initialized")
    $batLines.Add("echo   Project: $safeProject (Job $JobNumber)$safeTlNote")
    $batLines.Add("echo   Working in: !WORK_DIR!")
    $batLines.Add("echo ========================================================")
    $batLines.Add("echo.")
    $batLines.Add("set ""MOVED_COUNT=0""")
    $batLines.Add("set ""MISSING_COUNT=0""")
    $batLines.Add("")

    foreach ($g in $plateItems) {
        $first = $g.Group[0]
        $partMark = $first.PartMark
        if (-not $partMark) { continue }

        $dec = Get-DecimalThickness $first.Size
        $thickFolder = Format-ThicknessFolder $dec

        $batLines.Add("if not exist ""$thickFolder\"" mkdir ""$thickFolder"" >nul 2>&1")
        $batLines.Add("if exist ""$partMark.dxf"" (")
        $batLines.Add("    move /y ""$partMark.dxf"" ""$thickFolder\"" >nul 2>&1")
        $batLines.Add("    echo [SUCCESS] $partMark.dxf -^> $thickFolder")
        $batLines.Add("    set /a MOVED_COUNT+=1")
        $batLines.Add(") else (")
        $batLines.Add("    echo [NOT FOUND] $partMark.dxf")
        $batLines.Add("    set /a MISSING_COUNT+=1")
        $batLines.Add(")")
        $batLines.Add("")
    }

    $batLines.Add("echo.")
    $batLines.Add("echo ========================================================")
    $batLines.Add("echo   Sorting Complete!")
    $batLines.Add("echo   Successfully Moved: !MOVED_COUNT! file(s)")
    $batLines.Add("echo   Missing / Not Found: !MISSING_COUNT! file(s)")
    $batLines.Add("echo ========================================================")
    $batLines.Add("echo.")
    $batLines.Add("pause")
    Save-BatFile $sortName $batLines

    $rolledPlates = @($plateItems | ForEach-Object {
        [PSCustomObject]@{
            Mark  = $_.Group[0].PartMark
            Qty   = ($_.Group | Measure-Object -Property Qty -Sum).Sum
            Grade = $_.Group[0].Grade
            Thick = Get-DecimalThickness $_.Group[0].Size
        }
    } | Sort-Object Mark)

    Write-Host ""
    Write-Host ("Plate / thickness / quantity preview for {0}:" -f $Label)
    foreach ($row in $rolledPlates) {
        $fmtThick = Format-DecimalThickness $row.Thick
        Write-Host ("  {0} | thickness={1} | qty={2} | grade={3}" -f $row.Mark, $fmtThick, $row.Qty, $row.Grade)
    }

    $qtyRows   = @($rolledPlates | ForEach-Object { [PSCustomObject]@{ Mark = $_.Mark; Value = $_.Qty } })
    $gradeRows = @($rolledPlates | ForEach-Object { [PSCustomObject]@{ Mark = $_.Mark; Value = $_.Grade } })
    Save-BatFile $qtyName (New-TypeMacroLines "Quantities" "quantity" "Qty" $qtyRows $ProjectName $JobNumber $TL)
    Save-BatFile $gradeName (New-TypeMacroLines "Grades" "grade" "Grade" $gradeRows $ProjectName $JobNumber $TL)

    return [PSCustomObject]@{ Label = $Label; TL = $TL; Sort = $sortName; Quantities = $qtyName; Grades = $gradeName }
}

function Write-MasterScript ($ScriptSets, $ProjectName, $JobNumber, $TLs) {
    $tlPart = if (@($TLs).Count -gt 0) { " - " + (@($TLs) -join "_") } else { "" }
    if ($tlPart.Length -gt 60) { $tlPart = " - $(@($TLs)[0])-$(@($TLs)[-1])" }
    $masterName = "Master-ShopScripts$tlPart.bat" -replace '[\\/:*?"<>|]', '_'
    $safeProject = Escape-BatchText ($ProjectName -replace '[&|<>^%]', '')

    $m = [System.Collections.Generic.List[string]]::new()
    $m.Add("@echo off")
    $safeProjectHeader = Escape-BatchText $ProjectName
    $m.Add(":: ========================================================")
    $m.Add(":: Master Shop Script Launcher")
    $m.Add(":: Project: $safeProjectHeader (Job $JobNumber)")
    $m.Add(":: ========================================================")
    $m.Add("setlocal enabledelayedexpansion")
    $m.Add("")
    $m.Add(":menu")
    $m.Add("cls")
    $m.Add("echo ========================================================")
    $m.Add("echo   Master Shop Script Launcher")
    $m.Add("echo   Project: $safeProject (Job $JobNumber)")
    $m.Add("echo ========================================================")
    $m.Add("echo.")
    $m.Add("echo   [S] Run ALL DXF sort scripts")
    $m.Add("echo   [Q] Run ALL quantity typing scripts, one after another")
    $m.Add("echo   [G] Run ALL grade typing scripts, one after another")
    $m.Add("echo.")

    $kinds = @(
        @{ Key = "Sort";       Text = "Sort DXFs by thickness" },
        @{ Key = "Quantities"; Text = "Type quantities" },
        @{ Key = "Grades";     Text = "Type grades" }
    )
    $entries = [System.Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($s in $ScriptSets) {
        $safeLabel = $s.Label -replace '[&|<>^%()]', ''
        foreach ($k in $kinds) {
            $n++
            $entries.Add([PSCustomObject]@{ Number = $n; File = $s.($k.Key) })
            $m.Add("echo   [$n] $safeLabel - $($k.Text)")
        }
    }

    $m.Add("echo.")
    $m.Add("echo   [0] Exit")
    $m.Add("echo.")
    $m.Add('set "SEL="')
    $m.Add('set /p "SEL=Select an option: "')
    $m.Add('if /i "!SEL!"=="0" goto :end')
    $m.Add('if /i "!SEL!"=="S" goto :runsort')
    $m.Add('if /i "!SEL!"=="Q" goto :runqty')
    $m.Add('if /i "!SEL!"=="G" goto :rungrade')
    foreach ($e in $entries) {
        $m.Add(('if "!SEL!"=="{0}" (' -f $e.Number))
        $m.Add(('    call "%~dp0{0}"' -f $e.File))
        $m.Add('    goto :menu')
        $m.Add(')')
    }
    $m.Add('goto :menu')
    $m.Add('')

    foreach ($run in @(@{ Label = "runsort"; Key = "Sort" }, @{ Label = "runqty"; Key = "Quantities" }, @{ Label = "rungrade"; Key = "Grades" })) {
        $m.Add(":$($run.Label)")
        foreach ($s in $ScriptSets) { $m.Add(('call "%~dp0{0}"' -f $s.($run.Key))) }
        $m.Add('goto :menu')
        $m.Add('')
    }

    $m.Add(':end')
    $m.Add('endlocal')
    Save-BatFile $masterName $m
    return $masterName
}

function ConvertFrom-SoundwallCsv ($CsvPath, $AnchorsPerPost) {
    $rows = @(Import-Csv -Path $CsvPath)
    $items = [System.Collections.Generic.List[PSCustomObject]]::new()
    $labor = [System.Collections.Generic.List[PSCustomObject]]::new()
    $drawings = @{}
    $sourceFile = [System.IO.Path]::GetFileName($CsvPath)

    $drawingColumn = $null
    if ($rows.Count -gt 0) {
        $colNames = $rows[0].PSObject.Properties.Name
        $drawingColumn = $colNames | Where-Object { $_ -match '^(Drawing|Dwg( No)?)$' } | Select-Object -First 1
    }

    foreach ($row in $rows) {
        $postId = $row.'Post ID'
        if (-not $postId) { continue }
        $assm = $postId.Trim()

        if ($drawingColumn -and $row.$drawingColumn) {
            $drawings[$assm] = $row.$drawingColumn.Trim()
        }

        # POST (Profile)
        $postSize = $row.'Post Size'
        $postType = $row.'Post Type'
        $postLength = $row.'Post Length'
        if ($postSize -and $postType) {
            $items.Add([PSCustomObject]@{
                Assembly    = $assm
                MainMark    = $assm
                PartMark    = $postType.Trim()
                RawPartMark = $postType.Trim()
                Qty         = 1
                Shape       = Get-ShapePrefix $postSize
                Size        = $postSize
                Grade       = ""
                Length      = $postLength
                Finish      = ""
                Description = "Soundwall Post"
                Weight      = Get-WShapeUnitWeight $postSize $postLength
                Category    = "Profile"
                SourceFile  = $sourceFile
            })
        }

        # BASE PLATE
        $bpType = $row.'Base Plate Type'
        if ($bpType) {
            $items.Add([PSCustomObject]@{
                Assembly    = $assm
                MainMark    = $assm
                PartMark    = $bpType.Trim()
                RawPartMark = $bpType.Trim()
                Qty         = 1
                Shape       = "BPL"
                Size        = "-"
                Grade       = ""
                Length      = "-"
                Finish      = ""
                Description = "Base Plate"
                Weight      = 0.0
                Category    = "Plate"
                SourceFile  = $sourceFile
            })
        }

        # ANCHOR RODS
        $arType = $row.'Anchor Rod Type'
        if ($arType) {
            $items.Add([PSCustomObject]@{
                Assembly    = $assm
                MainMark    = $assm
                PartMark    = $arType.Trim()
                RawPartMark = $arType.Trim()
                Qty         = $AnchorsPerPost
                Shape       = "AB"
                Size        = "-"
                Grade       = ""
                Length      = "-"
                Finish      = ""
                Description = "Anchor Rod"
                Weight      = 0.0
                Category    = "Hardware"
                SourceFile  = $sourceFile
            })
        }

        # FLAT BARS
        foreach ($side in @('Left', 'Right')) {
            $sizeProp = "Flat Bar $side Size"
            $lenProp  = "Flat Bar $side Length"
            $sizeField = $row.$sizeProp
            $lenField  = $row.$lenProp
            if ($sizeField) {
                $parsed = Get-QtyAndSize $sizeField
                $mark = Get-OrCreatePartMark "FB" "$($parsed.Size)|$lenField"
                $items.Add([PSCustomObject]@{
                    Assembly    = $assm
                    MainMark    = $assm
                    PartMark    = $mark
                    RawPartMark = $mark
                    Qty         = $parsed.Qty
                    Shape       = "PL"
                    Size        = $parsed.Size
                    Grade       = ""
                    Length      = $lenField
                    Finish      = ""
                    Description = "Flat Bar ($side)"
                    Weight      = Get-FlatPlateUnitWeight $parsed.Size $lenField
                    Category    = "Plate"
                    SourceFile  = $sourceFile
                })
            }
        }

        # SEAT PLATES
        foreach ($side in @('Left', 'Right')) {
            $sizeProp   = "Seat Plate $side Size"
            $heightProp = "Seat Plate $side Height"
            $sizeField   = $row.$sizeProp
            $heightField = $row.$heightProp
            if ($sizeField) {
                $mark = Get-OrCreatePartMark "SPL" "$sizeField|$heightField"
                $items.Add([PSCustomObject]@{
                    Assembly    = $assm
                    MainMark    = $assm
                    PartMark    = $mark
                    RawPartMark = $mark
                    Qty         = 1
                    Shape       = "PL"
                    Size        = $sizeField
                    Grade       = ""
                    Length      = $heightField
                    Finish      = ""
                    Description = "Seat Plate ($side) @ Elev $heightField"
                    Weight      = Get-FlatPlateUnitWeight $sizeField $heightField
                    Category    = "Plate"
                    SourceFile  = $sourceFile
                })
            }
        }
    }

    return [PSCustomObject]@{
        MasterData       = $items
        LaborData        = $labor
        AssemblyDrawings = $drawings
    }
}

function Write-GeneratedKssFile ($MasterData, $LaborData, $AssemblyDrawings, $ProjectName, $JobNumber, $SourceCsv, $TL = "") {
    $laborMap = @{}
    foreach ($l in $LaborData) {
        $key = "$($l.Assembly)|$($l.PartMark)"
        if (-not $laborMap.ContainsKey($key)) { $laborMap[$key] = [System.Collections.Generic.List[object]]::new() }
        $laborMap[$key].Add($l)
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("* Generated from $([System.IO.Path]::GetFileName($SourceCsv)) on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("H,$JobNumber,$ProjectName")

    $grouped = $MasterData | Group-Object Assembly
    foreach ($g in $grouped) {
        $mainAssemblyQty = ($g.Group | Where-Object { $_.PartMark -eq $_.Assembly -or $_.RawPartMark -eq $_.Assembly } | Measure-Object -Property Qty -Sum).Sum
        if (-not $mainAssemblyQty) { $mainAssemblyQty = ($g.Group | Measure-Object -Property Qty -Sum).Sum }
        if (-not $mainAssemblyQty) { $mainAssemblyQty = 1 }
        $prefixDigits = if ($g.Name -match '^(\d+)') { $Matches[1] } else { $g.Name }
        $prefix = if ($prefixDigits.Length -gt 4) { $prefixDigits.Substring(0, 4) } else { $prefixDigits }

        foreach ($item in $g.Group) {
            $safeDesc = ($item.Description -replace ',', ';')
            $fields = @(
                "D", $item.Assembly, "0", $item.MainMark, $item.RawPartMark, $item.Qty,
                $item.Shape, $item.Size, $item.Grade, $item.Length, $item.Finish,
                $safeDesc, "", $item.Weight
            )
            $lines.Add(($fields -join ','))

            $laborKey = "$($item.Assembly)|$($item.PartMark)"
            if ($laborMap.ContainsKey($laborKey)) {
                foreach ($l in $laborMap[$laborKey]) {
                    $lFields = @("L", $l.LaborType, $l.Qty, $l.Size1, $l.Size2, $l.Notes)
                    $lines.Add(($lFields -join ','))
                }
            }
        }

        $lines.Add("S,$prefix,$mainAssemblyQty")
        if ($AssemblyDrawings.ContainsKey($g.Name)) {
            $lines.Add("A,$($AssemblyDrawings[$g.Name])")
        }
        $lines.Add("*")
    }

    $safeProjectName = $ProjectName -replace '[\\/:*?"<>|]', '_'
    $tlPart = if ($TL) { " - $TL" } else { "" }
    $outPath = Join-Path (Get-Location) "$safeProjectName - Job $JobNumber$tlPart (Generated).kss"
    [System.IO.File]::WriteAllLines($outPath, $lines, [System.Text.Encoding]::ASCII)

    try {
        Test-KissAssemblyIntegrity -Path $outPath
        Write-Log "Validated generated KISS export: $outPath" "Green"
    }
    catch {
        Write-Log "Generated KISS export failed validation and was removed: $($_.Exception.Message)" "Red"
        Remove-Item -LiteralPath $outPath -ErrorAction SilentlyContinue
        throw
    }

    return $outPath
}

function Write-FilteredAssemblyKssFile ($MasterData, $AssemblyDrawings, $ProjectName, $JobNumber, $SourceFiles, $TLs, $SelectedAssemblies = $null) {
    $assemblies = if ($null -ne $SelectedAssemblies) { @($SelectedAssemblies | Select-Object -Unique) } else { @($MasterData | ForEach-Object { $_.Assembly } | Select-Object -Unique) }
    if ($assemblies.Count -eq 0) { return $null }

    $grouped = @($MasterData | Where-Object { $_.Assembly -in $assemblies } | Group-Object Assembly | Sort-Object Name)
    if ($grouped.Count -eq 0) { return $null }

    $headerLine = "H,$JobNumber,$ProjectName"
    foreach ($file in @($SourceFiles)) {
        $content = Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue
        foreach ($line in $content) {
            if ($line -and $line.Trim().StartsWith("H,")) {
                $headerLine = $line.Trim()
                break
            }
        }
        if ($headerLine -and $headerLine -ne "H,$JobNumber,$ProjectName") { break }
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("KISS,1.1,Tekla Structures")
    $lines.Add($headerLine)

    foreach ($g in $grouped) {
        $assmName = $g.Name
        $sourceGroups = @($g.Group | Group-Object SourceFile)
        $canonicalGroup = @($g.Group)
        if ($sourceGroups.Count -gt 1) {
            $sourceOrder = @{}
            for ($i = 0; $i -lt @($SourceFiles).Count; $i++) {
                $sourceOrder[(@($SourceFiles)[$i]).Name] = $i
            }
            $canonicalSource = $sourceGroups | Sort-Object {
                if ($sourceOrder.ContainsKey($_.Name)) { $sourceOrder[$_.Name] } else { 9999 }
            } | Select-Object -First 1
            $canonicalGroup = @($canonicalSource.Group)
        }

        $uniqueItems = @(
            $canonicalGroup |
            Group-Object -Property {
                "$($_.Assembly)|$($_.PartMark)|$($_.Shape)|$($_.Size)|$($_.Grade)|$($_.Length)|$($_.Finish)|$($_.Description)|$($_.MainMark)"
            } |
            ForEach-Object {
                $first = $_.Group[0]
                $totalQty = ($_.Group | Measure-Object -Property Qty -Sum).Sum
                [PSCustomObject]@{
                    Assembly    = $first.Assembly
                    MainMark    = $first.MainMark
                    PartMark    = $first.PartMark
                    RawPartMark = $first.RawPartMark
                    Qty         = $totalQty
                    Shape       = $first.Shape
                    Size        = $first.Size
                    Grade       = $first.Grade
                    Length      = $first.Length
                    Finish      = $first.Finish
                    Description = $first.Description
                    Weight      = $first.Weight
                    Rev         = $first.Rev
                    Tail        = $first.Tail
                    TL          = $first.TL
                }
            } |
            Sort-Object PartMark
        )

        foreach ($item in $uniqueItems) {
            $safeDesc = ($item.Description -replace ',', ';')
            if (-not $TransmittalInFinish) {
                # Finish (field 10) is untouched; the TL goes in the cost code slot (field 13); revision and fields 12+ pass through.
                $rev  = if ($item.Rev) { $item.Rev } else { "0" }
                $tail = if ($null -ne $item.Tail -and @($item.Tail).Count -ge 2) { @($item.Tail) } else { @("", "", $item.Weight) }
                if ($item.TL) { $tail[1] = $item.TL }
                $fields = @(
                    "D", $item.Assembly, $rev, $item.MainMark, $item.RawPartMark, $item.Qty,
                    $item.Shape, $item.Size, $item.Grade, $item.Length, $item.Finish,
                    $safeDesc
                ) + $tail
            }
            else {
                $fields = @(
                    "D", $item.Assembly, "0", $item.MainMark, $item.RawPartMark, $item.Qty,
                    $item.Shape, $item.Size, $item.Grade, $item.Length, $item.Finish,
                    $safeDesc, "", $item.Weight
                )
            }
            $lines.Add(($fields -join ','))
        }

        $prefixDigits = if ($assmName -match '^(\d+)') { $Matches[1] } else { $assmName }
        $prefix = if ($prefixDigits.Length -gt 4) { $prefixDigits.Substring(0, 4) } else { $prefixDigits }
        $assemblyQty = ($uniqueItems | Where-Object { $_.PartMark -eq $_.Assembly -or $_.RawPartMark -eq $_.Assembly } | Measure-Object -Property Qty -Sum).Sum
        if (-not $assemblyQty) { $assemblyQty = ($uniqueItems | Measure-Object -Property Qty -Sum).Sum }
        if (-not $assemblyQty) { $assemblyQty = 1 }
        $lines.Add("S,$prefix,$assemblyQty")

        if ($AssemblyDrawings.ContainsKey($assmName) -and $AssemblyDrawings[$assmName]) {
            $lines.Add("A,$($AssemblyDrawings[$assmName])")
        }
        $lines.Add("*")
    }

    $safeProjectName = $ProjectName -replace '[\\/:*?"<>|]', '_'
    $tlNamePart = ""
    if (@($TLs).Count -gt 0) {
        $tlNamePart = " - " + (@($TLs) -join "_")
        if (($safeProjectName.Length + $jobNumber.Length + $tlNamePart.Length + 35) -gt 220) {
            $tlNamePart = " - $(@($TLs)[0])-$(@($TLs)[-1])"
        }
    }
    $outPath = Join-Path (Get-Location) "$safeProjectName - Job $JobNumber$tlNamePart - Filtered Assemblies.kss"
    [System.IO.File]::WriteAllLines($outPath, $lines, [System.Text.Encoding]::ASCII)

    try {
        Test-KissAssemblyIntegrity -Path $outPath
        Write-Log "Validated filtered KISS export: $outPath" "Green"
    }
    catch {
        Write-Log "Filtered KISS export failed validation and was removed: $($_.Exception.Message)" "Red"
        Remove-Item -LiteralPath $outPath -ErrorAction SilentlyContinue
        throw
    }

    return $outPath
}


# 1. SELECT DATA SOURCE
Write-Log "Initializing KSS Nesting, Hardware & Shop Management Engine..." "Yellow"
Write-Host ""
Write-Host "========================================================"
Write-Host "  ASSEMBLY LIST FORMAT (when using whitelist)"
Write-Host "========================================================"
Write-Host "  Simple list (comma/space/newline separated):"
Write-Host "    51328511"
Write-Host "    51318512"
Write-Host ""
Write-Host "  WITH TL PREFIXES (for TL-specific filtering):"
Write-Host "    TL1296  51328511 51318512 51316500"
Write-Host "    TL1313  51338507 51338511"
Write-Host ""
Write-Host "  WITH REV MARKER (R prefix = override older TL versions):"
Write-Host "    R51328511  <- REV: will replace this part from earlier TLs"
Write-Host "    R51318512  <- REV: will replace this part from earlier TLs"
Write-Host ""
Write-Host "  WITH MANUAL TL ORDERING (number prefix for sort order):"
Write-Host "    1_TL1301   <- Process TL1301 FIRST"
Write-Host "    2_TL1282   <- Process TL1282 SECOND"
Write-Host "    3_TL1296   <- Process TL1296 THIRD"
Write-Host ""
Write-Host "========================================================"
Write-Host "  Select Data Source"
Write-Host "========================================================"
Write-Host " [1] KSS Files (.kss nesting/detail export)"
Write-Host " [2] Soundwall CSV (post schedule import)"
Write-Host ""
$modeChoice = Read-Host "Enter selection (1 or 2)"

# 2. PARSE DATA & METADATA
$masterData = [System.Collections.Generic.List[PSCustomObject]]::new()
$laborData  = [System.Collections.Generic.List[PSCustomObject]]::new()
$assemblyDrawings = @{}
$projectName = "STEEL FABRICATION PROJECT"
$jobNumber = "000"
$allTLs = @()

if ($modeChoice -eq "2") {
    # --- SOUNDWALL CSV MODE ---
    $csvPath = Select-CsvFile
    if (-not $csvPath) {
        Write-Log "No CSV file selected. Exiting." "Red"
        return
    }
    Write-Log "Selected CSV: $csvPath" "Green"

    $csvBaseName = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
    $projectInput = Read-Host "Enter Project Name (default: $($csvBaseName.ToUpper()))"
    $projectName = if ($projectInput) { $projectInput } else { $csvBaseName.ToUpper() }

    $jobInput = Read-Host "Enter Job Number (default: 000)"
    if ($jobInput) { $jobNumber = $jobInput }

    $csvTL = ConvertTo-TransmittalTag (Read-Host "Enter TL # (e.g. TL012, blank for none)")
    $allTLs = @($csvTL | Where-Object { $_ })

    $anchorsPerPost = 4
    $arQtyInput = Read-Host "Anchor rods per post (default: 4)"
    if ($arQtyInput) { [int]::TryParse($arQtyInput, [ref]$anchorsPerPost) | Out-Null }

    $csvResult = ConvertFrom-SoundwallCsv -CsvPath $csvPath -AnchorsPerPost $anchorsPerPost
    foreach ($item in $csvResult.MasterData) {
        $item | Add-Member -NotePropertyName TL -NotePropertyValue $csvTL -Force
        $masterData.Add($item)
    }
    foreach ($item in $csvResult.LaborData) {
        $item | Add-Member -NotePropertyName TL -NotePropertyValue $csvTL -Force
        $laborData.Add($item)
    }
    foreach ($key in $csvResult.AssemblyDrawings.Keys) { $assemblyDrawings[$key] = $csvResult.AssemblyDrawings[$key] }

    if ($masterData.Count -eq 0) {
        Write-Log "No valid post data found in the CSV." "Red"
        return
    }

    $generatedKssPath = Write-GeneratedKssFile -MasterData $masterData -LaborData $laborData -AssemblyDrawings $assemblyDrawings -ProjectName $projectName -JobNumber $jobNumber -SourceCsv $csvPath -TL $csvTL
    Write-Log "Generated companion KSS file: $generatedKssPath" "Green"

    $continueInput = Read-Host "Continue and generate the full BOM spreadsheet now (tabs, DXF sorting script, etc.)? (Y/n)"
    if ($continueInput -match '^(n|no)$') {
        Write-Log "Skipping spreadsheet generation. The KSS file has been saved." "Yellow"
        return
    }
}
else {
    # --- KSS FILE MODE ---
    $kssFiles = @(Select-KssFiles)
    if ($kssFiles.Count -eq 0) {
        Write-Log "No KSS files selected. Exiting." "Red"
        return
    }
    $allTLs = @($kssFiles | ForEach-Object { $_.TL } | Where-Object { $_ } | Select-Object -Unique)
    Write-Log "Processing $($kssFiles.Count) KSS file(s)..."
    $headerLocked = $false

    foreach ($file in $kssFiles) {
        $currentAssembly = ""
        $currentDetailItem = $null
        $lines = Get-Content -LiteralPath $file.FullName
        foreach ($line in $lines) {
            if (-not $line -or $line.StartsWith("*")) { continue }

            $parts = $line.Split(',')
            $lineType = $parts[0].Trim().ToUpper()

            # Header Line
            if ($lineType -eq "H") {
                if ($parts.Length -ge 3) {
                    $fileJob = $parts[1].Trim()
                    $fileProject = $parts[2].Trim()
                    if (-not $headerLocked) {
                        if ($fileJob) { $jobNumber = $fileJob }
                        if ($fileProject) { $projectName = $fileProject }
                        $headerLocked = $true
                    }
                    elseif (($fileJob -and $fileJob -ne $jobNumber) -or ($fileProject -and $fileProject -ne $projectName)) {
                        Write-Log "Header in $($file.Name) ($fileJob / $fileProject) differs from the first file; keeping $jobNumber / $projectName." "Yellow"
                    }
                }
            }

            # Address / Assembly-Drawing Mapping
            elseif ($lineType -eq "A") {
                if ($parts.Length -ge 2 -and $currentAssembly) {
                    $dwgVal = $parts[1].Trim()
                    if ($dwgVal -and -not $dwgVal.StartsWith("A,")) {
                        $assemblyDrawings[$currentAssembly] = $dwgVal
                    }
                }
            }

            # Detail Line
            elseif ($lineType -eq "D") {
                if ($parts.Length -ge 10) {
                    $qty = 0
                    if (-not [int]::TryParse($parts[5].Trim(), [ref]$qty)) { continue }

                    $weight = 0.00
                    if ($StandardWeight) {
                        # KISS v1.1: field 14 = gross part weight, field 15 = net. Blank/absent = 0 (Tekla exports carry none).
                        $weight = [decimal]0
                        foreach ($wIdx in 14, 15) {
                            if ($parts.Length -gt $wIdx) {
                                $candidate = $parts[$wIdx].Trim()
                                $parsedWeight = [decimal]0
                                if ($candidate -and [decimal]::TryParse($candidate, [ref]$parsedWeight) -and $parsedWeight -gt 0) {
                                    $weight = $parsedWeight
                                    break
                                }
                            }
                        }
                    }
                    else {
                        for ($idx = $parts.Length - 1; $idx -ge 0; $idx--) {
                            $candidate = $parts[$idx].Trim()
                            if ($candidate -and [decimal]::TryParse($candidate, [ref]$weight)) {
                                break
                            }
                        }
                    }

                    $assm = $parts[1].Trim()
                    $currentAssembly = $assm

                    $rawPartMark = $parts[4].Trim()
                    $shape = $parts[6].Trim().ToUpper()
                    $size = $parts[7].Trim()
                    $grade = $parts[8].Trim()
                    $length = $parts[9].Trim()
                    $finish = if ($parts.Length -ge 11) { $parts[10].Trim() } else { "" }
                    $desc = if ($parts.Length -ge 12) { $parts[11].Trim() } else { "" }

                    # Format part mark for loose hardware items with blank mark fields
                    $effectiveMark = $rawPartMark
                    if (-not $effectiveMark) {
                        if ($desc) { $effectiveMark = "$shape-$desc" }
                        else { $effectiveMark = "$shape-$size" }
                    }

                    $category = Get-MaterialCategory -shape $shape -partMark $effectiveMark -description $desc -size $size

                    $currentDetailItem = [PSCustomObject]@{
                        Assembly    = $assm
                        MainMark    = $parts[3].Trim()
                        PartMark    = $effectiveMark
                        RawPartMark = $rawPartMark
                        Qty         = $qty
                        Shape       = $shape
                        Size        = $size
                        Grade       = $grade
                        Length      = $length
                        Finish      = $finish
                        Description = $desc
                        Weight      = $weight
                        Rev         = $parts[2].Trim()
                        Tail        = @($parts | Select-Object -Skip 12 | ForEach-Object { $_.Trim() })
                        Category    = $category
                        SourceFile  = $file.Name
                        TL          = $file.TL
                    }

                    $masterData.Add($currentDetailItem)
                }
            }

            # Labor / Operations Line (Holes, Welds, Cuts, Copes)
            elseif ($lineType -eq "L") {
                if ($currentDetailItem -and $parts.Length -ge 3) {
                    $lType = $parts[1].Trim()
                    $lQty = 0
                    [int]::TryParse($parts[2].Trim(), [ref]$lQty) | Out-Null
                    $lSize1 = if ($parts.Length -ge 4) { $parts[3].Trim() } else { "" }
                    $lSize2 = if ($parts.Length -ge 5) { $parts[4].Trim() } else { "" }
                    $lNotes = if ($parts.Length -ge 6) { $parts[5].Trim() } else { "" }

                    $laborData.Add([PSCustomObject]@{
                        Assembly    = $currentDetailItem.Assembly
                        PartMark    = $currentDetailItem.PartMark
                        LaborType   = $lType
                        Qty         = $lQty
                        Size1       = $lSize1
                        Size2       = $lSize2
                        Notes       = $lNotes
                        SourceFile  = $file.Name
                        TL          = $file.TL
                    })
                }
            }
        }
    }

    if ($masterData.Count -eq 0) {
        Write-Log "No valid part data found in the KSS files." "Red"
        return
    }
}

Write-Log "Extracted Project: $projectName (Job $jobNumber)" "Green"
Write-Log "Total Detail Records Parsed: $($masterData.Count)" "Cyan"
if ($allTLs.Count -gt 0) { Write-Log "Transmittals: $($allTLs -join ', ')" "Cyan" }

# Folder labels: optional Required Assm workbook (asked once, built later inside the BOM workbook)
$requiredAssmPath = $null
$ourJobNumber = ""
$issuedBy = ""
if ($modeChoice -ne "2") {
    $requiredAssmPath = Select-RequiredAssmFile
    if ($requiredAssmPath) {
        $guessJob = ""
        if ([System.IO.Path]::GetFileName($requiredAssmPath) -match '^\s*(\d{4,6})') { $guessJob = $Matches[1] }
        $jobPrompt = if ($guessJob) { "Our job number (default: $guessJob)" } else { "Our job number (shown first on the label)" }
        $ourJobInput = (Read-Host $jobPrompt).Trim()
        $ourJobNumber = if ($ourJobInput) { $ourJobInput } else { $guessJob }
        $issuedBy = (Read-Host "Initials of who is making the folders (Enter to leave blank)").Trim().ToUpper()
    }
}

$whitelistResult = $null
Write-Host ""
Write-Host "========================================================"
Write-Host "  ASSEMBLY FILTERING OPTIONS"
Write-Host "========================================================"
Write-Host "  [Y]es: Load assembly whitelist to include ONLY specific assemblies"
Write-Host "        - Use for REV (revision) runs, partial builds, or focused subsets"
Write-Host "        - Assembly list can include R prefix for REV overrides (e.g., R51328511)"
Write-Host "        - Can include TL prefixes to organize by transmittal (e.g., TL1301 51328511)"
Write-Host ""
Write-Host "  [N]o: Include ALL assemblies from selected KSS files (default)"
Write-Host "       - Use for complete project BOMs"
Write-Host ""
$whitelistChoice = Read-Host "Limit the run to a specific assembly list? [Y]es / [N]o (default: No)"
if ($whitelistChoice -match '^(y|yes)$') {
    $whitelistText = Read-Host "Paste assembly list (comma/newline/space separated, works from Excel), or type @path.txt to load a file. Press Enter to use clipboard text if available."
    if ([string]::IsNullOrWhiteSpace($whitelistText)) {
        try {
            $whitelistText = Get-Clipboard -ErrorAction Stop
        }
        catch {
            $whitelistText = ""
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($whitelistText)) {
        $whitelistResult = Apply-AssemblyWhitelist -MasterData $masterData -LaborData $laborData -AssemblyDrawings $assemblyDrawings -AssemblyListText $whitelistText
        if ($whitelistResult.Applied) {
            if ($whitelistResult.Missing.Count -gt 0 -and $whitelistResult.MasterData.Count -eq 0) {
                throw "Assembly whitelist produced zero matching records. Check the filter values in '$whitelistText' or the source data."
            }
            $masterData = $whitelistResult.MasterData
            $laborData = $whitelistResult.LaborData
            $assemblyDrawings = $whitelistResult.AssemblyDrawings
        }
    }
    else {
        Write-Log "No assembly list entered; keeping all assemblies." "Yellow"
    }
}

Write-Log "Filtered Detail Records: $($masterData.Count)" "Cyan"

# Check if REV assemblies exist and offer override option
$revAssemblies = @{}
$applyRevOverride = $true
if ($null -ne $whitelistResult -and $whitelistResult.Applied) {
    $revAssemblies = $whitelistResult.RevMap
}

if ($revAssemblies.Count -gt 0) {
    Write-Host ""
    Write-Host "========================================================"
    Write-Host "  REV ASSEMBLY OVERRIDE"
    Write-Host "========================================================"
    Write-Host "  Found $($revAssemblies.Count) REV-marked assembly(ies): $($revAssemblies.Keys -join ', ')"
    Write-Host ""
    Write-Host "  [Y]es (default): REV assemblies OVERRIDE earlier TL versions"
    Write-Host "    If 51328511 is in TL1296 AND TL1301 (marked R), keep ONLY TL1301 version"
    Write-Host ""
    Write-Host "  [N]o: Keep all versions, show duplicates in output"
    Write-Host "    51328511 will appear for both TL1296 and TL1301"
    Write-Host ""
    $revChoice = Read-Host "Apply REV overrides? [Y]es (default) / [N]o"
    $applyRevOverride = if ($revChoice -match '^(n|no)$') { $false } else { $true }
}

# Duplicate handling: merge identical marks across files, or keep each file's lines separate
$mergeDuplicates = $true
$sourceNames = @($masterData | ForEach-Object { $_.SourceFile } | Select-Object -Unique)
if ($sourceNames.Count -gt 1) {
    $spansFiles = { param($group) @($group.Group | ForEach-Object { $_.SourceFile } | Select-Object -Unique).Count -gt 1 }
    $sharedAssemblies = @($masterData | Group-Object Assembly | Where-Object { & $spansFiles $_ })
    $sharedMarks = @($masterData | Group-Object PartMark | Where-Object { & $spansFiles $_ })
    Write-Host ""
    Write-Host "========================================================"
    Write-Host "  DUPLICATE HANDLING (parts appearing in multiple TLs)"
    Write-Host "========================================================"
    Write-Host "Assemblies appearing in more than one file: $($sharedAssemblies.Count)"
    Write-Host "Part marks appearing in more than one file: $($sharedMarks.Count)"
    Write-Host ""
    Write-Host "  [M]erge (default): Combine identical parts into ONE row"
    Write-Host "    Example: Part 7p352 qty=90 in TL1296 + qty=90 in TL1313 = 180 total"
    Write-Host "    Use for: Complete project BOMs, purchasing takeoffs"
    Write-Host ""
    Write-Host "  [S]eparate: Keep each file's version SEPARATE by TL/file"
    Write-Host "    Example: 7p352 appears as two rows (TL1296 qty=90, TL1313 qty=90)"
    Write-Host "    Use for: Comparing TL versions, REV tracking, split manufacturing"
    Write-Host "    Note: REV deduplication will remove older TL versions automatically"
    Write-Host ""
    $dupChoice = Read-Host "Duplicates across files: [M]erge (default) or [S]eparate?"
    if ($dupChoice -match '^\s*s') { $mergeDuplicates = $false }

    if ($mergeDuplicates) {
        $conflicts = @($sharedMarks | Where-Object { @($_.Group | ForEach-Object { "$($_.Shape)|$($_.Size)|$($_.Grade)|$($_.Length)" } | Select-Object -Unique).Count -gt 1 })
        if ($conflicts.Count -gt 0) {
            $sample = ($conflicts | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ', '
            Write-Log "WARNING: $($conflicts.Count) part mark(s) have different size/grade/length in different files but will be merged (e.g. $sample). Rerun with [S]eparate to keep them apart." "Yellow"
        }
    }
}

# Apply REV overrides to masterData (works in BOTH Merge and Separate modes)
if ($applyRevOverride -and $revAssemblies.Count -gt 0 -and $sourceNames.Count -gt 1) {
    $tlPattern = { param($name) if ($name -match '(?i)(?:(\d+)_)?TL[\s_-]?(\d+)') {
        [int]$Matches[2]
    } else { -1 } }

    $tlOrder = @{}
    foreach ($file in @($masterData | ForEach-Object { $_.SourceFile } | Select-Object -Unique)) {
        $tlNum = & $tlPattern $file
        $tlOrder[$file] = $tlNum
    }

    $filteredData = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seenAssemblies = @{}

    # Sort by TL number descending (newer first)
    $sortedItems = @($masterData | Sort-Object { $tlOrder[[string]$_.SourceFile] } -Descending)

    foreach ($item in $sortedItems) {
        $assmKey = $item.Assembly
        $isRevAssembly = $revAssemblies.ContainsKey($assmKey)

        if ($seenAssemblies.ContainsKey($assmKey)) {
            # Skip earlier version UNLESS it's marked REV (override case)
            if ($isRevAssembly) {
                $idx = $filteredData.FindIndex({ param($x) $x.Assembly -eq $assmKey })
                if ($idx -ge 0) {
                    $filteredData[$idx] = $item
                    $seenAssemblies[$assmKey] = $item.SourceFile
                }
            }
        } else {
            $seenAssemblies[$assmKey] = $item.SourceFile
            $filteredData.Add($item)
        }
    }

    if ($filteredData.Count -lt $masterData.Count) {
        $removed = $masterData.Count - $filteredData.Count
        Write-Log "REV override: removed $removed duplicate assembly records, keeping REV versions." "Green"
        $masterData = $filteredData
    }
}

$partGroupKeys = @(if ($mergeDuplicates) { 'PartMark' } else { 'SourceFile'; 'PartMark' })
$assemblyGroupKeys = @(if ($mergeDuplicates) { 'Assembly' } else { 'SourceFile'; 'Assembly' })

# REV deduplication: when parts appear in multiple TLs with "Separate" mode selected,
# keep only the highest-numbered (newest) TL version, with REV-marked assemblies taking priority
if (-not $mergeDuplicates -and $sourceNames.Count -gt 1) {
    # Get RevMap from whitelist result (if any)
    $revAssemblies = @{}
    if ($null -ne $whitelistResult -and $whitelistResult.Applied) {
        $revAssemblies = $whitelistResult.RevMap
    }

    # Extract TL number and manual sort order (e.g., "1_TL3" -> order=1, tlNum=3)
    $tlPattern = { param($name) if ($name -match '(?i)(?:(\d+)_)?TL[\s_-]?(\d+)') {
        $order = if ($Matches[1]) { [int]$Matches[1] } else { 999 }
        $tlNum = [int]$Matches[2]
        [PSCustomObject]@{ Order = $order; TlNum = $tlNum }
    } else { [PSCustomObject]@{ Order = 999; TlNum = -1 } } }

    $tlOrder = @{}
    foreach ($file in @($masterData | ForEach-Object { $_.SourceFile } | Select-Object -Unique)) {
        $result = & $tlPattern $file
        $tlOrder[$file] = $result
    }

    $dedupData = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seenParts = @{}

    # Sort by: 1) manual order number, 2) TL number descending (newer first), 3) REV-marked assemblies last (override)
    $sortedItems = @($masterData | Sort-Object -Property @(
        @{ Expression = { $tlOrder[[string]$_.SourceFile].Order }; Ascending = $true },
        @{ Expression = { $tlOrder[[string]$_.SourceFile].TlNum }; Ascending = $false },
        @{ Expression = { if ($revAssemblies.ContainsKey($_.Assembly)) { 0 } else { 1 } }; Ascending = $true }
    ))

    foreach ($item in $sortedItems) {
        $partKey = "$($item.Assembly)|$($item.PartMark)"
        $isRevAssembly = $revAssemblies.ContainsKey($item.Assembly)

        if ($seenParts.ContainsKey($partKey)) {
            # If this item is marked as REV, replace the earlier version
            if ($isRevAssembly) {
                $idx = $dedupData.FindIndex({ param($x) "$($x.Assembly)|$($x.PartMark)" -eq $partKey })
                if ($idx -ge 0) {
                    $dedupData[$idx] = $item
                    $seenParts[$partKey] = $item.SourceFile
                }
            }
        } else {
            $seenParts[$partKey] = $item.SourceFile
            $dedupData.Add($item)
        }
    }

    if ($dedupData.Count -lt $masterData.Count) {
        $removed = $masterData.Count - $dedupData.Count
        Write-Log "REV deduplication: removed $removed duplicate part(s), keeping newest versions (REV-marked assemblies override earlier TL versions)." "Green"
        $masterData = $dedupData
    }
}

$combinedShopScripts = $false
if ($sourceNames.Count -gt 1) {
    Write-Host ""
    Write-Host "========================================================"
    Write-Host "  SHOP SCRIPT GENERATION OPTIONS"
    Write-Host "========================================================"
    Write-Host "  [C]ombined (default): ONE set of scripts for all TLs together"
    Write-Host "    Files: Sort-PlatesByThickness-TL1282_TL1296_TL1301_TL1313.bat"
    Write-Host "           Type-Quantities-TL1282_TL1296_TL1301_TL1313.bat"
    Write-Host "           Type-Grades-TL1282_TL1296_TL1301_TL1313.bat"
    Write-Host "    Use for: Complete shop floor, all plates/quantities/grades in one list"
    Write-Host ""
    Write-Host "  [S]eparate: SEPARATE scripts for EACH KSS file/TL"
    Write-Host "    Files: Sort-PlatesByThickness-TL1282.bat, Sort-PlatesByThickness-TL1296.bat, etc."
    Write-Host "    Use for: Phase manufacturing, different shop stages, TL-specific work orders"
    Write-Host ""
    $shopScriptChoice = Read-Host "Generate shop scripts as [C]ombined (default) or [S]eparate per KSS file?"
    if ($shopScriptChoice -match '^(s|separate)$') { $combinedShopScripts = $false } else { $combinedShopScripts = $true }
}

# Repackage the filtered combined pool into a clean KISS file with only selected assemblies.
$kssFilesInScope = $ExecutionContext.SessionState.PSVariable.GetValue('kssFiles', $null)
if ($null -ne $kssFilesInScope -and $kssFilesInScope.Count -gt 0) {
    $filteredKssOutput = Write-FilteredAssemblyKssFile -MasterData $masterData -AssemblyDrawings $assemblyDrawings -ProjectName $projectName -JobNumber $jobNumber -SourceFiles $kssFilesInScope -TLs $allTLs
    if ($filteredKssOutput) {
        Write-Log "Generated filtered KISS package: $filteredKssOutput" "Green"
    }
}


# 3. EXCEL COM SETUP
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false 
$workbook = $excel.Workbooks.Add()


# --- TAB 1: MASTER BOM ---
Write-Log "Building Tab 1: Master BOM..."
$sheet1 = $workbook.Sheets.Item(1)
$sheet1.Name = "Master BOM"

$mHeaders = @("TL #", "Assembly", "Drawing", "MainMark", "PartMark", "Qty", "Category", "Shape", "Size", "Length", "Grade", "Description", "Unit Wt (lbs)", "Total Wt (lbs)", "Status", "SourceFile")
for ($c = 0; $c -lt $mHeaders.Count; $c++) { $sheet1.Cells.Item(1, $c + 1) = $mHeaders[$c] }

$mRows = $masterData.Count
$mData = New-Object "object[,]" $mRows, $mHeaders.Count

for ($i = 0; $i -lt $mRows; $i++) {
    $item = $masterData[$i]
    $dwg = if ($assemblyDrawings.ContainsKey($item.Assembly)) { $assemblyDrawings[$item.Assembly] } else { "-" }
    
    $mData[$i, 0]  = $item.TL
    $mData[$i, 1]  = $item.Assembly
    $mData[$i, 2]  = $dwg
    $mData[$i, 3]  = $item.MainMark
    $mData[$i, 4]  = $item.PartMark
    $mData[$i, 5]  = $item.Qty
    $mData[$i, 6]  = $item.Category
    $mData[$i, 7]  = $item.Shape
    $mData[$i, 8]  = $item.Size
    $mData[$i, 9]  = $item.Length
    $mData[$i, 10] = $item.Grade
    $mData[$i, 11] = $item.Description
    $mData[$i, 12] = $item.Weight
    $mData[$i, 13] = "=F$(2+$i)*M$(2+$i)"
    $mData[$i, 14] = "PENDING"
    $mData[$i, 15] = $item.SourceFile
}

$sheet1.Range("A2").Resize($mRows, $mHeaders.Count).Value2 = $mData
$sheet1.Range("M2:N$($mRows + 1)").NumberFormat = "#,##0.00"

$mRange = $sheet1.Range($sheet1.Cells.Item(1, 1), $sheet1.Cells.Item($mRows + 1, $mHeaders.Count))
$mTable = $sheet1.ListObjects.Add(1, $mRange, $null, 1)
$mTable.Name = "MasterBOMTable"
$mTable.TableStyle = "TableStyleMedium2"
$mTable.ShowTotals = $true
$mRange.EntireColumn.AutoFit()

$sheet1.Activate()
$excel.ActiveWindow.SplitRow = 1
$excel.ActiveWindow.FreezePanes = $true


# --- TAB 2: PLATES, SHEETMETAL & BENT PLATES ---
$sheet2 = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
$plateGroups = $masterData | Where-Object { $_.Category -eq "Plate" } | Group-Object -Property $partGroupKeys

if ($plateGroups.Count -gt 0) {
    Write-Log "Building Tab 2: Plates & Sheets Nesting..."
    $sheet2.Name = "Plates & Sheets (PL-SHT-BPL)"

    $pHeaders = @("TL #", "Shape", "PartMark", "Thickness - Part", "Total Qty", "Size", "Length", "Grade", "Description", "Unit Wt (lbs)", "Total Wt (lbs)", "SourceFile", "Used In Assemblies")
    for ($c = 0; $c -lt $pHeaders.Count; $c++) { $sheet2.Cells.Item(1, $c + 1) = $pHeaders[$c] }

    $rolledPlates = @($(foreach ($g in $plateGroups) {
        $first = $g.Group[0]
        [PSCustomObject]@{
            Shape       = $first.Shape
            PartMark    = $first.PartMark
            TotalQty    = ($g.Group | Measure-Object -Property Qty -Sum).Sum
            Size        = $first.Size
            Length      = $first.Length
            Grade       = $first.Grade
            Description = $first.Description
            Weight      = $first.Weight
            SourceFile  = ($g.Group | ForEach-Object { $_.SourceFile } | Select-Object -Unique) -join ', '
            Assemblies  = ($g.Group.Assembly | Select-Object -Unique) -join ', '
            TL          = ($g.Group | ForEach-Object { $_.TL } | Where-Object { $_ } | Select-Object -Unique) -join ', '
        }
    }) | Sort-Object Shape, Size, Grade, Length)

    $pCount = $rolledPlates.Count
    $pData = New-Object "object[,]" $pCount, $pHeaders.Count

    for ($i = 0; $i -lt $pCount; $i++) {
        $r = $rolledPlates[$i]
        $rowNum = $i + 2
        $decThick = Get-DecimalThickness $r.Size
        $thickPartStr = "$(Format-DecimalThickness $decThick) - $($r.PartMark)"

        $pData[$i, 0]  = $r.TL
        $pData[$i, 1]  = $r.Shape
        $pData[$i, 2]  = $r.PartMark
        $pData[$i, 3]  = $thickPartStr
        $pData[$i, 4]  = $r.TotalQty
        $pData[$i, 5]  = $r.Size
        $pData[$i, 6]  = $r.Length
        $pData[$i, 7]  = $r.Grade
        $pData[$i, 8]  = $r.Description
        $pData[$i, 9]  = $r.Weight
        $pData[$i, 10] = "=E$rowNum*J$rowNum"
        $pData[$i, 11] = $r.SourceFile
        $pData[$i, 12] = $r.Assemblies
    }

    $sheet2.Range("A2").Resize($pCount, $pHeaders.Count).Value2 = $pData
    $sheet2.Range("J2:K$($pCount + 1)").NumberFormat = "#,##0.00"

    $pRange = $sheet2.Range($sheet2.Cells.Item(1, 1), $sheet2.Cells.Item($pCount + 1, $pHeaders.Count))
    $pTable = $sheet2.ListObjects.Add(1, $pRange, $null, 1)
    $pTable.Name = "PlatesNestingTable"
    $pTable.TableStyle = "TableStyleMedium2"
    $pTable.ShowTotals = $true
    
    $pTable.ListColumns.Item("Total Qty").TotalsCalculation = 1
    $pTable.ListColumns.Item("Total Wt (lbs)").TotalsCalculation = 1
    $pRange.EntireColumn.AutoFit()
    $sheet2.Columns.Item(13).WrapText = $false
    $sheet2.Columns.Item(13).ColumnWidth = 35

    $sheet2.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
} else {
    $sheet2.Name = "Plates & Sheets"
    $sheet2.Cells.Item(1,1) = "No Plate Data Found"
}


# --- TAB 3: STRUCTURAL PROFILES ---
$sheet3 = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
$profileGroups = $masterData | Where-Object { $_.Category -eq "Profile" } | Group-Object -Property $partGroupKeys

if ($profileGroups.Count -gt 0) {
    Write-Log "Building Tab 3: Structural Profiles Nesting..."
    $sheet3.Name = "Profiles (W-HSS-L)"

    $prHeaders = @("TL #", "Shape", "PartMark", "Total Qty", "Size", "Length", "Grade", "Description", "Unit Wt (lbs)", "Total Wt (lbs)", "SourceFile", "Used In Assemblies")
    for ($c = 0; $c -lt $prHeaders.Count; $c++) { $sheet3.Cells.Item(1, $c + 1) = $prHeaders[$c] }

    $rolledProfiles = @($(foreach ($g in $profileGroups) {
        $first = $g.Group[0]
        [PSCustomObject]@{
            Shape       = $first.Shape
            PartMark    = $first.PartMark
            TotalQty    = ($g.Group | Measure-Object -Property Qty -Sum).Sum
            Size        = $first.Size
            Length      = $first.Length
            Grade       = $first.Grade
            Description = $first.Description
            Weight      = $first.Weight
            SourceFile  = ($g.Group | ForEach-Object { $_.SourceFile } | Select-Object -Unique) -join ', '
            Assemblies  = ($g.Group.Assembly | Select-Object -Unique) -join ', '
            TL          = ($g.Group | ForEach-Object { $_.TL } | Where-Object { $_ } | Select-Object -Unique) -join ', '
        }
    }) | Sort-Object Shape, Size, Length)

    $prCount = $rolledProfiles.Count
    $prData = New-Object "object[,]" $prCount, $prHeaders.Count

    for ($i = 0; $i -lt $prCount; $i++) {
        $r = $rolledProfiles[$i]
        $rowNum = $i + 2

        $prData[$i, 0]  = $r.TL
        $prData[$i, 1]  = $r.Shape
        $prData[$i, 2]  = $r.PartMark
        $prData[$i, 3]  = $r.TotalQty
        $prData[$i, 4]  = $r.Size
        $prData[$i, 5]  = $r.Length
        $prData[$i, 6]  = $r.Grade
        $prData[$i, 7]  = $r.Description
        $prData[$i, 8]  = $r.Weight
        $prData[$i, 9]  = "=D$rowNum*I$rowNum"
        $prData[$i, 10] = $r.SourceFile
        $prData[$i, 11] = $r.Assemblies
    }

    $sheet3.Range("A2").Resize($prCount, $prHeaders.Count).Value2 = $prData
    $sheet3.Range("I2:J$($prCount + 1)").NumberFormat = "#,##0.00"

    $prRange = $sheet3.Range($sheet3.Cells.Item(1, 1), $sheet3.Cells.Item($prCount + 1, $prHeaders.Count))
    $prTable = $sheet3.ListObjects.Add(1, $prRange, $null, 1)
    $prTable.Name = "ProfilesNestingTable"
    $prTable.TableStyle = "TableStyleMedium2"
    $prTable.ShowTotals = $true
    
    $prTable.ListColumns.Item("Total Qty").TotalsCalculation = 1
    $prTable.ListColumns.Item("Total Wt (lbs)").TotalsCalculation = 1
    $prRange.EntireColumn.AutoFit()
    $sheet3.Columns.Item(12).WrapText = $false
    $sheet3.Columns.Item(12).ColumnWidth = 35

    $sheet3.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
} else {
    $sheet3.Name = "Profiles"
    $sheet3.Cells.Item(1,1) = "No Structural Profile Data Found"
}


# --- TAB 4: HARDWARE & FASTENERS (BOLTS, NUTS, WASHERS, ANCHORS) ---
$sheet4 = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
$hardwareGroups = $masterData | Where-Object { $_.Category -eq "Hardware" } | Group-Object -Property ($partGroupKeys + @('Size', 'Description'))

if ($hardwareGroups.Count -gt 0) {
    Write-Log "Building Tab 4: Hardware & Fastener Takeoff..."
    $sheet4.Name = "Hardware & Fasteners"

    $hHeaders = @("TL #", "Fastener / Item Mark", "Type / Shape", "Size / Spec", "Grade", "Total Qty", "Location / Notes", "Description", "Unit Wt (lbs)", "Total Wt (lbs)", "SourceFile", "Used In Assemblies")
    for ($c = 0; $c -lt $hHeaders.Count; $c++) { $sheet4.Cells.Item(1, $c + 1) = $hHeaders[$c] }

    $rolledHardware = @($(foreach ($g in $hardwareGroups) {
        $first = $g.Group[0]
        [PSCustomObject]@{
            PartMark    = $first.PartMark
            Shape       = $first.Shape
            Size        = $first.Size
            Grade       = $first.Grade
            TotalQty    = ($g.Group | Measure-Object -Property Qty -Sum).Sum
            Description = $first.Description
            Finish      = $first.Finish
            Weight      = $first.Weight
            SourceFile  = ($g.Group | ForEach-Object { $_.SourceFile } | Select-Object -Unique) -join ', '
            Assemblies  = ($g.Group.Assembly | Select-Object -Unique) -join ', '
            TL          = ($g.Group | ForEach-Object { $_.TL } | Where-Object { $_ } | Select-Object -Unique) -join ', '
        }
    }) | Sort-Object Shape, Size, PartMark)

    $hCount = $rolledHardware.Count
    $hData = New-Object "object[,]" $hCount, $hHeaders.Count

    for ($i = 0; $i -lt $hCount; $i++) {
        $r = $rolledHardware[$i]
        $rowNum = $i + 2

        $hData[$i, 0]  = $r.TL
        $hData[$i, 1]  = $r.PartMark
        $hData[$i, 2]  = $r.Shape
        $hData[$i, 3]  = $r.Size
        $hData[$i, 4]  = $r.Grade
        $hData[$i, 5]  = $r.TotalQty
        $hData[$i, 6]  = if ($r.Description) { $r.Description } else { $r.Finish }
        $hData[$i, 7]  = "$($r.Shape) $($r.Size)"
        $hData[$i, 8]  = $r.Weight
        $hData[$i, 9]  = "=F$rowNum*I$rowNum"
        $hData[$i, 10] = $r.SourceFile
        $hData[$i, 11] = $r.Assemblies
    }

    $sheet4.Range("A2").Resize($hCount, $hHeaders.Count).Value2 = $hData
    $sheet4.Range("I2:J$($hCount + 1)").NumberFormat = "#,##0.00"

    $hRange = $sheet4.Range($sheet4.Cells.Item(1, 1), $sheet4.Cells.Item($hCount + 1, $hHeaders.Count))
    $hTable = $sheet4.ListObjects.Add(1, $hRange, $null, 1)
    $hTable.Name = "HardwareTakeoffTable"
    $hTable.TableStyle = "TableStyleMedium2"
    $hTable.ShowTotals = $true
    
    $hTable.ListColumns.Item("Total Qty").TotalsCalculation = 1
    $hTable.ListColumns.Item("Total Wt (lbs)").TotalsCalculation = 1
    $hRange.EntireColumn.AutoFit()
    $sheet4.Columns.Item(12).WrapText = $false
    $sheet4.Columns.Item(12).ColumnWidth = 35

    $sheet4.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
} else {
    $sheet4.Name = "Hardware & Fasteners"
    $sheet4.Cells.Item(1,1) = "No Hardware or Fastener Data Found"
}


# --- TAB 5: SHOP ASSEMBLY BREAKDOWN ---
Write-Log "Building Tab 5: Shop Assembly Breakdown..."
$sheet5 = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
$sheet5.Name = "Shop Assemblies"

$aRow = 1
$groupedAssms = $masterData | Group-Object -Property $assemblyGroupKeys | Sort-Object { $_.Group[0].Assembly }

foreach ($g in $groupedAssms) {
    $assmName = $g.Group[0].Assembly
    $dwgNo = if ($assemblyDrawings.ContainsKey($assmName)) { $assemblyDrawings[$assmName] } else { "-" }
    $sourceFiles = @($g.Group | ForEach-Object { $_.SourceFile } | Select-Object -Unique)
    $sourceText = if ($sourceFiles.Count -gt 0) { ($sourceFiles -join ', ') } else { "-" }
    $tlText = (($g.Group | ForEach-Object { $_.TL } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
    $varianceNote = ""
    $hasVariance = $false
    $distinctSpecRows = @()
    $variantSources = @()
    if ($sourceFiles.Count -gt 1) {
        $partMarkVarianceGroups = @($g.Group | Group-Object -Property PartMark | Where-Object {
            $markFiles = @($_.Group | ForEach-Object { $_.SourceFile } | Select-Object -Unique)
            $markFiles.Count -gt 1 -and (@($_.Group | ForEach-Object { "$($_.Shape)|$($_.Size)|$($_.Grade)|$($_.Length)|$($_.Description)|$($_.Qty)" } | Select-Object -Unique).Count -gt 1)
        })
        $hasVariance = ($partMarkVarianceGroups.Count -gt 0)

        if ($hasVariance) {
            $distinctSpecRows = @($partMarkVarianceGroups | ForEach-Object { $_.Group | ForEach-Object { "$($_.Shape)|$($_.Size)|$($_.Grade)|$($_.Length)|$($_.Description)" } } | Select-Object -Unique)
            $varianceNote = "VARIANCE: part definition differs across KSS files. " + (($distinctSpecRows | Select-Object -First 3) -join '; ')
            Write-Log "Assembly $assmName appears in multiple KSS files with different definitions: $sourceText | $varianceNote" "Yellow"
            $canonicalSource = @($sourceFiles | Sort-Object) | Select-Object -First 1
            if ($canonicalSource) {
                $variantSources = @($sourceFiles | Sort-Object | Select-Object -Unique | Where-Object { $_ -ne $canonicalSource })
            }
        }
        else {
            $varianceNote = "KSS sources match across files: $sourceText"
        }
    }
    
    $sheet5.Cells.Item($aRow, 1) = "ASSEMBLY: $assmName"
    $sheet5.Cells.Item($aRow, 2) = "DRAWING: $dwgNo"
    $sheet5.Cells.Item($aRow, 3) = "TOTAL PARTS: $($g.Group.Count)"
    $sheet5.Cells.Item($aRow, 4) = "TL #: $tlText"
    $sheet5.Cells.Item($aRow, 5) = "KSS FILES: $sourceText"
    $sheet5.Cells.Item($aRow, 8) = if ($varianceNote) { $varianceNote } else { "KSS FILES MATCH ACROSS FILES" }
    
    $bannerRange = $sheet5.Range("A$aRow", "I$aRow")
    $bannerRange.Interior.Color = Get-OleColor 31 78 120
    $bannerRange.Font.Color = Get-OleColor 255 255 255
    $bannerRange.Font.Bold = $true

    if ($variantSources.Count -gt 0) {
        $variantBlockIndex = 0
        foreach ($variantSource in $variantSources) {
            $variantBlockIndex++
            $variantBannerStartCol = 11 + (($variantBlockIndex - 1) * 10)
            $variantBannerEndCol = $variantBannerStartCol + 8
            $variantBannerRange = $sheet5.Range($sheet5.Cells.Item($aRow, $variantBannerStartCol), $sheet5.Cells.Item($aRow, $variantBannerEndCol))
            $sheet5.Cells.Item($aRow, $variantBannerStartCol) = "ASSEMBLY: $assmName"
            $sheet5.Cells.Item($aRow, $variantBannerStartCol + 1) = "DRAWING: $dwgNo"
            $sheet5.Cells.Item($aRow, $variantBannerStartCol + 2) = "TOTAL PARTS: $($g.Group.Count)"
            $sheet5.Cells.Item($aRow, $variantBannerStartCol + 3) = "TL #: $tlText"
            $sheet5.Cells.Item($aRow, $variantBannerStartCol + 4) = "KSS FILES: $sourceText"
            $sheet5.Cells.Item($aRow, $variantBannerStartCol + 7) = if ($varianceNote) { $varianceNote } else { "KSS FILES MATCH ACROSS FILES" }
            $variantBannerRange.Interior.Color = Get-OleColor 31 78 120
            $variantBannerRange.Font.Color = Get-OleColor 255 255 255
            $variantBannerRange.Font.Bold = $true
        }
    }
    $aRow++

    $subHeaders = @("TL #", "PartMark", "Qty", "Category", "Shape", "Size", "Length", "Grade", "Description")
    for ($sc = 0; $sc -lt $subHeaders.Count; $sc++) {
        $sheet5.Cells.Item($aRow, $sc + 1) = $subHeaders[$sc]
    }
    $sheet5.Range("A$aRow", "I$aRow").Font.Bold = $true
    $sheet5.Range("A$aRow", "I$aRow").Interior.Color = Get-OleColor 220 225 230

    if ($variantSources.Count -gt 0) {
        $variantBlockIndex = 0
        foreach ($variantSource in $variantSources) {
            $variantBlockIndex++
            $variantHeaderStartCol = 11 + (($variantBlockIndex - 1) * 10)
            $variantHeaderEndCol = $variantHeaderStartCol + 8
            $variantHeaderRange = $sheet5.Range($sheet5.Cells.Item($aRow, $variantHeaderStartCol), $sheet5.Cells.Item($aRow, $variantHeaderEndCol))
            for ($sc = 0; $sc -lt $subHeaders.Count; $sc++) {
                $sheet5.Cells.Item($aRow, $variantHeaderStartCol + $sc) = $subHeaders[$sc]
            }
            $variantHeaderRange.Font.Bold = $true
            $variantHeaderRange.Interior.Color = Get-OleColor 220 225 230
        }
    }
    $aRow++

    $displayGroup = @($g.Group)
    $canonicalSource = $null
    if ($sourceFiles.Count -gt 1) {
        $canonicalSource = @($sourceFiles | Sort-Object) | Select-Object -First 1
        if ($canonicalSource) {
            $displayGroup = @($g.Group | Where-Object { $_.SourceFile -eq $canonicalSource })
        }
    }
    if ($displayGroup.Count -eq 0) { $displayGroup = @($g.Group) }

    foreach ($p in $displayGroup) {
        $sheet5.Cells.Item($aRow, 1) = $p.TL
        $sheet5.Cells.Item($aRow, 2) = $p.PartMark
        $sheet5.Cells.Item($aRow, 3) = $p.Qty
        $sheet5.Cells.Item($aRow, 4) = $p.Category
        $sheet5.Cells.Item($aRow, 5) = $p.Shape
        $sheet5.Cells.Item($aRow, 6) = $p.Size
        $sheet5.Cells.Item($aRow, 7) = $p.Length
        $sheet5.Cells.Item($aRow, 8) = $p.Grade
        $sheet5.Cells.Item($aRow, 9) = $p.Description

        if ($variantSources.Count -gt 0) {
            $variantCounter = 0
            foreach ($variantSource in $variantSources) {
                $variantCounter++
                $variantRows = @($g.Group | Where-Object { $_.SourceFile -eq $variantSource } | Sort-Object PartMark)
                $variantMatch = $variantRows | Where-Object { $_.PartMark -eq $p.PartMark } | Select-Object -First 1
                if (-not $variantMatch) { $variantMatch = $variantRows | Select-Object -First 1 }

                $compareStartCol = 11 + (($variantCounter - 1) * 10)
                $compareEndCol = $compareStartCol + 8
                $leftCols = @("B", "C", "D", "E", "F", "G", "H", "I")
                $compareCols = @()
                for ($cc = $compareStartCol; $cc -le $compareEndCol; $cc++) {
                    $compareCols += Get-ExcelColumnLetter($cc)
                }
                $compareValueCols = @($compareCols[1..8])

                $sheet5.Cells.Item($aRow, $compareStartCol) = if ($variantMatch) { $variantMatch.TL } else { "NO MATCH" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 1) = if ($variantMatch) { $variantMatch.PartMark } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 2) = if ($variantMatch) { $variantMatch.Qty } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 3) = if ($variantMatch) { $variantMatch.Category } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 4) = if ($variantMatch) { $variantMatch.Shape } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 5) = if ($variantMatch) { $variantMatch.Size } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 6) = if ($variantMatch) { $variantMatch.Length } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 7) = if ($variantMatch) { $variantMatch.Grade } else { "-" }
                $sheet5.Cells.Item($aRow, $compareStartCol + 8) = if ($variantMatch) { $variantMatch.Description } else { "-" }

                $compareRange = $sheet5.Range($sheet5.Cells.Item($aRow, $compareStartCol), $sheet5.Cells.Item($aRow, $compareEndCol))
                $compareRange.Interior.Color = Get-OleColor 255 255 255
                $compareRange.Font.Color = Get-OleColor 44 62 80
                $compareRange.Borders.LineStyle = 1
                $compareRange.Borders.Weight = 2
                $compareRange.Borders.Color = Get-OleColor 200 200 200

                $formulaParts = @()
                for ($idx = 0; $idx -lt $leftCols.Count; $idx++) {
                    $formulaParts += '$' + $leftCols[$idx] + '$' + $aRow + '=$' + $compareValueCols[$idx] + '$' + $aRow
                }

                $matchCondition = "AND(" + ($formulaParts -join ',') + ")"
                $matchFormula = "=" + $matchCondition
                $matchFormat = $compareRange.FormatConditions.Add(2, 0, $matchFormula)
                $matchFormat.Interior.Color = Get-OleColor 198 239 206
                $matchFormat.Font.Color = Get-OleColor 0 97 0
                $matchFormat.Font.Bold = $true

                $diffFormula = "=NOT(" + $matchCondition + ")"
                $diffFormat = $compareRange.FormatConditions.Add(2, 0, $diffFormula)
                $diffFormat.Interior.Color = Get-OleColor 255 199 206
                $diffFormat.Font.Color = Get-OleColor 128 0 0
                $diffFormat.Font.Bold = $true
            }
        }

        $aRow++
    }
    $aRow++
}
$sheet5.Columns.AutoFit()


# --- TAB 6: LABOR & WELD BREAKDOWN (HOLES, CUTS, WELDS) ---
$sheet6 = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
$sheet6.Name = "Labor & Weld Operations"

if ($laborData.Count -gt 0) {
    Write-Log "Building Tab 6: Labor & Weld Operations Breakdown..."
    $lHeaders = @("TL #", "Assembly", "PartMark", "Labor Type", "Qty", "Size 1 / Diameter", "Size 2 / Thickness", "Notes / Profile", "Source File")
    for ($c = 0; $c -lt $lHeaders.Count; $c++) { $sheet6.Cells.Item(1, $c + 1) = $lHeaders[$c] }

    $lRows = $laborData.Count
    $lDataArr = New-Object "object[,]" $lRows, $lHeaders.Count

    for ($i = 0; $i -lt $lRows; $i++) {
        $item = $laborData[$i]
        $lDataArr[$i, 0] = $item.TL
        $lDataArr[$i, 1] = $item.Assembly
        $lDataArr[$i, 2] = $item.PartMark
        $lDataArr[$i, 3] = $item.LaborType
        $lDataArr[$i, 4] = $item.Qty
        $lDataArr[$i, 5] = $item.Size1
        $lDataArr[$i, 6] = $item.Size2
        $lDataArr[$i, 7] = $item.Notes
        $lDataArr[$i, 8] = $item.SourceFile
    }

    $sheet6.Range("A2").Resize($lRows, $lHeaders.Count).Value2 = $lDataArr
    $lRange = $sheet6.Range($sheet6.Cells.Item(1, 1), $sheet6.Cells.Item($lRows + 1, $lHeaders.Count))
    $lTable = $sheet6.ListObjects.Add(1, $lRange, $null, 1)
    $lTable.Name = "LaborOperationsTable"
    $lTable.TableStyle = "TableStyleMedium2"
    $lTable.ShowTotals = $true
    $lTable.ListColumns.Item("Qty").TotalsCalculation = 1
    $lRange.EntireColumn.AutoFit()

    $sheet6.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
} else {
    $sheet6.Cells.Item(1, 1) = "No Labor Record Data Found in KSS Input"
}


# --- TAB 0: PROJECT STATISTICS (Excel formulas only) ---
Write-Log "Building front statistics dashboard..."
$dashboard = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
$dashboard.Name = "Statistics"

# Determine last data rows across sheets for clean formula ranges
$lastMbRow = 2
if ($mRows -gt 0) { $lastMbRow = $mRows + 1 }

$lastPlRow = 2
if ($pCount -gt 0) { $lastPlRow = $pCount + 1 }

$lastPrRow = 2
if ($prCount -gt 0) { $lastPrRow = $prCount + 1 }

$lastHwRow = 2
if ($hCount -gt 0) { $lastHwRow = $hCount + 1 }

$lastLbRow = 2
if ($lRows -gt 0) { $lastLbRow = $lRows + 1 }

# =========================================================================
# TITLE BANNER & EXECUTIVE HEADER (Rows 1 - 5)
# =========================================================================
$dashboard.Range("A1:W1").Merge()
$dashboard.Cells.Item(1, 1) = "PROJECT FABRICATION, ESTIMATING & PRODUCTION DASHBOARD"
$dashboard.Range("A1:W1").Font.Size = 16
$dashboard.Range("A1:W1").Font.Bold = $true
$dashboard.Range("A1:W1").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("A1:W1").Interior.Color = Get-OleColor 27 54 93
$dashboard.Range("A1:W1").HorizontalAlignment = -4108
$dashboard.Range("A1:W1").VerticalAlignment = -4108
$dashboard.Range("A1:W1").RowHeight = 32

$dashboard.Range("A2:W2").Merge()
$dashboard.Cells.Item(2, 1) = "100% Live Excel Equations | Real-time Rollups, Category Takeoffs, TL Sequencing, Steel Grades, Nesting Audits & KPI Ratios"
$dashboard.Range("A2:W2").Font.Size = 9
$dashboard.Range("A2:W2").Font.Italic = $true
$dashboard.Range("A2:W2").Font.Color = Get-OleColor 230 235 245
$dashboard.Range("A2:W2").Interior.Color = Get-OleColor 44 82 130
$dashboard.Range("A2:W2").HorizontalAlignment = -4108
$dashboard.Range("A2:W2").VerticalAlignment = -4108
$dashboard.Range("A2:W2").RowHeight = 18

# Executive Information Cards (Row 4 - 5)
$dashboard.Cells.Item(4, 1) = "PROJECT NAME"
$dashboard.Cells.Item(5, 1) = $projectName
$dashboard.Range("A4:B4").Merge()
$dashboard.Range("A5:B5").Merge()

$dashboard.Cells.Item(4, 3) = "JOB NUMBER"
$dashboard.Cells.Item(5, 3) = $jobNumber
$dashboard.Range("C4:D4").Merge()
$dashboard.Range("C5:D5").Merge()

$dashboard.Cells.Item(4, 5) = "TRANSMITTALS (TL)"
$dashboard.Cells.Item(5, 5) = if ($allTLs.Count -gt 0) { ($allTLs -join ', ') } else { "ALL" }
$dashboard.Range("E4:H4").Merge()
$dashboard.Range("E5:H5").Merge()

$dashboard.Cells.Item(4, 9) = "TIMESTAMP GENERATED"
$dashboard.Cells.Item(5, 9) = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$dashboard.Range("I4:K4").Merge()
$dashboard.Range("I5:K5").Merge()

$dashboard.Cells.Item(4, 12) = "CALCULATION ENGINE"
$dashboard.Cells.Item(5, 12) = "EXCEL FORMULA DRIVEN"
$dashboard.Range("L4:N4").Merge()
$dashboard.Range("L5:N5").Merge()

$dashboard.Range("A4:N4").Font.Bold = $true
$dashboard.Range("A4:N4").Font.Size = 9
$dashboard.Range("A4:N4").Font.Color = Get-OleColor 100 110 120
$dashboard.Range("A4:N4").Interior.Color = Get-OleColor 240 243 246
$dashboard.Range("A4:N4").HorizontalAlignment = -4108
$dashboard.Range("A4:N4").Borders.LineStyle = 1
$dashboard.Range("A4:N4").Borders.Weight = 2
$dashboard.Range("A4:N4").Borders.Color = Get-OleColor 210 215 220

$dashboard.Range("A5:N5").Font.Bold = $true
$dashboard.Range("A5:N5").Font.Size = 10
$dashboard.Range("A5:N5").Font.Color = Get-OleColor 20 40 70
$dashboard.Range("A5:N5").Interior.Color = Get-OleColor 255 255 255
$dashboard.Range("A5:N5").HorizontalAlignment = -4108
$dashboard.Range("A5:N5").Borders.LineStyle = 1
$dashboard.Range("A5:N5").Borders.Weight = 2
$dashboard.Range("A5:N5").Borders.Color = Get-OleColor 210 215 220
$dashboard.Range("A5:N5").RowHeight = 22

# =========================================================================
# SECTION 1: MASTER BOM TOTALS & STATS (Cols A-C, Rows 7-26)
# =========================================================================
$dashboard.Range("A7:C7").Merge()
$dashboard.Cells.Item(7, 1) = "MASTER BOM METRICS & TOTALS"
$dashboard.Range("A7:C7").Font.Bold = $true
$dashboard.Range("A7:C7").Font.Size = 10
$dashboard.Range("A7:C7").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("A7:C7").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("A7:C7").HorizontalAlignment = -4108

$dashboard.Cells.Item(8, 1) = "Metric Description"
$dashboard.Cells.Item(8, 2) = "Formula Value"
$dashboard.Cells.Item(8, 3) = "Source / Scope"
$dashboard.Range("A8:C8").Font.Bold = $true
$dashboard.Range("A8:C8").Font.Size = 9
$dashboard.Range("A8:C8").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("A8:C8").Borders.LineStyle = 1
$dashboard.Range("A8:C8").Borders.Weight = 2
$dashboard.Range("A8:C8").Borders.Color = Get-OleColor 166 166 166

$masterStats = @(
    @{ Name = "Total Detail Line Items"; Formula = "=COUNTA('Master BOM'!`$B`$2:`$B`$$lastMbRow)"; Tab = "Master BOM (D rows)"; Fmt = "#,##0" },
    @{ Name = "Total Part Piece Count"; Formula = "=SUM('Master BOM'!`$F`$2:`$F`$$lastMbRow)"; Tab = "Master BOM [Qty]"; Fmt = "#,##0" },
    @{ Name = "Total Material Weight (lbs)"; Formula = "=SUM('Master BOM'!`$N`$2:`$N`$$lastMbRow)"; Tab = "Master BOM [Total Wt]"; Fmt = "#,##0.00" },
    @{ Name = "Total Material Weight (Tons)"; Formula = "=B11/2000"; Tab = "Total lbs / 2000"; Fmt = "#,##0.000" },
    @{ Name = "Distinct Assemblies Count"; Formula = "=COUNTA(UNIQUE('Master BOM'!`$B`$2:`$B`$$lastMbRow))"; Tab = "Master BOM [Assembly]"; Fmt = "#,##0" },
    @{ Name = "Distinct Part Marks Count"; Formula = "=COUNTA(UNIQUE('Master BOM'!`$E`$2:`$E`$$lastMbRow))"; Tab = "Master BOM [PartMark]"; Fmt = "#,##0" },
    @{ Name = "Distinct Material Shapes"; Formula = "=COUNTA(UNIQUE('Master BOM'!`$H`$2:`$H`$$lastMbRow))"; Tab = "Master BOM [Shape]"; Fmt = "#,##0" },
    @{ Name = "Distinct Steel Grades"; Formula = "=COUNTA(UNIQUE('Master BOM'!`$K`$2:`$K`$$lastMbRow))"; Tab = "Master BOM [Grade]"; Fmt = "#,##0" },
    @{ Name = "Distinct Transmittals (TL #)"; Formula = "=COUNTA(UNIQUE('Master BOM'!`$A`$2:`$A`$$lastMbRow))"; Tab = "Master BOM [TL #]"; Fmt = "#,##0" },
    @{ Name = "Distinct Source KSS Files"; Formula = "=COUNTA(UNIQUE('Master BOM'!`$P`$2:`$P`$$lastMbRow))"; Tab = "Master BOM [Source]"; Fmt = "#,##0" },
    @{ Name = "Average Part Quantity / Line"; Formula = "=AVERAGE('Master BOM'!`$F`$2:`$F`$$lastMbRow)"; Tab = "Master BOM [Qty]"; Fmt = "#,##0.0" },
    @{ Name = "Maximum Quantity on Single Line"; Formula = "=MAX('Master BOM'!`$F`$2:`$F`$$lastMbRow)"; Tab = "Master BOM [Qty]"; Fmt = "#,##0" },
    @{ Name = "Minimum Quantity on Single Line"; Formula = "=MIN('Master BOM'!`$F`$2:`$F`$$lastMbRow)"; Tab = "Master BOM [Qty]"; Fmt = "#,##0" },
    @{ Name = "Average Unit Weight (lbs)"; Formula = "=AVERAGE('Master BOM'!`$M`$2:`$M`$$lastMbRow)"; Tab = "Master BOM [Unit Wt]"; Fmt = "#,##0.00" },
    @{ Name = "Maximum Unit Weight (lbs)"; Formula = "=MAX('Master BOM'!`$M`$2:`$M`$$lastMbRow)"; Tab = "Master BOM [Unit Wt]"; Fmt = "#,##0.00" },
    @{ Name = "Maximum Extended Weight (lbs)"; Formula = "=MAX('Master BOM'!`$N`$2:`$N`$$lastMbRow)"; Tab = "Master BOM [Total Wt]"; Fmt = "#,##0.00" },
    @{ Name = "Avg Weight per Assembly (lbs)"; Formula = "=B11/B13"; Tab = "Total lbs / Assms"; Fmt = "#,##0.00" },
    @{ Name = "Avg Parts per Assembly"; Formula = "=B10/B13"; Tab = "Total Pcs / Assms"; Fmt = "#,##0.0" }
)

$rIdx = 9
foreach ($ms in $masterStats) {
    $dashboard.Cells.Item($rIdx, 1) = $ms.Name
    $dashboard.Cells.Item($rIdx, 2).Formula = $ms.Formula
    $dashboard.Cells.Item($rIdx, 3) = $ms.Tab
    $dashboard.Cells.Item($rIdx, 2).NumberFormat = $ms.Fmt
    $dashboard.Range("A$rIdx:C$rIdx").Borders.LineStyle = 1
    $dashboard.Range("A$rIdx:C$rIdx").Borders.Weight = 2
    $dashboard.Range("A$rIdx:C$rIdx").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($rIdx, 1).Interior.Color = Get-OleColor 248 249 251
    $rIdx++
}

# =========================================================================
# SECTION 2: WORKBOOK SHEET ROLLUP & AUDIT (Cols E-H, Rows 7-15)
# =========================================================================
$dashboard.Range("E7:H7").Merge()
$dashboard.Cells.Item(7, 5) = "WORKBOOK SHEET ROLLUP & CROSS-AUDIT"
$dashboard.Range("E7:H7").Font.Bold = $true
$dashboard.Range("E7:H7").Font.Size = 10
$dashboard.Range("E7:H7").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("E7:H7").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("E7:H7").HorizontalAlignment = -4108

$dashboard.Cells.Item(8, 5) = "Sheet / Department"
$dashboard.Cells.Item(8, 6) = "Line Items"
$dashboard.Cells.Item(8, 7) = "Total Qty"
$dashboard.Cells.Item(8, 8) = "Total Wt (lbs)"
$dashboard.Range("E8:H8").Font.Bold = $true
$dashboard.Range("E8:H8").Font.Size = 9
$dashboard.Range("E8:H8").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("E8:H8").Borders.LineStyle = 1
$dashboard.Range("E8:H8").Borders.Weight = 2
$dashboard.Range("E8:H8").Borders.Color = Get-OleColor 166 166 166

$sheetRollups = @(
    @{ Name = "Plates & Sheets Nesting"; RowFml = "=COUNTA('Plates & Sheets (PL-SHT-BPL)'!`$C`$2:`$C`$$lastPlRow)"; QtyFml = "=SUM('Plates & Sheets (PL-SHT-BPL)'!`$E`$2:`$E`$$lastPlRow)"; WtFml = "=SUM('Plates & Sheets (PL-SHT-BPL)'!`$K`$2:`$K`$$lastPlRow)" },
    @{ Name = "Structural Profiles"; RowFml = "=COUNTA('Profiles (W-HSS-L)'!`$C`$2:`$C`$$lastPrRow)"; QtyFml = "=SUM('Profiles (W-HSS-L)'!`$D`$2:`$D`$$lastPrRow)"; WtFml = "=SUM('Profiles (W-HSS-L)'!`$J`$2:`$J`$$lastPrRow)" },
    @{ Name = "Hardware & Fasteners"; RowFml = "=COUNTA('Hardware & Fasteners'!`$B`$2:`$B`$$lastHwRow)"; QtyFml = "=SUM('Hardware & Fasteners'!`$F`$2:`$F`$$lastHwRow)"; WtFml = "=SUM('Hardware & Fasteners'!`$J`$2:`$J`$$lastHwRow)" },
    @{ Name = "Labor & Operations"; RowFml = "=COUNTA('Labor & Weld Operations'!`$C`$2:`$C`$$lastLbRow)"; QtyFml = "=SUM('Labor & Weld Operations'!`$E`$2:`$E`$$lastLbRow)"; WtFml = 'NA' }
)

$srRow = 9
foreach ($sr in $sheetRollups) {
    $dashboard.Cells.Item($srRow, 5) = $sr.Name
    $dashboard.Cells.Item($srRow, 6).Formula = $sr.RowFml
    $dashboard.Cells.Item($srRow, 7).Formula = $sr.QtyFml
    if ($sr.WtFml -eq 'NA') {
        $dashboard.Cells.Item($srRow, 8) = "-"
    } else {
        $dashboard.Cells.Item($srRow, 8).Formula = $sr.WtFml
        $dashboard.Cells.Item($srRow, 8).NumberFormat = "#,##0.00"
    }
    $dashboard.Cells.Item($srRow, 6).NumberFormat = "#,##0"
    $dashboard.Cells.Item($srRow, 7).NumberFormat = "#,##0"
    $dashboard.Range("E$srRow:H$srRow").Borders.LineStyle = 1
    $dashboard.Range("E$srRow:H$srRow").Borders.Weight = 2
    $dashboard.Range("E$srRow:H$srRow").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($srRow, 5).Interior.Color = Get-OleColor 248 249 251
    $srRow++
}

# Total Audit Row
$dashboard.Cells.Item($srRow, 5) = "TOTAL AUDITED SUM"
$dashboard.Cells.Item($srRow, 6).Formula = "=SUM(F9:F12)"
$dashboard.Cells.Item($srRow, 7).Formula = "=SUM(G9:G11)"
$dashboard.Cells.Item($srRow, 8).Formula = "=SUM(H9:H11)"
$dashboard.Range("E$srRow:H$srRow").Font.Bold = $true
$dashboard.Range("E$srRow:H$srRow").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($srRow, 6).NumberFormat = "#,##0"
$dashboard.Cells.Item($srRow, 7).NumberFormat = "#,##0"
$dashboard.Cells.Item($srRow, 8).NumberFormat = "#,##0.00"
$dashboard.Range("E$srRow:H$srRow").Borders.LineStyle = 1
$dashboard.Range("E$srRow:H$srRow").Borders.Weight = 2
$dashboard.Range("E$srRow:H$srRow").Borders.Color = Get-OleColor 166 166 166

# Variance Check Row
$auditRow = $srRow + 1
$dashboard.Cells.Item($auditRow, 5) = "Variance vs Master BOM"
$dashboard.Cells.Item($auditRow, 6) = "-"
$dashboard.Cells.Item($auditRow, 7).Formula = "=G$srRow-B10"
$dashboard.Cells.Item($auditRow, 8).Formula = "=H$srRow-B11"
$dashboard.Range("E$auditRow:H$auditRow").Font.Italic = $true
$dashboard.Cells.Item($auditRow, 7).NumberFormat = "#,##0"
$dashboard.Cells.Item($auditRow, 8).NumberFormat = "#,##0.00"
$dashboard.Range("E$auditRow:H$auditRow").Borders.LineStyle = 1
$dashboard.Range("E$auditRow:H$auditRow").Borders.Weight = 2
$dashboard.Range("E$auditRow:H$auditRow").Borders.Color = Get-OleColor 220 225 230

# =========================================================================
# SECTION 3: MATERIAL CATEGORY TAKEOFF & SHARES (Cols E-H, Rows 16-24)
# =========================================================================
$dashboard.Range("E16:H16").Merge()
$dashboard.Cells.Item(16, 5) = "MATERIAL CATEGORY TAKEOFF & WEIGHT SHARES"
$dashboard.Range("E16:H16").Font.Bold = $true
$dashboard.Range("E16:H16").Font.Size = 10
$dashboard.Range("E16:H16").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("E16:H16").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("E16:H16").HorizontalAlignment = -4108

$dashboard.Cells.Item(17, 5) = "Category"
$dashboard.Cells.Item(17, 6) = "Line Items"
$dashboard.Cells.Item(17, 7) = "Piece Qty"
$dashboard.Cells.Item(17, 8) = "Total Wt (lbs)"
$dashboard.Range("E17:H17").Font.Bold = $true
$dashboard.Range("E17:H17").Font.Size = 9
$dashboard.Range("E17:H17").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("E17:H17").Borders.LineStyle = 1
$dashboard.Range("E17:H17").Borders.Weight = 2
$dashboard.Range("E17:H17").Borders.Color = Get-OleColor 166 166 166

$catList = @("Plate", "Profile", "Hardware")
$cRow = 18
foreach ($cName in $catList) {
    $dashboard.Cells.Item($cRow, 5) = $cName
    $dashboard.Cells.Item($cRow, 6).Formula = "=COUNTIF('Master BOM'!`$G`$2:`$G`$$lastMbRow, E$cRow)"
    $dashboard.Cells.Item($cRow, 7).Formula = "=SUMIF('Master BOM'!`$G`$2:`$G`$$lastMbRow, E$cRow, 'Master BOM'!`$F`$2:`$F`$$lastMbRow)"
    $dashboard.Cells.Item($cRow, 8).Formula = "=SUMIF('Master BOM'!`$G`$2:`$G`$$lastMbRow, E$cRow, 'Master BOM'!`$N`$2:`$N`$$lastMbRow)"
    $dashboard.Cells.Item($cRow, 6).NumberFormat = "#,##0"
    $dashboard.Cells.Item($cRow, 7).NumberFormat = "#,##0"
    $dashboard.Cells.Item($cRow, 8).NumberFormat = "#,##0.00"
    $dashboard.Range("E$cRow:H$cRow").Borders.LineStyle = 1
    $dashboard.Range("E$cRow:H$cRow").Borders.Weight = 2
    $dashboard.Range("E$cRow:H$cRow").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($cRow, 5).Interior.Color = Get-OleColor 248 249 251
    $cRow++
}

# Category Total Row
$dashboard.Cells.Item($cRow, 5) = "TOTAL MATERIAL"
$dashboard.Cells.Item($cRow, 6).Formula = "=SUM(F18:F20)"
$dashboard.Cells.Item($cRow, 7).Formula = "=SUM(G18:G20)"
$dashboard.Cells.Item($cRow, 8).Formula = "=SUM(H18:H20)"
$dashboard.Range("E$cRow:H$cRow").Font.Bold = $true
$dashboard.Range("E$cRow:H$cRow").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($cRow, 6).NumberFormat = "#,##0"
$dashboard.Cells.Item($cRow, 7).NumberFormat = "#,##0"
$dashboard.Cells.Item($cRow, 8).NumberFormat = "#,##0.00"
$dashboard.Range("E$cRow:H$cRow").Borders.LineStyle = 1
$dashboard.Range("E$cRow:H$cRow").Borders.Weight = 2
$dashboard.Range("E$cRow:H$cRow").Borders.Color = Get-OleColor 166 166 166

# Category Share Rows (Weight % and Pcs %)
$cShareRow = $cRow + 1
$dashboard.Cells.Item($cShareRow, 5) = "Weight Distribution %"
$dashboard.Cells.Item($cShareRow, 6) = "-"
$dashboard.Cells.Item($cShareRow, 7) = "-"
$dashboard.Cells.Item($cShareRow, 8).Formula = "=IF(B11>0, H$cRow/B11, 0)"
$dashboard.Cells.Item($cShareRow, 8).NumberFormat = "0.0%"
$dashboard.Range("E$cShareRow:H$cShareRow").Font.Italic = $true
$dashboard.Range("E$cShareRow:H$cShareRow").Borders.LineStyle = 1
$dashboard.Range("E$cShareRow:H$cShareRow").Borders.Weight = 2
$dashboard.Range("E$cShareRow:H$cShareRow").Borders.Color = Get-OleColor 220 225 230

$cPcsShareRow = $cShareRow + 1
$dashboard.Cells.Item($cPcsShareRow, 5) = "Piece Count Distribution %"
$dashboard.Cells.Item($cPcsShareRow, 6) = "-"
$dashboard.Cells.Item($cPcsShareRow, 7).Formula = "=IF(B10>0, G$cRow/B10, 0)"
$dashboard.Cells.Item($cPcsShareRow, 8) = "-"
$dashboard.Cells.Item($cPcsShareRow, 7).NumberFormat = "0.0%"
$dashboard.Range("E$cPcsShareRow:H$cPcsShareRow").Font.Italic = $true
$dashboard.Range("E$cPcsShareRow:H$cPcsShareRow").Borders.LineStyle = 1
$dashboard.Range("E$cPcsShareRow:H$cPcsShareRow").Borders.Weight = 2
$dashboard.Range("E$cPcsShareRow:H$cPcsShareRow").Borders.Color = Get-OleColor 220 225 230

# =========================================================================
# SECTION 4: TRANSMITTAL (TL #) BREAKDOWN (Cols J-M, Rows 7-26)
# =========================================================================
$dashboard.Range("J7:M7").Merge()
$dashboard.Cells.Item(7, 10) = "TRANSMITTAL (TL #) SEQUENCING & TOTALS"
$dashboard.Range("J7:M7").Font.Bold = $true
$dashboard.Range("J7:M7").Font.Size = 10
$dashboard.Range("J7:M7").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("J7:M7").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("J7:M7").HorizontalAlignment = -4108

$dashboard.Cells.Item(8, 10) = "TL #"
$dashboard.Cells.Item(8, 11) = "Line Items"
$dashboard.Cells.Item(8, 12) = "Total Qty"
$dashboard.Cells.Item(8, 13) = "Total Wt (lbs)"
$dashboard.Range("J8:M8").Font.Bold = $true
$dashboard.Range("J8:M8").Font.Size = 9
$dashboard.Range("J8:M8").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("J8:M8").Borders.LineStyle = 1
$dashboard.Range("J8:M8").Borders.Weight = 2
$dashboard.Range("J8:M8").Borders.Color = Get-OleColor 166 166 166

$displayTLs = if ($allTLs.Count -gt 0) { @($allTLs) } else { @("TL000") }
$tlIdx = 9
foreach ($tlVal in $displayTLs) {
    $dashboard.Cells.Item($tlIdx, 10) = $tlVal
    $dashboard.Cells.Item($tlIdx, 11).Formula = "=COUNTIF('Master BOM'!`$A`$2:`$A`$$lastMbRow, J$tlIdx)"
    $dashboard.Cells.Item($tlIdx, 12).Formula = "=SUMIF('Master BOM'!`$A`$2:`$A`$$lastMbRow, J$tlIdx, 'Master BOM'!`$F`$2:`$F`$$lastMbRow)"
    $dashboard.Cells.Item($tlIdx, 13).Formula = "=SUMIF('Master BOM'!`$A`$2:`$A`$$lastMbRow, J$tlIdx, 'Master BOM'!`$N`$2:`$N`$$lastMbRow)"
    $dashboard.Cells.Item($tlIdx, 11).NumberFormat = "#,##0"
    $dashboard.Cells.Item($tlIdx, 12).NumberFormat = "#,##0"
    $dashboard.Cells.Item($tlIdx, 13).NumberFormat = "#,##0.00"
    $dashboard.Range("J$tlIdx:M$tlIdx").Borders.LineStyle = 1
    $dashboard.Range("J$tlIdx:M$tlIdx").Borders.Weight = 2
    $dashboard.Range("J$tlIdx:M$tlIdx").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($tlIdx, 10).Interior.Color = Get-OleColor 248 249 251
    $tlIdx++
}

# TL Total Row
$dashboard.Cells.Item($tlIdx, 10) = "TOTAL TRANSMITTALS"
$dashboard.Cells.Item($tlIdx, 11).Formula = "=SUM(K9:K$($tlIdx - 1))"
$dashboard.Cells.Item($tlIdx, 12).Formula = "=SUM(L9:L$($tlIdx - 1))"
$dashboard.Cells.Item($tlIdx, 13).Formula = "=SUM(M9:M$($tlIdx - 1))"
$dashboard.Range("J$tlIdx:M$tlIdx").Font.Bold = $true
$dashboard.Range("J$tlIdx:M$tlIdx").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($tlIdx, 11).NumberFormat = "#,##0"
$dashboard.Cells.Item($tlIdx, 12).NumberFormat = "#,##0"
$dashboard.Cells.Item($tlIdx, 13).NumberFormat = "#,##0.00"
$dashboard.Range("J$tlIdx:M$tlIdx").Borders.LineStyle = 1
$dashboard.Range("J$tlIdx:M$tlIdx").Borders.Weight = 2
$dashboard.Range("J$tlIdx:M$tlIdx").Borders.Color = Get-OleColor 166 166 166

# =========================================================================
# SECTION 5: SOURCE KSS FILES BREAKDOWN (Cols O-R, Rows 7-26)
# =========================================================================
$dashboard.Range("O7:R7").Merge()
$dashboard.Cells.Item(7, 15) = "ORIGINATING KSS SOURCE FILES"
$dashboard.Range("O7:R7").Font.Bold = $true
$dashboard.Range("O7:R7").Font.Size = 10
$dashboard.Range("O7:R7").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("O7:R7").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("O7:R7").HorizontalAlignment = -4108

$dashboard.Cells.Item(8, 15) = "KSS File Name"
$dashboard.Cells.Item(8, 16) = "Line Items"
$dashboard.Cells.Item(8, 17) = "Total Qty"
$dashboard.Cells.Item(8, 18) = "Total Wt (lbs)"
$dashboard.Range("O8:R8").Font.Bold = $true
$dashboard.Range("O8:R8").Font.Size = 9
$dashboard.Range("O8:R8").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("O8:R8").Borders.LineStyle = 1
$dashboard.Range("O8:R8").Borders.Weight = 2
$dashboard.Range("O8:R8").Borders.Color = Get-OleColor 166 166 166

$srcFiles = @($masterData | ForEach-Object { $_.SourceFile } | Where-Object { $_ } | Select-Object -Unique)
if ($srcFiles.Count -eq 0) { $srcFiles = @("Source.kss") }
$sfIdx = 9
foreach ($sf in $srcFiles) {
    $dashboard.Cells.Item($sfIdx, 15) = $sf
    $dashboard.Cells.Item($sfIdx, 16).Formula = "=COUNTIF('Master BOM'!`$P`$2:`$P`$$lastMbRow, O$sfIdx)"
    $dashboard.Cells.Item($sfIdx, 17).Formula = "=SUMIF('Master BOM'!`$P`$2:`$P`$$lastMbRow, O$sfIdx, 'Master BOM'!`$F`$2:`$F`$$lastMbRow)"
    $dashboard.Cells.Item($sfIdx, 18).Formula = "=SUMIF('Master BOM'!`$P`$2:`$P`$$lastMbRow, O$sfIdx, 'Master BOM'!`$N`$2:`$N`$$lastMbRow)"
    $dashboard.Cells.Item($sfIdx, 16).NumberFormat = "#,##0"
    $dashboard.Cells.Item($sfIdx, 17).NumberFormat = "#,##0"
    $dashboard.Cells.Item($sfIdx, 18).NumberFormat = "#,##0.00"
    $dashboard.Range("O$sfIdx:R$sfIdx").Borders.LineStyle = 1
    $dashboard.Range("O$sfIdx:R$sfIdx").Borders.Weight = 2
    $dashboard.Range("O$sfIdx:R$sfIdx").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($sfIdx, 15).Interior.Color = Get-OleColor 248 249 251
    $sfIdx++
}

# Source File Total Row
$dashboard.Cells.Item($sfIdx, 15) = "TOTAL KSS SOURCES"
$dashboard.Cells.Item($sfIdx, 16).Formula = "=SUM(P9:P$($sfIdx - 1))"
$dashboard.Cells.Item($sfIdx, 17).Formula = "=SUM(Q9:Q$($sfIdx - 1))"
$dashboard.Cells.Item($sfIdx, 18).Formula = "=SUM(R9:R$($sfIdx - 1))"
$dashboard.Range("O$sfIdx:R$sfIdx").Font.Bold = $true
$dashboard.Range("O$sfIdx:R$sfIdx").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($sfIdx, 16).NumberFormat = "#,##0"
$dashboard.Cells.Item($sfIdx, 17).NumberFormat = "#,##0"
$dashboard.Cells.Item($sfIdx, 18).NumberFormat = "#,##0.00"
$dashboard.Range("O$sfIdx:R$sfIdx").Borders.LineStyle = 1
$dashboard.Range("O$sfIdx:R$sfIdx").Borders.Weight = 2
$dashboard.Range("O$sfIdx:R$sfIdx").Borders.Color = Get-OleColor 166 166 166

# =========================================================================
# SECTION 6: STEEL GRADE & METALLURGY (Cols T-W, Rows 7-26)
# =========================================================================
$dashboard.Range("T7:W7").Merge()
$dashboard.Cells.Item(7, 20) = "STEEL GRADE & SPECIFICATION DISTRIBUTION"
$dashboard.Range("T7:W7").Font.Bold = $true
$dashboard.Range("T7:W7").Font.Size = 10
$dashboard.Range("T7:W7").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("T7:W7").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("T7:W7").HorizontalAlignment = -4108

$dashboard.Cells.Item(8, 20) = "Grade"
$dashboard.Cells.Item(8, 21) = "Line Items"
$dashboard.Cells.Item(8, 22) = "Piece Qty"
$dashboard.Cells.Item(8, 23) = "Total Wt (lbs)"
$dashboard.Range("T8:W8").Font.Bold = $true
$dashboard.Range("T8:W8").Font.Size = 9
$dashboard.Range("T8:W8").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("T8:W8").Borders.LineStyle = 1
$dashboard.Range("T8:W8").Borders.Weight = 2
$dashboard.Range("T8:W8").Borders.Color = Get-OleColor 166 166 166

$topGrades = @($masterData | ForEach-Object { $_.Grade } | Where-Object { $_ } | Select-Object -Unique -First 10)
if ($topGrades.Count -eq 0) { $topGrades = @("A572-50") }
$grIdx = 9
foreach ($gr in $topGrades) {
    $dashboard.Cells.Item($grIdx, 20) = $gr
    $dashboard.Cells.Item($grIdx, 21).Formula = "=COUNTIF('Master BOM'!`$K`$2:`$K`$$lastMbRow, T$grIdx)"
    $dashboard.Cells.Item($grIdx, 22).Formula = "=SUMIF('Master BOM'!`$K`$2:`$K`$$lastMbRow, T$grIdx, 'Master BOM'!`$F`$2:`$F`$$lastMbRow)"
    $dashboard.Cells.Item($grIdx, 23).Formula = "=SUMIF('Master BOM'!`$K`$2:`$K`$$lastMbRow, T$grIdx, 'Master BOM'!`$N`$2:`$N`$$lastMbRow)"
    $dashboard.Cells.Item($grIdx, 21).NumberFormat = "#,##0"
    $dashboard.Cells.Item($grIdx, 22).NumberFormat = "#,##0"
    $dashboard.Cells.Item($grIdx, 23).NumberFormat = "#,##0.00"
    $dashboard.Range("T$grIdx:W$grIdx").Borders.LineStyle = 1
    $dashboard.Range("T$grIdx:W$grIdx").Borders.Weight = 2
    $dashboard.Range("T$grIdx:W$grIdx").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($grIdx, 20).Interior.Color = Get-OleColor 248 249 251
    $grIdx++
}

# Grade Total Row
$dashboard.Cells.Item($grIdx, 20) = "TOTAL (Top Grades)"
$dashboard.Cells.Item($grIdx, 21).Formula = "=SUM(U9:U$($grIdx - 1))"
$dashboard.Cells.Item($grIdx, 22).Formula = "=SUM(V9:V$($grIdx - 1))"
$dashboard.Cells.Item($grIdx, 23).Formula = "=SUM(W9:W$($grIdx - 1))"
$dashboard.Range("T$grIdx:W$grIdx").Font.Bold = $true
$dashboard.Range("T$grIdx:W$grIdx").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($grIdx, 21).NumberFormat = "#,##0"
$dashboard.Cells.Item($grIdx, 22).NumberFormat = "#,##0"
$dashboard.Cells.Item($grIdx, 23).NumberFormat = "#,##0.00"
$dashboard.Range("T$grIdx:W$grIdx").Borders.LineStyle = 1
$dashboard.Range("T$grIdx:W$grIdx").Borders.Weight = 2
$dashboard.Range("T$grIdx:W$grIdx").Borders.Color = Get-OleColor 166 166 166

# =========================================================================
# SECTION 7: FASTENER TYPE DISTRIBUTION (Cols E-H, Rows 28-36)
# =========================================================================
$dashboard.Range("E28:H28").Merge()
$dashboard.Cells.Item(28, 5) = "HARDWARE & FASTENER SHAPE / TYPE DISTRIBUTION"
$dashboard.Range("E28:H28").Font.Bold = $true
$dashboard.Range("E28:H28").Font.Size = 10
$dashboard.Range("E28:H28").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("E28:H28").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("E28:H28").HorizontalAlignment = -4108

$dashboard.Cells.Item(29, 5) = "Fastener Type"
$dashboard.Cells.Item(29, 6) = "Line Items"
$dashboard.Cells.Item(29, 7) = "Piece Quantity"
$dashboard.Cells.Item(29, 8) = "Total Wt (lbs)"
$dashboard.Range("E29:H29").Font.Bold = $true
$dashboard.Range("E29:H29").Font.Size = 9
$dashboard.Range("E29:H29").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("E29:H29").Borders.LineStyle = 1
$dashboard.Range("E29:H29").Borders.Weight = 2
$dashboard.Range("E29:H29").Borders.Color = Get-OleColor 166 166 166

$fastenerShapes = @("NUT", "WASHER", "BOLT", "ROD", "ANCHOR")
$fsRow = 30
foreach ($fShape in $fastenerShapes) {
    $dashboard.Cells.Item($fsRow, 5) = $fShape
    $dashboard.Cells.Item($fsRow, 6).Formula = "=COUNTIF('Hardware & Fasteners'!`$C`$2:`$C`$$lastHwRow, E$fsRow)"
    $dashboard.Cells.Item($fsRow, 7).Formula = "=SUMIF('Hardware & Fasteners'!`$C`$2:`$C`$$lastHwRow, E$fsRow, 'Hardware & Fasteners'!`$F`$2:`$F`$$lastHwRow)"
    $dashboard.Cells.Item($fsRow, 8).Formula = "=SUMIF('Hardware & Fasteners'!`$C`$2:`$C`$$lastHwRow, E$fsRow, 'Hardware & Fasteners'!`$J`$2:`$J`$$lastHwRow)"
    $dashboard.Cells.Item($fsRow, 6).NumberFormat = "#,##0"
    $dashboard.Cells.Item($fsRow, 7).NumberFormat = "#,##0"
    $dashboard.Cells.Item($fsRow, 8).NumberFormat = "#,##0.00"
    $dashboard.Range("E$fsRow:H$fsRow").Borders.LineStyle = 1
    $dashboard.Range("E$fsRow:H$fsRow").Borders.Weight = 2
    $dashboard.Range("E$fsRow:H$fsRow").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($fsRow, 5).Interior.Color = Get-OleColor 248 249 251
    $fsRow++
}

# Fasteners Total Row
$dashboard.Cells.Item($fsRow, 5) = "TOTAL FASTENERS"
$dashboard.Cells.Item($fsRow, 6).Formula = "=SUM(F30:F$($fsRow - 1))"
$dashboard.Cells.Item($fsRow, 7).Formula = "=SUM(G30:G$($fsRow - 1))"
$dashboard.Cells.Item($fsRow, 8).Formula = "=SUM(H30:H$($fsRow - 1))"
$dashboard.Range("E$fsRow:H$fsRow").Font.Bold = $true
$dashboard.Range("E$fsRow:H$fsRow").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($fsRow, 6).NumberFormat = "#,##0"
$dashboard.Cells.Item($fsRow, 7).NumberFormat = "#,##0"
$dashboard.Cells.Item($fsRow, 8).NumberFormat = "#,##0.00"
$dashboard.Range("E$fsRow:H$fsRow").Borders.LineStyle = 1
$dashboard.Range("E$fsRow:H$fsRow").Borders.Weight = 2
$dashboard.Range("E$fsRow:H$fsRow").Borders.Color = Get-OleColor 166 166 166

# =========================================================================
# SECTION 8: LABOR OPERATIONS BREAKDOWN (Cols J-M, Rows 28-36)
# =========================================================================
$dashboard.Range("J28:M28").Merge()
$dashboard.Cells.Item(28, 10) = "SHOP LABOR & WELD OPERATIONS BREAKDOWN"
$dashboard.Range("J28:M28").Font.Bold = $true
$dashboard.Range("J28:M28").Font.Size = 10
$dashboard.Range("J28:M28").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("J28:M28").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("J28:M28").HorizontalAlignment = -4108

$dashboard.Cells.Item(29, 10) = "Operation Type"
$dashboard.Cells.Item(29, 11) = "Operation Count"
$dashboard.Cells.Item(29, 12) = "Total Qty"
$dashboard.Cells.Item(29, 13) = "Avg Qty / Op"
$dashboard.Range("J29:M29").Font.Bold = $true
$dashboard.Range("J29:M29").Font.Size = 9
$dashboard.Range("J29:M29").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("J29:M29").Borders.LineStyle = 1
$dashboard.Range("J29:M29").Borders.Weight = 2
$dashboard.Range("J29:M29").Borders.Color = Get-OleColor 166 166 166

$laborTypes = @("HOLE", "WELD", "CUT", "COPE", "OTHER")
$loRow = 30
foreach ($lt in $laborTypes) {
    $dashboard.Cells.Item($loRow, 10) = $lt
    $dashboard.Cells.Item($loRow, 11).Formula = "=COUNTIF('Labor & Weld Operations'!`$D`$2:`$D`$$lastLbRow, J$loRow)"
    $dashboard.Cells.Item($loRow, 12).Formula = "=SUMIF('Labor & Weld Operations'!`$D`$2:`$D`$$lastLbRow, J$loRow, 'Labor & Weld Operations'!`$E`$2:`$E`$$lastLbRow)"
    $dashboard.Cells.Item($loRow, 13).Formula = "=IF(K$loRow>0, L$loRow/K$loRow, 0)"
    $dashboard.Cells.Item($loRow, 11).NumberFormat = "#,##0"
    $dashboard.Cells.Item($loRow, 12).NumberFormat = "#,##0"
    $dashboard.Cells.Item($loRow, 13).NumberFormat = "#,##0.0"
    $dashboard.Range("J$loRow:M$loRow").Borders.LineStyle = 1
    $dashboard.Range("J$loRow:M$loRow").Borders.Weight = 2
    $dashboard.Range("J$loRow:M$loRow").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($loRow, 10).Interior.Color = Get-OleColor 248 249 251
    $loRow++
}

# Labor Total Row
$dashboard.Cells.Item($loRow, 10) = "TOTAL OPERATIONS"
$dashboard.Cells.Item($loRow, 11).Formula = "=SUM(K30:K$($loRow - 1))"
$dashboard.Cells.Item($loRow, 12).Formula = "=SUM(L30:L$($loRow - 1))"
$dashboard.Cells.Item($loRow, 13).Formula = "=IF(K$loRow>0, L$loRow/K$loRow, 0)"
$dashboard.Range("J$loRow:M$loRow").Font.Bold = $true
$dashboard.Range("J$loRow:M$loRow").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($loRow, 11).NumberFormat = "#,##0"
$dashboard.Cells.Item($loRow, 12).NumberFormat = "#,##0"
$dashboard.Cells.Item($loRow, 13).NumberFormat = "#,##0.0"
$dashboard.Range("J$loRow:M$loRow").Borders.LineStyle = 1
$dashboard.Range("J$loRow:M$loRow").Borders.Weight = 2
$dashboard.Range("J$loRow:M$loRow").Borders.Color = Get-OleColor 166 166 166

# =========================================================================
# SECTION 9: ASSEMBLY DENSITY & SAMPLE TOTALS (Cols O-R, Rows 28-36)
# =========================================================================
$dashboard.Range("O28:R28").Merge()
$dashboard.Cells.Item(28, 15) = "ASSEMBLY DENSITY & SAMPLE TOTALS"
$dashboard.Range("O28:R28").Font.Bold = $true
$dashboard.Range("O28:R28").Font.Size = 10
$dashboard.Range("O28:R28").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("O28:R28").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("O28:R28").HorizontalAlignment = -4108

$dashboard.Cells.Item(29, 15) = "Assembly Sample"
$dashboard.Cells.Item(29, 16) = "Line Items"
$dashboard.Cells.Item(29, 17) = "Total Qty"
$dashboard.Cells.Item(29, 18) = "Total Wt (lbs)"
$dashboard.Range("O29:R29").Font.Bold = $true
$dashboard.Range("O29:R29").Font.Size = 9
$dashboard.Range("O29:R29").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("O29:R29").Borders.LineStyle = 1
$dashboard.Range("O29:R29").Borders.Weight = 2
$dashboard.Range("O29:R29").Borders.Color = Get-OleColor 166 166 166

$sampleAssms = @($masterData | ForEach-Object { $_.Assembly } | Where-Object { $_ } | Select-Object -Unique -First 5)
if ($sampleAssms.Count -eq 0) { $sampleAssms = @("11137101") }
$asRow = 30
foreach ($asm in $sampleAssms) {
    $dashboard.Cells.Item($asRow, 15) = $asm
    $dashboard.Cells.Item($asRow, 16).Formula = "=COUNTIF('Master BOM'!`$B`$2:`$B`$$lastMbRow, O$asRow)"
    $dashboard.Cells.Item($asRow, 17).Formula = "=SUMIF('Master BOM'!`$B`$2:`$B`$$lastMbRow, O$asRow, 'Master BOM'!`$F`$2:`$F`$$lastMbRow)"
    $dashboard.Cells.Item($asRow, 18).Formula = "=SUMIF('Master BOM'!`$B`$2:`$B`$$lastMbRow, O$asRow, 'Master BOM'!`$N`$2:`$N`$$lastMbRow)"
    $dashboard.Cells.Item($asRow, 16).NumberFormat = "#,##0"
    $dashboard.Cells.Item($asRow, 17).NumberFormat = "#,##0"
    $dashboard.Cells.Item($asRow, 18).NumberFormat = "#,##0.00"
    $dashboard.Range("O$asRow:R$asRow").Borders.LineStyle = 1
    $dashboard.Range("O$asRow:R$asRow").Borders.Weight = 2
    $dashboard.Range("O$asRow:R$asRow").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($asRow, 15).Interior.Color = Get-OleColor 248 249 251
    $asRow++
}

# Assembly Sample Total
$dashboard.Cells.Item($asRow, 15) = "SAMPLE TOTAL"
$dashboard.Cells.Item($asRow, 16).Formula = "=SUM(P30:P$($asRow - 1))"
$dashboard.Cells.Item($asRow, 17).Formula = "=SUM(Q30:Q$($asRow - 1))"
$dashboard.Cells.Item($asRow, 18).Formula = "=SUM(R30:R$($asRow - 1))"
$dashboard.Range("O$asRow:R$asRow").Font.Bold = $true
$dashboard.Range("O$asRow:R$asRow").Interior.Color = Get-OleColor 235 241 222
$dashboard.Cells.Item($asRow, 16).NumberFormat = "#,##0"
$dashboard.Cells.Item($asRow, 17).NumberFormat = "#,##0"
$dashboard.Cells.Item($asRow, 18).NumberFormat = "#,##0.00"
$dashboard.Range("O$asRow:R$asRow").Borders.LineStyle = 1
$dashboard.Range("O$asRow:R$asRow").Borders.Weight = 2
$dashboard.Range("O$asRow:R$asRow").Borders.Color = Get-OleColor 166 166 166

# =========================================================================
# SECTION 10: ESTIMATING & PRODUCTION RATIOS (Cols T-W, Rows 28-36)
# =========================================================================
$dashboard.Range("T28:W28").Merge()
$dashboard.Cells.Item(28, 20) = "ESTIMATING & PRODUCTION RATIOS"
$dashboard.Range("T28:W28").Font.Bold = $true
$dashboard.Range("T28:W28").Font.Size = 10
$dashboard.Range("T28:W28").Font.Color = Get-OleColor 255 255 255
$dashboard.Range("T28:W28").Interior.Color = Get-OleColor 31 78 120
$dashboard.Range("T28:W28").HorizontalAlignment = -4108

$dashboard.Cells.Item(29, 20) = "KPI / Ratio Metric"
$dashboard.Range("T29:U29").Merge()
$dashboard.Cells.Item(29, 22) = "Formula Value"
$dashboard.Cells.Item(29, 23) = "Units"
$dashboard.Range("T29:W29").Font.Bold = $true
$dashboard.Range("T29:W29").Font.Size = 9
$dashboard.Range("T29:W29").Interior.Color = Get-OleColor 216 228 188
$dashboard.Range("T29:W29").Borders.LineStyle = 1
$dashboard.Range("T29:W29").Borders.Weight = 2
$dashboard.Range("T29:W29").Borders.Color = Get-OleColor 166 166 166

$ratios = @(
    @{ Name = "Overall Avg Wt per Piece"; Fml = "=IF(B10>0, B11/B10, 0)"; Unit = "lbs / pc"; Fmt = "#,##0.00" },
    @{ Name = "Plate Weight Percentage"; Fml = "=IF(B11>0, H18/B11, 0)"; Unit = "% of total"; Fmt = "0.0%" },
    @{ Name = "Profile Weight Percentage"; Fml = "=IF(B11>0, H19/B11, 0)"; Unit = "% of total"; Fmt = "0.0%" },
    @{ Name = "Hardware Weight Percentage"; Fml = "=IF(B11>0, H20/B11, 0)"; Unit = "% of total"; Fmt = "0.0%" },
    @{ Name = "Shop Operations per Assembly"; Fml = "=IF(B13>0, L35/B13, 0)"; Unit = "ops / assm"; Fmt = "#,##0.0" },
    @{ Name = "Shop Operations per Ton"; Fml = "=IF(B12>0, L35/B12, 0)"; Unit = "ops / ton"; Fmt = "#,##0.0" }
)

$rtRow = 30
foreach ($rt in $ratios) {
    $dashboard.Cells.Item($rtRow, 20) = $rt.Name
    $dashboard.Range("T$rtRow:U$rtRow").Merge()
    $dashboard.Cells.Item($rtRow, 22).Formula = $rt.Fml
    $dashboard.Cells.Item($rtRow, 23) = $rt.Unit
    $dashboard.Cells.Item($rtRow, 22).NumberFormat = $rt.Fmt
    $dashboard.Range("T$rtRow:W$rtRow").Borders.LineStyle = 1
    $dashboard.Range("T$rtRow:W$rtRow").Borders.Weight = 2
    $dashboard.Range("T$rtRow:W$rtRow").Borders.Color = Get-OleColor 220 225 230
    $dashboard.Cells.Item($rtRow, 20).Interior.Color = Get-OleColor 248 249 251
    $rtRow++
}

# --- ALIGNMENT & STYLING ---
$dashboard.Range("A8:W37").HorizontalAlignment = -4108
$dashboard.Range("A8:W37").VerticalAlignment = -4108

# Left-align text / description columns
$dashboard.Range("A9:A26").HorizontalAlignment = -4131
$dashboard.Range("E9:E15").HorizontalAlignment = -4131
$dashboard.Range("E18:E23").HorizontalAlignment = -4131
$dashboard.Range("E30:E35").HorizontalAlignment = -4131
$dashboard.Range("J9:J26").HorizontalAlignment = -4131
$dashboard.Range("J30:J35").HorizontalAlignment = -4131
$dashboard.Range("O9:O26").HorizontalAlignment = -4131
$dashboard.Range("O30:O35").HorizontalAlignment = -4131
$dashboard.Range("T9:T26").HorizontalAlignment = -4131
$dashboard.Range("T30:U35").HorizontalAlignment = -4131

# Column widths
$dashboard.Columns.Item(1).ColumnWidth = 30
$dashboard.Columns.Item(2).ColumnWidth = 16
$dashboard.Columns.Item(3).ColumnWidth = 22
$dashboard.Columns.Item(4).ColumnWidth = 4
$dashboard.Columns.Item(5).ColumnWidth = 24
$dashboard.Columns.Item(6).ColumnWidth = 12
$dashboard.Columns.Item(7).ColumnWidth = 12
$dashboard.Columns.Item(8).ColumnWidth = 16
$dashboard.Columns.Item(9).ColumnWidth = 4
$dashboard.Columns.Item(10).ColumnWidth = 14
$dashboard.Columns.Item(11).ColumnWidth = 12
$dashboard.Columns.Item(12).ColumnWidth = 12
$dashboard.Columns.Item(13).ColumnWidth = 16
$dashboard.Columns.Item(14).ColumnWidth = 4
$dashboard.Columns.Item(15).ColumnWidth = 26
$dashboard.Columns.Item(16).ColumnWidth = 12
$dashboard.Columns.Item(17).ColumnWidth = 12
$dashboard.Columns.Item(18).ColumnWidth = 16
$dashboard.Columns.Item(19).ColumnWidth = 4
$dashboard.Columns.Item(20).ColumnWidth = 16
$dashboard.Columns.Item(21).ColumnWidth = 12
$dashboard.Columns.Item(22).ColumnWidth = 12
$dashboard.Columns.Item(23).ColumnWidth = 16

# --- TAB: FOLDER LABELS (only when a Required Assm workbook was chosen) ---
if ($requiredAssmPath) {
    Write-Log "Building Folder Labels tab from $([System.IO.Path]::GetFileName($requiredAssmPath))..."
    try {
        $labelData = @(Read-RequiredAssm -Excel $excel -Path $requiredAssmPath)
        $allWorkbookTLs = $labelData
        if ($allTLs.Count -gt 0) {
            $runNumbers = @($allTLs | ForEach-Object { [int]($_ -replace '\D', '0') })
            $skipped = @($labelData | Where-Object { $runNumbers -notcontains $_.Number } | ForEach-Object { $_.TL })
            if ($skipped.Count -gt 0) { Write-Log "No label for $($skipped -join ', '): not in the KSS files selected for this run." "Yellow" }
            $labelData = @($labelData | Where-Object { $runNumbers -contains $_.Number })
            $missingLabels = @($runNumbers | Where-Object { $n = $_; -not (@($labelData | Where-Object { $_.Number -eq $n })) })
            if ($missingLabels.Count -gt 0) { Write-Log "No Required Assm entry for TL$($missingLabels -join ', TL'); no label made for it." "Yellow" }
        }
        if ($labelData.Count -eq 0) {
            Write-Log "No transmittals with assemblies found in the Required Assm workbook; skipping the Folder Labels tab." "Yellow"
        }
        else {
            # One label per folder = one per (TL, sequence), ordered by TL number then sequence
            $labelData = @($labelData | Sort-Object Number | ForEach-Object {
                $tl = $_
                foreach ($seq in $tl.Sequences) {
                    $note = $tl.SeqNotes[$seq]
                    $due = if ($note -and $note.DueDate) { $note.DueDate } else { $tl.DueDate }
                    $coat = if ($note -and $note.Coating) { $note.Coating } else { $tl.Coating }
                    $notes = $coat
                    if ($tl.IsRev) {
                        # REV: name the transmittals that hold the same assembly numbers in this sequence (non-REV ones preferred)
                        $mine = $tl.SeqAssemblies[$seq]
                        $hits = @($allWorkbookTLs | Where-Object {
                            $_.TL -ne $tl.TL -and $_.SeqAssemblies.ContainsKey($seq) -and $mine.Overlaps($_.SeqAssemblies[$seq])
                        } | Sort-Object Number)
                        $plain = @($hits | Where-Object { -not $_.IsRev })
                        if ($plain.Count -gt 0) { $hits = $plain }
                        if ($hits.Count -gt 0) { $notes = (@("REVISES $(($hits | ForEach-Object { $_.TL }) -join ', ')", $coat) | Where-Object { $_ }) -join " - " }
                    }
                    [PSCustomObject]@{ TL = $tl.TL; IsRev = $tl.IsRev; Sequence = $seq; DueDate = $due; Notes = $notes }
                }
            })
            [void](Add-FolderLabelSheet -Workbook $workbook -Labels $labelData -OurJob $ourJobNumber -Title $projectName.ToUpper() -CustomerJob $jobNumber -IssuedBy $issuedBy)
            foreach ($l in $labelData) {
                $line1 = (@($ourJobNumber, $l.TL, $l.Sequence) | Where-Object { $_ }) -join " - "
                Write-Log "  Label: $line1 | $($l.Notes) | $($l.DueDate)" "Cyan"
            }
            Write-Log "Folder Labels tab ready: $($labelData.Count) label(s). Print that tab on the Dymo (3.5 x 1.125 in)." "Green"
        }
    }
    catch {
        Write-Log "Could not build the Folder Labels tab: $($_.Exception.Message)" "Red"
    }
}

# Sheet presentation setup
$dashboard.Move($workbook.Sheets.Item(1))
$dashboard.Activate()
$excel.ActiveWindow.DisplayGridlines = $true
$dashboard.Range("A7").Select()
$excel.ActiveWindow.FreezePanes = $true

# 4. SAVE FILE AND CLEANUP
Write-Log "Finalizing and saving workbook..."
$safeProjectName = $projectName -replace '[\\/:*?"<>|]', '_'
$tlNamePart = ""
if ($allTLs.Count -gt 0) {
    $tlNamePart = " - " + ($allTLs -join "_")
    if (((Get-Location).Path.Length + $safeProjectName.Length + $jobNumber.Length + $tlNamePart.Length) -gt 220) {
        $tlNamePart = " - $($allTLs[0])-$($allTLs[$allTLs.Count - 1]) ($($allTLs.Count) TLs)"
    }
}
$exportPath = Join-Path (Get-Location) "$safeProjectName - Job $jobNumber$tlNamePart BOM.xlsx"

if (Test-Path $exportPath) {
    try {
        Remove-Item $exportPath -Force -ErrorAction Stop
    }
    catch {
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $exportPath = Join-Path (Get-Location) "$safeProjectName - Job $jobNumber$tlNamePart BOM $stamp.xlsx"
        Write-Log "Existing workbook locked; saving to a unique file instead." "Yellow"
        Write-Log " -> $exportPath" "Yellow"
    }
}

$sheet1.Activate()
$workbook.SaveAs($exportPath)
$workbook.Close($false)
$excel.Quit()

Write-Log "SUCCESS! Formatted workbook saved directly to:" "Green"
Write-Log " -> $exportPath" "Yellow"

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
[gc]::collect()
[gc]::WaitForPendingFinalizers()


# 5. GENERATE PER-FILE SHOP SCRIPTS (DXF SORT, QUANTITIES, GRADES) + MASTER LAUNCHER
Write-Log "Generating shop scripts and master launcher..." "Yellow"
$scriptSets = [System.Collections.Generic.List[PSCustomObject]]::new()
$usedLabels = @{}

if ($combinedShopScripts -and $masterData.Count -gt 0) {
    $combinedLabel = if ($allTLs.Count -gt 0) { ($allTLs -join '_') } else { "ALL-KSS" }
    $combinedLabel = $combinedLabel -replace '[\\/:*?"<>|]', '_'
    $combinedScriptSet = Write-ShopScripts -Items $masterData -ProjectName $projectName -JobNumber $jobNumber -TL $combinedLabel -Label $combinedLabel
    if ($combinedScriptSet) { $scriptSets.Add($combinedScriptSet) }
    else { Write-Log "No plates in the combined output; skipping combined shop scripts." "Yellow" }
}
else {
    foreach ($fileGroup in ($masterData | Group-Object SourceFile)) {
        $fileBase = [System.IO.Path]::GetFileNameWithoutExtension($fileGroup.Name)
        $fileTL = $fileGroup.Group[0].TL
        $label = if ($fileTL) { $fileTL } else { $fileBase }
        $label = $label -replace '[\\/:*?"<>|]', '_'
        if ($usedLabels.ContainsKey($label)) { $label = "$label ($fileBase)" -replace '[\\/:*?"<>|]', '_' }
        $usedLabels[$label] = $true

        $scriptSet = Write-ShopScripts -Items $fileGroup.Group -ProjectName $projectName -JobNumber $jobNumber -TL $fileTL -Label $label
        if ($scriptSet) { $scriptSets.Add($scriptSet) }
        else { Write-Log "No plates in $($fileGroup.Name); skipping its shop scripts." "Yellow" }
    }
}

if ($scriptSets.Count -gt 0) {
    Write-Log "Generated shop scripts without a master launcher." "Yellow"
}





function Write-QualityProcessReport {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$MasterData,
        [string]$ProjectName,
        [string]$JobNumber,
        [string]$QprJobNumber,  # Job number for sheet naming
        [string]$GroupBy  # "TL", "Assembly", or "Sequence"
    )

    $xlContinuous    = 1
    $xlThin          = 2
    $xlMedium        = -4138
    $xlNoneColorIdx  = -4142
    $xlCenter        = -4108
    $xlBottom        = -4107
    $xlLeft          = -4131
    $xlHoriz         = -4128
    $grayBg          = 14277081

    # Helper: set a single border side by literal index (avoids variable-cast COM error)
    # Uses the Range's BorderAround trick: set all borders then we handle sides separately via
    # individual cell references. For targeted single-side ops we call this small helper.
    function Set-BorderSide ($cellOrRange, [int]$side, [int]$style, [int]$weight) {
        # $side: 7=left 8=top 9=bottom 10=right
        try {
            $bd = $cellOrRange.Borders.Item($side)
            $bd.LineStyle = $style
            $bd.Weight    = $weight
        } catch {
            # fallback: skip per-side tweak if COM refuses
        }
    }

    if ($GroupBy -eq "TL") {
        $groups = @($MasterData | Group-Object TL | Sort-Object Name)
    } elseif ($GroupBy -eq "Sequence") {
        # One sheet per TL + sequence (first 4 digits of assembly), so the sequence sheets still say which TL they came from
        $groups = @($MasterData | Group-Object {
            $assm = $_.Assembly
            $seq = if ($assm -match '^(\d{4})') { $Matches[1] } else { $assm.Substring(0, [Math]::Min(4, $assm.Length)) }
            "$($_.TL)`t$seq"
        } | Sort-Object Name)
    } else {
        $groups = @($MasterData | Group-Object Assembly | Sort-Object Name)
    }

    if ($groups.Count -eq 0) {
        Write-Log "No data found for Quality Process Report." "Yellow"
        return $null
    }

    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $false
    $xlApp.DisplayAlerts = $false
    $xlApp.ScreenUpdating = $false

    try {
        $wb = $xlApp.Workbooks.Add()
        while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

        $headerDefs = @(
            [PSCustomObject]@{ Col=1;  Label="QTY";               Orient=$xlHoriz; Gray=$false; Va=$xlCenter; ThickL=$true  }
            [PSCustomObject]@{ Col=2;  Label="PART #";             Orient=$xlHoriz; Gray=$false; Va=$xlCenter; ThickL=$true  }
            [PSCustomObject]@{ Col=3;  Label="ITEM DESCRIPTION";   Orient=$xlHoriz; Gray=$false; Va=$xlCenter; ThickL=$false }
            [PSCustomObject]@{ Col=4;  Label="MATERIAL SIZE";      Orient=45;       Gray=$true;  Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=5;  Label="CUT LENGTH";         Orient=45;       Gray=$false; Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=6;  Label="HOLE SIZE";          Orient=45;       Gray=$true;  Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=7;  Label="LAYOUT";             Orient=45;       Gray=$false; Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=8;  Label="FIT-UP";             Orient=45;       Gray=$true;  Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=9;  Label="WELD SIZE";          Orient=45;       Gray=$false; Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=10; Label="WELDER I.D. STAMP";  Orient=45;       Gray=$true;  Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=11; Label="PART NUMBER STAMP";  Orient=45;       Gray=$false; Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=12; Label="SURFACE PREP.";      Orient=45;       Gray=$true;  Va=$xlBottom; ThickL=$false }
            [PSCustomObject]@{ Col=13; Label="COATING";            Orient=45;       Gray=$false; Va=$xlBottom; ThickL=$false }
        )

        $sheetIndex = 0
        foreach ($grp in $groups) {
            $sheetIndex++
            # Sheet / job label: {job} - {TL}, or {job} - {TL} - {sequence} when grouped by sequence
            $labelParts = @()
            if ($QprJobNumber) { $labelParts += $QprJobNumber }
            if ($GroupBy -eq "Sequence") {
                $seqTl, $seqNum = $grp.Name -split "`t", 2
                if ($seqTl) { $labelParts += $seqTl }
                $labelParts += $seqNum
            }
            else {
                $labelParts += $(if ($grp.Name) { $grp.Name } else { "(No TL)" })
            }
            $jobLabel  = $labelParts -join " - "
            $sheetName = ($jobLabel -replace '[\\\/\?\*\[\]\:]', '_')
            if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }

            $ws = if ($sheetIndex -eq 1) {
                $wb.Sheets.Item(1)
            } else {
                $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
            }
            $ws.Name = $sheetName

            # Column widths
            $ws.Columns.Item(1).ColumnWidth  = 6
            $ws.Columns.Item(2).ColumnWidth  = 10
            $ws.Columns.Item(3).ColumnWidth  = 30
            for ($c = 4; $c -le 16; $c++) { $ws.Columns.Item($c).ColumnWidth = 6 }
            $ws.Columns.Item(17).ColumnWidth = 3.43

            $ps = $ws.PageSetup
            $ps.Orientation   = 2
            $ps.Zoom          = $false
            $ps.FitToPagesWide = 1
            $ps.FitToPagesTall = $false
            $ps.LeftMargin    = 7.2
            $ps.RightMargin   = 7.2
            $ps.TopMargin     = 7.2
            $ps.BottomMargin  = 7.2
            $ps.HeaderMargin  = 0
            $ps.FooterMargin  = 0
            $ps.CenterHeader  = ""
            $ps.LeftHeader    = ""
            $ps.RightHeader   = ""
            $ps.CenterFooter  = ""
            $ps.LeftFooter    = ""
            $ps.RightFooter   = ""

            # Row heights
            $ws.Rows.Item(1).RowHeight = 16.5
            $ws.Rows.Item(2).RowHeight = 16.5
            $ws.Rows.Item(3).RowHeight = 16.5
            $ws.Rows.Item(4).RowHeight = 21.75
            $ws.Rows.Item(5).RowHeight = 74.25

            # --- Title R1C3:R3C11 ---
            $ws.Range($ws.Cells.Item(1,3), $ws.Cells.Item(3,11)).Merge()
            $ws.Cells.Item(1,3).Value2              = "          PROJECT QUALITY PROCESS REPORT"
            $ws.Cells.Item(1,3).Font.Name           = "Arial Narrow"
            $ws.Cells.Item(1,3).Font.Size           = 16
            $ws.Cells.Item(1,3).Font.Bold           = $true
            $ws.Cells.Item(1,3).HorizontalAlignment = $xlLeft
            $ws.Cells.Item(1,3).VerticalAlignment   = $xlCenter

            # Job Number label R1C12:C13, value R1C14:C16 with bottom underline
            $ws.Range($ws.Cells.Item(1,12), $ws.Cells.Item(1,13)).Merge()
            $ws.Cells.Item(1,12).Value2    = "Job:"
            $ws.Cells.Item(1,12).Font.Name = "Arial Narrow"
            $ws.Cells.Item(1,12).Font.Bold = $true
            $jnRange = $ws.Range($ws.Cells.Item(1,14), $ws.Cells.Item(1,16))
            $jnRange.Merge()
            $ws.Cells.Item(1,14).Value2    = $jobLabel   # mirrors the sheet name (Excel cuts sheet names at 31 characters, this cell does not)
            $ws.Cells.Item(1,14).Font.Name = "Arial Narrow"
            $jnRange.ShrinkToFit = $true
            Set-BorderSide $jnRange 9 $xlContinuous $xlThin   # bottom underline

            # Inspector R2C12:C13, write-in area R2C14:C16 top+bottom lines
            $ws.Range($ws.Cells.Item(2,12), $ws.Cells.Item(2,13)).Merge()
            $ws.Cells.Item(2,12).Value2    = "Inspector:"
            $ws.Cells.Item(2,12).Font.Name = "Arial Narrow"
            $ws.Cells.Item(2,12).Font.Bold = $true
            $insRange = $ws.Range($ws.Cells.Item(2,14), $ws.Cells.Item(2,16))
            $insRange.Merge()
            Set-BorderSide $insRange 8 $xlContinuous $xlThin   # top
            Set-BorderSide $insRange 9 $xlContinuous $xlThin   # bottom

            # Date R3C12:C13, write-in area R3C14:C16 top+bottom lines
            $ws.Range($ws.Cells.Item(3,12), $ws.Cells.Item(3,13)).Merge()
            $ws.Cells.Item(3,12).Value2    = "Date:"
            $ws.Cells.Item(3,12).Font.Name = "Arial Narrow"
            $ws.Cells.Item(3,12).Font.Bold = $true
            $dtRange = $ws.Range($ws.Cells.Item(3,14), $ws.Cells.Item(3,16))
            $dtRange.Merge()
            Set-BorderSide $dtRange 8 $xlContinuous $xlThin    # top
            Set-BorderSide $dtRange 9 $xlContinuous $xlThin    # bottom

            # --- Row 4 separator: medium bottom border across C3:C12 and C13:C16 ---
            $sep1 = $ws.Range($ws.Cells.Item(4,3), $ws.Cells.Item(4,12))
            $sep1.Merge()
            Set-BorderSide $sep1 9 $xlContinuous $xlMedium
            $sep2 = $ws.Range($ws.Cells.Item(4,13), $ws.Cells.Item(4,16))
            $sep2.Merge()
            Set-BorderSide $sep2 9 $xlContinuous $xlMedium

            # --- Header row 5 ---
            foreach ($hd in $headerDefs) {
                $cell = $ws.Cells.Item(5, $hd.Col)
                $cell.Value2              = $hd.Label
                $cell.Font.Name          = "Arial Narrow"
                $cell.Font.Size          = 10
                $cell.Font.Bold          = $true
                $cell.Orientation        = $hd.Orient
                $cell.VerticalAlignment  = $hd.Va
                $cell.HorizontalAlignment = $xlCenter
                $cell.WrapText           = $false
                if ($hd.Gray) {
                    $cell.Interior.Color = $grayBg
                } else {
                    $cell.Interior.ColorIndex = $xlNoneColorIdx
                }
                $cell.Borders.LineStyle = $xlContinuous
                $cell.Borders.Weight    = $xlThin
                if ($hd.ThickL) { Set-BorderSide $cell 7 $xlContinuous $xlMedium }
            }

            # --- Data rows: main-mark assemblies only ---
            $mainMarks = @($grp.Group |
                Where-Object { $_.PartMark -eq $_.Assembly -or $_.RawPartMark -eq $_.Assembly } |
                Group-Object Assembly |
                ForEach-Object { $_.Group[0] } |
                Sort-Object Assembly)

            $dataRow = 6
            foreach ($part in $mainMarks) {
                $ws.Rows.Item($dataRow).RowHeight = 16

                # C1: QTY â€” medium left border
                $ws.Cells.Item($dataRow, 1) = [int]$part.Qty
                $c1 = $ws.Cells.Item($dataRow, 1)
                $c1.Font.Name            = "Arial Narrow"
                $c1.Font.Size            = 11
                $c1.HorizontalAlignment  = $xlCenter
                $c1.Interior.ColorIndex  = $xlNoneColorIdx
                $c1.Borders.LineStyle    = $xlContinuous
                $c1.Borders.Weight       = $xlThin
                Set-BorderSide $c1 7 $xlContinuous $xlMedium

                # C2: PART NUMBER â€” medium left border
                $ws.Cells.Item($dataRow, 2) = [string]$part.PartMark
                $c2 = $ws.Cells.Item($dataRow, 2)
                $c2.Font.Name            = "Arial Narrow"
                $c2.Font.Size            = 11
                $c2.HorizontalAlignment  = $xlCenter
                $c2.Interior.ColorIndex  = $xlNoneColorIdx
                $c2.Borders.LineStyle    = $xlContinuous
                $c2.Borders.Weight       = $xlThin
                Set-BorderSide $c2 7 $xlContinuous $xlMedium

                # C3: ITEM DESCRIPTION
                $ws.Cells.Item($dataRow, 3) = [string]$part.Description
                $c3 = $ws.Cells.Item($dataRow, 3)
                $c3.Font.Name            = "Arial Narrow"
                $c3.Font.Size            = 11
                $c3.HorizontalAlignment  = $xlLeft
                $c3.Interior.ColorIndex  = $xlNoneColorIdx
                $c3.Borders.LineStyle    = $xlContinuous
                $c3.Borders.Weight       = $xlThin

                for ($cc = 4; $cc -le 13; $cc++) {
                    $cell = $ws.Cells.Item($dataRow, $cc)
                    if ($headerDefs[$cc - 1].Gray) { $cell.Interior.Color = $grayBg } else { $cell.Interior.ColorIndex = $xlNoneColorIdx }
                    $cell.Borders.LineStyle = $xlContinuous
                    $cell.Borders.Weight    = $xlThin
                }

                # C14-C16: merged notes column, thin borders all around
                $noteRange = $ws.Range($ws.Cells.Item($dataRow, 14), $ws.Cells.Item($dataRow, 16))
                $noteRange.Merge()
                $noteRange.Borders.LineStyle = $xlContinuous
                $noteRange.Borders.Weight    = $xlThin

                $dataRow++
            }

            # --- Bottom section (after last data row) ---
            # Blank separator
            $ws.Rows.Item($dataRow).RowHeight = 16
            Set-BorderSide ($ws.Cells.Item($dataRow, 1)) 8 $xlContinuous $xlThin
            $dataRow++

            # Instructions text: **INITIALS ARE TO BE PLACED...** merged C2:C16
            $ws.Rows.Item($dataRow).RowHeight = 16
            $instrRange = $ws.Range($ws.Cells.Item($dataRow, 2), $ws.Cells.Item($dataRow, 16))
            $instrRange.Merge()
            $ws.Cells.Item($dataRow, 2).Value2    = "**INITIALS ARE TO BE PLACED IN EACH BOX TO IDENTIFY THE CORRESPONDING ACTIVITY**"
            $ws.Cells.Item($dataRow, 2).Font.Name = "Arial Narrow"
            $ws.Cells.Item($dataRow, 2).Font.Size = 9
            $ws.Cells.Item($dataRow, 2).Font.Bold = $true
            $dataRow += 3  # skip 2 blank rows

            # Inspector signature row
            $ws.Rows.Item($dataRow).RowHeight = 16
            $ws.Range($ws.Cells.Item($dataRow, 1), $ws.Cells.Item($dataRow, 2)).Merge()
            $ws.Cells.Item($dataRow, 1).Value2    = "Inspector:"
            $ws.Cells.Item($dataRow, 1).Font.Name = "Arial Narrow"
            $ws.Cells.Item($dataRow, 1).Font.Bold = $true
            $sigR1 = $ws.Range($ws.Cells.Item($dataRow, 3), $ws.Cells.Item($dataRow, 5))
            $sigR1.Merge()
            Set-BorderSide $sigR1 9 $xlContinuous $xlThin
            $ws.Cells.Item($dataRow, 8).Value2    = "Date:"
            $ws.Cells.Item($dataRow, 8).Font.Name = "Arial Narrow"
            $ws.Cells.Item($dataRow, 8).Font.Bold = $true
            $dtR1 = $ws.Range($ws.Cells.Item($dataRow, 9), $ws.Cells.Item($dataRow, 10))
            $dtR1.Merge()
            Set-BorderSide $dtR1 9 $xlContinuous $xlThin
            $dataRow += 3  # skip 2 blank rows

            # Production Manager signature row
            $ws.Rows.Item($dataRow).RowHeight = 16
            $ws.Range($ws.Cells.Item($dataRow, 1), $ws.Cells.Item($dataRow, 2)).Merge()
            $ws.Cells.Item($dataRow, 1).Value2    = "Production Manager:"
            $ws.Cells.Item($dataRow, 1).Font.Name = "Arial Narrow"
            $ws.Cells.Item($dataRow, 1).Font.Bold = $true
            $sigR2 = $ws.Range($ws.Cells.Item($dataRow, 3), $ws.Cells.Item($dataRow, 5))
            $sigR2.Merge()
            Set-BorderSide $sigR2 9 $xlContinuous $xlThin
            $ws.Cells.Item($dataRow, 8).Value2    = "Date:"
            $ws.Cells.Item($dataRow, 8).Font.Name = "Arial Narrow"
            $ws.Cells.Item($dataRow, 8).Font.Bold = $true
            $dtR2 = $ws.Range($ws.Cells.Item($dataRow, 9), $ws.Cells.Item($dataRow, 10))
            $dtR2.Merge()
            Set-BorderSide $dtR2 9 $xlContinuous $xlThin
        }

        # Save
        $safeProjectName = $ProjectName -replace '[\\/:*?"<>|]', '_'
        $stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
        $qprPath = Join-Path (Get-Location) "$safeProjectName - Job $JobNumber - QPR $stamp.xlsx"
        if (Test-Path $qprPath) { Remove-Item $qprPath -Force -ErrorAction SilentlyContinue }
        $wb.Sheets.Item(1).Activate()
        $wb.SaveAs($qprPath)
        $wb.Close($false)

        Write-Log "Quality Process Report saved to:" "Green"
        Write-Log " -> $qprPath" "Yellow"
        return $qprPath
    }
    finally {
        $xlApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xlApp) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}


# =====================================================================
# 6. OPTIONAL: QUALITY PROCESS REPORT
# =====================================================================

Write-Host ""
Write-Host "========================================================"
Write-Host "  QUALITY PROCESS REPORT (QPR) OPTIONS"
Write-Host "========================================================"
Write-Host "  [Y]es: Generate detailed shop floor quality/inspection sheets"
Write-Host "  [N]o: Skip QPR generation (default)"
Write-Host ""
$qprChoice = Read-Host "Generate Quality Process Report (QPR)? [Y]es / [N]o (default: No)"
if ($qprChoice -match '^(y|yes)$') {
    Write-Host ""
    Write-Host "  Job name for the QPR (used as the sheet prefix and printed in the Job: box on each sheet)"
    Write-Host "  Example: enter '25-351' -> sheets named '25-351 - TL1282', '25-351 - TL1296', etc."
    $qprJobInput = Read-Host "Job name (default: $jobNumber)"
    $qprJobNumber = if ($qprJobInput) { $qprJobInput } else { $jobNumber }

    Write-Host ""
    Write-Host "  QPR Sheet Organization:"
    Write-Host "  [T]L (default): Organize by Transmittal number"
    Write-Host "    Sheets: $qprJobNumber - TL1282 | $qprJobNumber - TL1296 | $qprJobNumber - TL1301"
    Write-Host "    Use for: Manufacturing by transmittal/revision phase"
    Write-Host ""
    Write-Host "  [S]equence: Organize by Sequence (first 4 digits of assembly), still labeled with the TL"
    Write-Host "    Sheets: $qprJobNumber - TL1296 - 5131 | $qprJobNumber - TL1296 - 5132 | $qprJobNumber - TL1313 - 5133"
    Write-Host "    Use for: Manufacturing by structural family/sequence"
    Write-Host ""
    $qprGroupChoice = Read-Host "Organize sheets by [T]L number (default) or [S]equence?"
    $qprGroupBy = if ($qprGroupChoice -match '^(s|sequence)$') { "Sequence" } else { "TL" }

    Write-QualityProcessReport -MasterData $masterData -ProjectName $projectName -JobNumber $jobNumber -QprJobNumber $qprJobNumber -GroupBy $qprGroupBy
}

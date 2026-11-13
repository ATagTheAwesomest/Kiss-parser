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

# Any unexpected error ends up here: one plain message instead of a red wall of text, and a hidden Excel is not left running.
trap {
    Microsoft.PowerShell.Utility\Write-Host ""
    Microsoft.PowerShell.Utility\Write-Host "  Something went wrong and the run stopped." -ForegroundColor Red
    Microsoft.PowerShell.Utility\Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Microsoft.PowerShell.Utility\Write-Host "  (script line $($_.InvocationInfo.ScriptLineNumber))" -ForegroundColor DarkGray
    Microsoft.PowerShell.Utility\Write-Host "  Your .kss files were not changed. Take a screenshot of this window and send it to whoever maintains KISS." -ForegroundColor Yellow
    foreach ($xl in @($excel, $xlApp)) { if ($xl) { try { $xl.DisplayAlerts = $false; $xl.Quit() } catch { } } }
    Microsoft.PowerShell.Utility\Write-Host ""
    Microsoft.PowerShell.Utility\Read-Host "  Press Enter to close" | Out-Null
    exit 1
}

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

# --- SCREEN HELPERS ---
$script:OutputFiles = [System.Collections.Generic.List[string]]::new()
function Add-OutputFile ([string]$Path) {
    if ($Path -and -not $script:OutputFiles.Contains($Path)) { $script:OutputFiles.Add($Path) }
}

function Write-Section ([string]$Title) {
    Write-Host ""
    Write-Host ("=" * 56)
    Write-Host "  $Title"
    Write-Host ("=" * 56)
}

# --- MAIN MENU (centered, bold / italic / colors; falls back to plain colors where the console has no ANSI support) ---
$script:UseAnsi = $false
function Enable-AnsiConsole {
    try {
        if ($Host.Name -ne 'ConsoleHost') { return $false }
        if ($env:WT_SESSION) { return $true }
        if (-not ('Kiss.Con' -as [type])) {
            Add-Type -Namespace Kiss -Name Con -MemberDefinition '[DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int n); [DllImport("kernel32.dll")] public static extern bool GetConsoleMode(IntPtr h, out int m); [DllImport("kernel32.dll")] public static extern bool SetConsoleMode(IntPtr h, int m);'
        }
        $h = [Kiss.Con]::GetStdHandle(-11); $mode = 0
        if (-not [Kiss.Con]::GetConsoleMode($h, [ref]$mode)) { return $false }
        return [Kiss.Con]::SetConsoleMode($h, ($mode -bor 4))
    }
    catch { return $false }
}

# One styled piece of a line: @(text, color, 'b' bold / 'i' italic / 'bi' both)
function Write-Styled ($Segments) {
    $ansi = @{ Black = 30; DarkBlue = 34; DarkGreen = 32; DarkCyan = 36; DarkRed = 31; DarkMagenta = 35; DarkYellow = 33; Gray = 37; DarkGray = 90; Blue = 94; Green = 92; Cyan = 96; Red = 91; Magenta = 95; Yellow = 93; White = 97 }
    $esc = [string][char]27
    foreach ($seg in $Segments) {
        $text = [string]$seg[0]; $color = if ($seg.Count -gt 1 -and $seg[1]) { [string]$seg[1] } else { 'Gray' }
        $style = if ($seg.Count -gt 2) { [string]$seg[2] } else { '' }
        if ($script:UseAnsi) {
            $codes = @($ansi[$color]); if ($style -match 'b') { $codes += 1 }; if ($style -match 'i') { $codes += 3 }
            Microsoft.PowerShell.Utility\Write-Host ("$esc[" + ($codes -join ';') + "m$text$esc[0m") -NoNewline
        }
        else { Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor $color -NoNewline }
    }
    Microsoft.PowerShell.Utility\Write-Host ""
}

# Centers a line (made of styled segments) on a screen $Width wide
function Write-CenteredStyled ($Segments, [int]$Width) {
    $len = 0; foreach ($seg in $Segments) { $len += ([string]$seg[0]).Length }
    $pad = [Math]::Max(0, [int][Math]::Floor(($Width - $len) / 2))
    Write-Styled (@(, @((' ' * $pad), 'Gray')) + @($Segments))
}

function Get-ScreenWidth {
    try { return [Math]::Max(60, [int]$Host.UI.RawUI.WindowSize.Width) } catch { return 100 }
}

# Draws the main menu and returns '1', '2', '3' or 'Q'. Menu text is a block (62 wide) centered on the screen; every
# entry has a colored key, a bold title, an italic explanation underneath and a colored tag on the right.
function Show-MainMenu ([string]$Status) {
    $script:UseAnsi = Enable-AnsiConsole
    $items = @(
        @{ Key = '1'; Title = 'Build the job paperwork'; Tag = 'MOST USED'; TagColor = 'Green';
           Lines = @('Reads the .kss files from Tekla and builds the BOM workbook,', 'folder labels, work orders, QPR sheets and the ProNest scripts.') },
        @{ Key = 'U'; Title = 'Update a job already built'; Tag = 'NEW TL / REV'; TagColor = 'Magenta';
           Lines = @('A TL or revision came in: shows what is new, changed or', 'removed and keeps your edits. Folders, work orders etc. follow.') },
        @{ Key = '2'; Title = 'Soundwall post schedule'; Tag = 'CSV'; TagColor = 'Cyan';
           Lines = @('Turns a post-schedule CSV into a KSS file and the same', 'workbook, so soundwall jobs follow the same path.') },
        @{ Key = '3'; Title = 'New blank Required Assm workbook'; Tag = 'TEMPLATE'; TagColor = 'Yellow';
           Lines = @('Writes the empty Required Assm file (TL entry, TL Notes,', 'Exceptions) to fill in for a new job. No .kss needed.') },
        @{ Key = 'H'; Title = 'Getting started guide'; Tag = 'HELP'; TagColor = 'Magenta';
           Lines = @('Opens the step-by-step guide for first-time users.') },
        @{ Key = 'Q'; Title = 'Quit'; Tag = ''; TagColor = 'DarkGray'; Lines = @() }
    )
    while ($true) {
        try { Clear-Host } catch { }
        $w = Get-ScreenWidth; $block = 62; $left = [Math]::Max(0, [int][Math]::Floor(($w - $block) / 2)); $ind = ' ' * $left
        Write-Host ""
        Write-CenteredStyled @(, @(('=' * $block), 'DarkCyan')) $w
        Write-Host ""
        Write-CenteredStyled @(, @('K   I   S   S', 'Cyan', 'b')) $w
        Write-CenteredStyled @(, @('KSS nesting, hardware and shop management', 'Gray', 'i')) $w
        Write-Host ""
        Write-CenteredStyled @(, @(('=' * $block), 'DarkCyan')) $w
        Write-Host ""
        Write-CenteredStyled @(, @('What would you like to do?', 'Yellow', 'b')) $w
        Write-Host ""
        foreach ($it in $items) {
            $titleText = "[$($it.Key)]  $($it.Title)"
            $gap = [Math]::Max(1, $block - 2 - $titleText.Length - $it.Tag.Length)
            $keyColor = if ($it.Key -eq 'Q') { 'DarkGray' } else { 'Green' }
            Write-Styled @(@(($ind + '  '), 'Gray'), @("[$($it.Key)]", $keyColor, 'b'), @('  ', 'Gray'), @($it.Title, 'White', 'b'), @((' ' * $gap), 'Gray'), @($it.Tag, $it.TagColor, 'b'))
            foreach ($ln in $it.Lines) { Write-Styled @(, @("$ind        $ln", 'DarkGray', 'i')) }
            Write-Host ""
        }
        Write-CenteredStyled @(, @(('-' * $block), 'DarkGray')) $w
        Write-CenteredStyled @(@('Type a letter or number, then press ', 'DarkGray', 'i'), @('Enter', 'Yellow', 'b'), @('.  Just Enter = 1', 'DarkGray', 'i')) $w
        Write-CenteredStyled @(, @("Folder: $((Get-Location).Path)", 'DarkGray')) $w
        if ($Status) { Write-CenteredStyled @(, @($Status, 'DarkGreen')) $w }
        Write-Host ""
        Microsoft.PowerShell.Utility\Write-Host (' ' * ([Math]::Max(0, $left + 2)) + '> ') -ForegroundColor Green -NoNewline
        $answer = (Microsoft.PowerShell.Utility\Read-Host).Trim().ToUpper()
        if ($answer -eq '') { return '1' }
        if ($answer -in @('1', '2', '3', 'U', 'Q')) { return $answer }
        if ($answer -eq 'H') {
            $guide = Join-Path (Split-Path -Parent $PSScriptRoot) 'docs\GETTING-STARTED.md'
            try { Start-Process -FilePath $guide } catch { $Status = "Could not open the guide: $guide" }
            continue
        }
        $Status = "'$answer' is not on the list. Pick 1, 2, 3, U, H or Q."
    }
}

# --- FIRST-RUN / EVERY-RUN CHECK ---
# Looks at what the script needs and fixes what it can by itself (unblocking downloaded files, folder write test).
# Excel cannot be installed from here, so that case stops with plain instructions instead of a stack trace.
function Test-KissEnvironment {
    $problems = [System.Collections.Generic.List[string]]::new()
    Write-Section "KISS - checking this computer"

    # Windows + PowerShell
    if ($env:OS -ne 'Windows_NT') { $problems.Add("This tool only runs on Windows (it drives Microsoft Excel).") }
    $psv = $PSVersionTable.PSVersion
    if ($psv.Major -lt 5 -or ($psv.Major -eq 5 -and $psv.Minor -lt 1)) {
        $problems.Add("Windows PowerShell 5.1 or newer is needed (this is $psv). Run Windows Update, then try again.")
    }
    else { Write-Host "  [OK] PowerShell $psv" }

    # Script permission: Start-KISS.bat runs this for this window only; nothing on the PC is changed
    Write-Host "  [OK] Script permission: $(Get-ExecutionPolicy) (no change made to the PC)"

    # Files downloaded from the internet or a zip are 'blocked' by Windows; unblock the tools quietly
    $toolsDir = $PSScriptRoot
    if ($toolsDir) {
        $unblocked = 0
        foreach ($f in @(Get-ChildItem -LiteralPath $toolsDir -Filter *.ps1 -File -ErrorAction SilentlyContinue)) {
            if (Get-Item -LiteralPath $f.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue) {
                try { Unblock-File -LiteralPath $f.FullName -ErrorAction Stop; $unblocked++ } catch { }
            }
        }
        if ($unblocked -gt 0) { Write-Host "  [OK] Unblocked $unblocked downloaded script file(s)" }
    }

    # Can we write the results here?
    $here = (Get-Location).Path
    $probe = Join-Path $here ".kiss-write-test.tmp"
    try {
        [System.IO.File]::WriteAllText($probe, "x")
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        Write-Host "  [OK] Can save files in: $here"
    }
    catch { $problems.Add("KISS cannot save files in this folder: $here`n       Run it from a normal folder (for example the jobs folder), not from Program Files or a read-only drive.") }

    # Excel (the engine builds every workbook through Excel)
    $xlType = $null
    try { $xlType = [type]::GetTypeFromProgID('Excel.Application') } catch { }
    if (-not $xlType) {
        $problems.Add("Microsoft Excel (the desktop app) was not found. Install Microsoft 365 / Office with Excel, sign in once, then run KISS again. Excel on the web or the Microsoft Store 'Excel' viewer will not work.")
    }
    else {
        $xl = $null
        try {
            $xl = New-Object -ComObject Excel.Application
            Write-Host "  [OK] Microsoft Excel $($xl.Version)"
        }
        catch { $problems.Add("Excel is installed but would not start ($($_.Exception.Message)).`n       Open Excel once by hand, sign in or accept the license prompt, close it, then run KISS again.") }
        finally {
            if ($xl) {
                try { $xl.Quit() } catch { }
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch { }
            }
        }
    }

    if ($problems.Count -gt 0) {
        Write-Host ""
        foreach ($p in $problems) { Write-Host "  [X] $p" -ForegroundColor Red }
        Write-Host ""
        Write-Host "  Fix the above and run KISS again. Nothing else needs to be installed." -ForegroundColor Yellow
        Microsoft.PowerShell.Utility\Read-Host "  Press Enter to close" | Out-Null
        exit 1
    }
    Write-Host ""
    Write-Host "  Everything KISS needs is in place. No extra modules or installs are required." -ForegroundColor Green
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
        $sz -match '(NUT|WASHER|HEAVY_HEX)' -or
        ($s -eq "" -and $sz -match '^HSB\b')) {          # loose bolt rows with no shape field (sample job: "HSB 3/4x1-3/4 A325")
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

# --- REQUIRED ASSM WORKBOOK (the run filter, the TL / sequence info, the Exceptions and the folder labels) ---
# Looks for a "... REQUIRED ASSM.xlsx" in the current folder. Its Sheet1 has one column per transmittal:
# TL name in the "TL Name" row, the AFCO assembly numbers underneath, and optional "Paint / TL Details" and
# "Due Date" rows above. A REV transmittal is named TLR####. Each transmittal + sequence (folder) becomes one 3.5" x 1.125" label.
function Select-RequiredAssmFile {
    param(
        [string[]]$ExtraFolders = @(),
        [string]$NamePattern = '(?i)required[\s_-]*assm',
        [string]$ExcludePattern = '',
        [string]$Heading = 'Required Assm workbook',
        [string]$Noun = 'Required Assm workbooks',
        [string]$SkipText = 'Skip (no workbook: all assemblies, or type your own list)'
    )

    # Where to look: working folder, its jobs folder, the folders of the chosen .kss files, the KISS jobs folder, then common spots.
    $roots = @()
    $roots += (Get-Location).Path
    $roots += (Join-Path (Get-Location).Path 'jobs')
    $roots += $ExtraFolders
    if ($PSScriptRoot) { $roots += (Join-Path (Split-Path -Parent $PSScriptRoot) 'jobs') }
    foreach ($sub in 'Downloads', 'Desktop', 'Documents') {
        if ($env:USERPROFILE) { $roots += (Join-Path $env:USERPROFILE $sub) }
    }
    $roots = @($roots | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) } |
        ForEach-Object { (Resolve-Path -LiteralPath $_).Path } | Select-Object -Unique)

    $root = (Get-Location).Path
    $found = @()
    foreach ($r in $roots) {
        $depth = if ($env:USERPROFILE -and $r.StartsWith($env:USERPROFILE, [StringComparison]::OrdinalIgnoreCase) -and -not $r.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { 0 } else { 2 }
        $found += @(Get-ChildItem -LiteralPath $r -Filter *.xlsx -File -Recurse -Depth $depth -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' -and $_.Name -match $NamePattern -and (-not $ExcludePattern -or $_.Name -notmatch $ExcludePattern) })
    }
    $files = @($found | Sort-Object FullName -Unique | Sort-Object LastWriteTime -Descending | Select-Object -First 9)
    Write-Section $Heading
    if ($files.Count -gt 0) {
        Write-Host "Found these $Noun (newest first):"
        for ($i = 0; $i -lt $files.Count; $i++) {
            $f = $files[$i]
            $where = $f.DirectoryName
            if ($where.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { $where = '.' + $where.Substring($root.Length) }
            Write-Host ("  [{0}] {1}   ({2:M/d/yy h:mm tt}, {3})" -f ($i + 1), $f.Name, $f.LastWriteTime, $where)
        }
    }
    else {
        Write-Host "No $Noun found nearby."
    }
    Write-Host "  [B] Browse for a file..."
    Write-Host "  [0] $SkipText"
    $dflt = if ($files.Count -gt 0) { "1" } else { "0" }

    while ($true) {
        $pick = (Read-Host "Which file? (default: $dflt)").Trim()
        if ($pick -eq '') { $pick = $dflt }
        if ($pick -match '^(0|n|no|s|skip)$') { return $null }
        if ($pick -match '^(b|browse)$') {
            try {
                Add-Type -AssemblyName System.Windows.Forms | Out-Null
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog.Filter = "Excel workbooks (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
                $dialog.InitialDirectory = $root
                $dialog.Title = "Select the $Heading"
                if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
                Write-Host "  Nothing picked." -ForegroundColor Yellow
            }
            catch { Write-Host "  The file picker could not open: $($_.Exception.Message)" -ForegroundColor Yellow }
            continue
        }
        if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $files.Count) { return $files[[int]$pick - 1].FullName }
        Write-Host "  Pick a number from the list, B to browse, or 0 to skip." -ForegroundColor Yellow
    }
}

# Reads the workbook in a short-lived Excel of its own so the assembly filter can be built before the main run's Excel starts.
function Read-RequiredAssmStandalone ([string]$Path) {
    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        return @(Read-RequiredAssm -Excel $xl -Path $Path | Where-Object { $_ })
    }
    finally {
        if ($xl) {
            try { $xl.Quit() } catch { }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        }
    }
}

# The Required Assm workbook IS the assembly filter: turns each transmittal column of Sheet1 into one line of the
# whitelist text that Parse-AssemblyWhitelist understands ("TL1296 51311000 51311001 ..."). Only transmittals whose .kss
# is in this run are used. A REV transmittal (TLR####) is written with the R prefix and goes last, so its assemblies
# replace the same assemblies from earlier transmittals.
# Returns Text, per-TL counts (Used), workbook TLs that are not in this run (NotInRun) and run TLs that are not in the workbook (NotInWorkbook).
function Get-WorkbookAssemblyFilter {
    param($WorkbookTLs, [string[]]$RunTLs = @())

    $runNums = @($RunTLs | Where-Object { $_ } | ForEach-Object { [int]($_ -replace '\D', '0') })
    $inRun = @($WorkbookTLs | Where-Object { $runNums.Count -eq 0 -or $runNums -contains [int]$_.Number })
    $lines = [System.Collections.Generic.List[string]]::new()
    $used = [System.Collections.Generic.List[object]]::new()
    foreach ($tl in @($inRun | Sort-Object @{ Expression = { [bool]$_.IsRev } }, Number)) {
        $prefix = if ($tl.IsRev) { 'R' } else { '' }
        $names = [System.Collections.Generic.List[string]]::new()
        foreach ($seq in @($tl.Sequences)) {
            foreach ($a in @($tl.SeqAssemblies[$seq] | Sort-Object)) { $names.Add("$prefix$a") }
        }
        if ($names.Count -eq 0) { continue }
        $lines.Add("TL$($tl.Number) " + ($names -join ' '))
        $used.Add([PSCustomObject]@{ TL = $tl.TL; IsRev = [bool]$tl.IsRev; Count = $names.Count })
    }
    $workbookNums = @($WorkbookTLs | ForEach-Object { [int]$_.Number })
    return [PSCustomObject]@{
        Text          = ($lines -join "`n")
        Used          = @($used)
        AssemblyCount = [int](($used | Measure-Object -Property Count -Sum).Sum)
        NotInRun      = @($WorkbookTLs | Where-Object { $runNums.Count -gt 0 -and $runNums -notcontains [int]$_.Number } | ForEach-Object { $_.TL })
        NotInWorkbook = @($RunTLs | Where-Object { $_ -and $workbookNums -notcontains [int]($_ -replace '\D', '0') })
    }
}

# Date cells come back from Excel as OLE doubles; typed text is kept as typed.
function ConvertTo-DueText ($d) {
    if ($d -is [double]) { return [DateTime]::FromOADate($d).ToString('M/d/yy') }
    return ("$d").Trim()
}

# "Exceptions" tab of the Required Assm workbook: ASSM-specific overrides of the sequence's Paint / Details, Explanation
# and Due Date. $Vals is the tab's cells as a 2-D array (Excel Value2). Only the three Override columns are read
# (the Default / Effective columns are formulas); an ASSM with no override filled in "uses default" and is ignored.
# Returns @{ Map = digits -> { Assembly; Paint; Explanation; DueDate }; Duplicates = digits listed more than once }.
function Read-ExceptionsFromArray ($Vals) {
    $map = @{}
    $dups = [System.Collections.Generic.List[string]]::new()
    if ($Vals -isnot [System.Array] -or $Vals.Rank -ne 2) { return [PSCustomObject]@{ Map = $map; Duplicates = @() } }
    $r0 = $Vals.GetLowerBound(0); $r1 = $Vals.GetUpperBound(0)
    $c0 = $Vals.GetLowerBound(1); $c1 = $Vals.GetUpperBound(1)
    $hr = -1; $asmCol = -1
    for ($r = $r0; $r -le $r1 -and $hr -lt 0; $r++) {
        for ($c = $c0; $c -le $c1; $c++) {
            if (("$($Vals[$r, $c])").Trim() -match '^(?i)ASSM(\s*Number)?$') { $hr = $r; $asmCol = $c; break }
        }
    }
    if ($hr -lt 0) { return [PSCustomObject]@{ Map = $map; Duplicates = @() } }
    $paintCol = -1; $explCol = -1; $dueCol = -1
    for ($c = $c0; $c -le $c1; $c++) {
        $h = ("$($Vals[$hr, $c])").Trim()
        if ($h -match '(?i)^override\s*paint|^override\s*coating') { $paintCol = $c }
        elseif ($h -match '(?i)^override\s*expl') { $explCol = $c }
        elseif ($h -match '(?i)^override\s*due') { $dueCol = $c }
    }
    for ($r = $hr + 1; $r -le $r1; $r++) {
        $digits = ("$($Vals[$r, $asmCol])").Trim() -replace '\D', ''     # drops an R (REV) prefix and anything else
        if ($digits.Length -lt 4) { continue }
        $paint = ""; $expl = ""; $due = ""
        if ($paintCol -ge 0 -and $null -ne $Vals[$r, $paintCol]) { $paint = ("$($Vals[$r, $paintCol])").Trim() }
        if ($explCol -ge 0 -and $null -ne $Vals[$r, $explCol]) { $expl = ("$($Vals[$r, $explCol])").Trim() }
        if ($dueCol -ge 0 -and $null -ne $Vals[$r, $dueCol]) { $due = ConvertTo-DueText $Vals[$r, $dueCol] }
        if (-not $paint -and -not $expl -and -not $due) { continue }       # "Uses default"
        if ($map.ContainsKey($digits)) { $dups.Add($digits) }               # last entry wins
        $map[$digits] = [PSCustomObject]@{ Assembly = $digits; Paint = $paint; Explanation = $expl; DueDate = $due }
    }
    return [PSCustomObject]@{ Map = $map; Duplicates = @($dups) }
}

# Writes a blank Required Assm workbook (no data): the TL entry sheet, the three summary sheets, TL Notes and Exceptions,
# with all formulas, tables and formatting. The workbook itself is embedded at the bottom of this file ($script:RequiredAssmTemplateBase64); this needs
# no Excel to write it, only to open it. Returns the path written, or $null.
function New-RequiredAssmTemplate {
    Write-Section "New Required Assm workbook (blank template)"
    Write-Host "  The file name starts with our job number, e.g.  12345 REQUIRED ASSM.xlsx  (KISS reads that number from the name)."
    Write-Host ""
    $jobIn = (Read-Host "Our job number for the file name (Enter = TEMPLATE, rename it later)").Trim()
    if (-not $jobIn) { $jobIn = "TEMPLATE" }
    $jobIn = $jobIn -replace '[\\/:*?"<>|]', '_'
    $path = Join-Path (Get-Location).Path "$jobIn REQUIRED ASSM.xlsx"
    if (Test-Path -LiteralPath $path) {
        $path = Join-Path (Get-Location).Path ("$jobIn REQUIRED ASSM {0}.xlsx" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Write-Log "That file already exists; writing a new copy instead so nothing is overwritten." "Yellow"
    }
    [System.IO.File]::WriteAllBytes($path, [Convert]::FromBase64String($script:RequiredAssmTemplateBase64))
    Write-Log "Wrote $path" "Green"
    Write-Host ""
    Write-Host "  How to fill it in:" -ForegroundColor Yellow
    Write-Host "    Sheet1 ........ replace an orange Column# header with the TL name (TL1296, REV = TLR1301), then paste its AFCO numbers below it."
    Write-Host "    TL Notes ...... sequences appear on the left by themselves; type paint / explanation / due date in the yellow table on the right."
    Write-Host "    Exceptions .... paste only the ASSM numbers whose paint, explanation or due date differ from their sequence; fill the yellow Override columns."
    Write-Host "  Needs Microsoft 365 Excel (the tabs use LET / LAMBDA / FILTER). Leave this file in the job folder; KISS finds it by its name."
    Write-Host ""
    $open = (Read-Host "Open it in Excel now? [Y]es / [N]o (default: Yes)").Trim()
    if ($open -notmatch '^(n|no)$') { try { Start-Process -FilePath $path } catch { Write-Log "Could not open it: $($_.Exception.Message)" "Yellow" } }
    return $path
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
                    $paintCol = 0; $dueCol = 0; $explCol = 0
                    for ($k = $hc + 1; $k -le [Math]::Min($nCols, $hc + 6); $k++) {
                        $h = ("$($vals[$hr, $k])").Trim()
                        if ($h -match '(?i)^paint|coating') { $paintCol = $k }
                        elseif ($h -match '(?i)^expl') { $explCol = $k }
                        elseif ($h -match '(?i)^due') { $dueCol = $k }
                    }
                    if ($paintCol -eq 0 -and $dueCol -eq 0) { continue }
                    for ($r = $hr + 1; $r -le $nRows; $r++) {
                        $seq = ("$($vals[$r, $hc])").Trim()
                        if ($seq -notmatch '^\d{4}$') { continue }
                        if (-not $seqNotes.ContainsKey($seq)) { $seqNotes[$seq] = @{ Coating = ""; DueDate = ""; Explanation = "" } }
                        if ($paintCol -gt 0 -and -not $seqNotes[$seq].Coating -and $null -ne $vals[$r, $paintCol]) {
                            $seqNotes[$seq].Coating = ("$($vals[$r, $paintCol])").Trim()
                        }
                        if ($explCol -gt 0 -and -not $seqNotes[$seq].Explanation -and $null -ne $vals[$r, $explCol]) {
                            $seqNotes[$seq].Explanation = ("$($vals[$r, $explCol])").Trim()
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
    # ASSM-specific overrides from the "Exceptions" tab (same for every transmittal; read once)
    $exceptions = [PSCustomObject]@{ Map = @{}; Duplicates = @() }
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name -notmatch '(?i)^\s*Exceptions\s*$') { continue }
            $last = $ws.Cells.SpecialCells(11)
            $vals = $ws.Range($ws.Cells.Item(1, 1), $last).Value2
            $exceptions = Read-ExceptionsFromArray $vals
            break
        }
    }
    finally {
        $wb.Close($false)
    }
    foreach ($tlEntry in $result) {
        $tlEntry | Add-Member -NotePropertyName SeqNotes -NotePropertyValue $seqNotes
        $tlEntry | Add-Member -NotePropertyName Exceptions -NotePropertyValue $exceptions.Map
        $tlEntry | Add-Member -NotePropertyName ExceptionDuplicates -NotePropertyValue $exceptions.Duplicates
    }
    return @($result | Sort-Object Number)
}

# Physical label: Dymo address label, 3.5" x 1.125" = 252 pt x 81 pt. Rows 40.5 + 15 + 24.75 + 0.75 (spacer so the bottom border prints) = 81 pt.
# Margins are 0 and the page is set to fit all columns on one page, so the text stays inside the printable area.
function Add-FolderLabelSheet {
    param($Workbook, $Labels)

    # Job, title, customer job and initials come from the Settings tab (named cells Set_*)

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
    & $setWidth $sheet.Columns.Item(1) 120
    & $setWidth $sheet.Columns.Item(2) 132

    # Alignment constants: xlLeft = -4131, xlRight = -4152, xlCenter = -4108
    $i = 0
    foreach ($label in $Labels) {
        $r1 = 4 * $i + 1; $r2 = $r1 + 1; $r3 = $r1 + 2; $r4 = $r1 + 3

        $sheet.Rows.Item($r1).RowHeight = 40.5
        $sheet.Rows.Item($r2).RowHeight = 15
        $sheet.Rows.Item($r3).RowHeight = 24.75
        $sheet.Rows.Item($r4).RowHeight = 0.75      # spacer: the page ends below the last border, so the bottom line is not clipped

        # Line 1: our job - transmittal - sequence (big)
        $cell1 = $sheet.Range($sheet.Cells.Item($r1, 1), $sheet.Cells.Item($r1, 2))
        $cell1.Merge()
        Set-CellFormulaText $cell1 "=TEXTJOIN("" - "",TRUE,Set_OurJob,""$($label.TL)"",""$($label.Sequence)"")"
        $cell1.Font.Size = 26
        if ($label.IsRev) { $cell1.Font.Italic = $true }   # a formula cell cannot italicize just the R of TLR####, so a REV label is italic throughout
        $cell1.HorizontalAlignment = -4108
        $cell1.VerticalAlignment = -4108
        $cell1.ShrinkToFit = $true
        Set-LabelCellBorder $cell1

        # Line 2 (small): project title - customer job on the left, "ISU'D: initials" on the right
        $cellL = $sheet.Cells.Item($r2, 1)
        Set-CellFormulaText $cellL "=TEXTJOIN("" - "",TRUE,Set_Title,Set_CustomerJob)"
        $cellL.Font.Size = 10
        $cellL.HorizontalAlignment = -4131
        $cellL.VerticalAlignment = -4108
        $cellL.ShrinkToFit = $true
        Set-LabelCellBorder $cellL
        $cellR = $sheet.Cells.Item($r2, 2)
        # REVISES (REV labels only) sits on this line, left of the "ISU'D: initials" text
        $revRef = Get-SeqNoteRef "E" "`"$($label.TL)|$($label.Sequence)`""
        Set-CellFormulaText $cellR "=TRIM(IF($revRef<>"""",""REVISES ""&$revRef&""   "","""")&TRIM(Set_IssuedLabel&"" ""&Set_Initials))"
        $cellR.Font.Size = 10
        $cellR.HorizontalAlignment = -4152
        $cellR.VerticalAlignment = -4108
        $cellR.ShrinkToFit = $true
        Set-LabelCellBorder $cellR

        # Line 3: comments / paint info on the left, due date on the right (blank date = slashes to fill in later)
        $cellC = $sheet.Cells.Item($r3, 1)
        $cellC.NumberFormat = "General"
        $cellC.Formula = "=" + (Get-SeqNoteRef "M" "`"$($label.TL)|$($label.Sequence)`"")   # paint / exception text from the Sequence Notes tab (REVISES is on line 2)
        $cellC.Font.Size = 10
        $cellC.HorizontalAlignment = -4131
        $cellC.VerticalAlignment = -4108
        $cellC.WrapText = $true          # long notes wrap onto a second line instead of shrinking to tiny text
        $cellC.ShrinkToFit = $false
        Set-LabelCellBorder $cellC
        $cellD = $sheet.Cells.Item($r3, 2)
        $cellD.NumberFormat = "General"
        $cellD.Formula = "=" + (Get-SeqNoteRef "I" "`"$($label.TL)|$($label.Sequence)`"")   # "DUE m/d/yy", or slashes to fill in
        $cellD.Font.Size = 17
        $cellD.HorizontalAlignment = -4152
        $cellD.VerticalAlignment = -4108
        $cellD.ShrinkToFit = $true
        Set-LabelCellBorder $cellD

        $i++
    }

    $lastRow = 4 * $Labels.Count
    $sheet.Range("D1").Value2 = "Print this tab on the Dymo (3.5 x 1.125 in address label). One label per page."
    $sheet.Range("D1").Font.Bold = $false
    $sheet.Range("D1").Font.Italic = $true

    Set-LabelPrintSetup -Sheet $sheet -LastCol "B" -RowsPerLabel 4 -LabelCount $Labels.Count
    return $sheet
}

# Thin black border on all four edges of a label cell (applied to the whole range, so merged cells get a full outline).
function Set-LabelCellBorder ($Range) {
    foreach ($edge in 7, 8, 9, 10) {            # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
        $b = $Range.Borders.Item($edge)
        $b.LineStyle = 1                          # xlContinuous
        $b.Weight = 2                             # xlThin
        $b.Color = 0
    }
}

# Print setup shared by the label tabs: one 3.5 x 1.125 in label per page, margins 0, fit all columns on one page.
function Set-LabelPrintSetup {
    param($Sheet, [string]$LastCol, [int]$RowsPerLabel, [int]$LabelCount)
    try {
        $ps = $Sheet.PageSetup
        $ps.PrintArea = "`$A`$1:`$$LastCol`$$($RowsPerLabel * $LabelCount)"
        $ps.Orientation = 2                      # landscape: the long side (3.5 in) is the width
        $ps.LeftMargin = 0; $ps.RightMargin = 0
        $ps.TopMargin = 0; $ps.BottomMargin = 0
        $ps.HeaderMargin = 0; $ps.FooterMargin = 0
        $ps.CenterHorizontally = $true
        $ps.CenterVertically = $true
        $Sheet.ResetAllPageBreaks()
        for ($b = 1; $b -lt $LabelCount; $b++) {
            [void]$Sheet.HPageBreaks.Add($Sheet.Cells.Item($RowsPerLabel * $b + 1, 1))
        }
        $ps.Zoom = $false                        # after ResetAllPageBreaks (it resets the scaling): fit all columns on one page
        $ps.FitToPagesWide = 1
        $ps.FitToPagesTall = $false
    }
    catch {
        Write-Log "$($Sheet.Name) tab built, but page setup failed (is a printer installed? select the Dymo and set the page to 3.5 x 1.125 in): $($_.Exception.Message)" "Yellow"
    }
}

# Transmittals (other than the REV itself) that hold any of the REV's assembly numbers in the given sequences.
# Non-REV transmittals are preferred; other REVs are returned only when no non-REV one matches.
function Get-RevisedTransmittals {
    param($Tl, $AllTLs, [string[]]$Sequences)
    $hits = @($AllTLs | Where-Object {
        $other = $_
        $other.TL -ne $Tl.TL -and @($Sequences | Where-Object { $Tl.SeqAssemblies.ContainsKey($_) -and $other.SeqAssemblies.ContainsKey($_) -and $Tl.SeqAssemblies[$_].Overlaps($other.SeqAssemblies[$_]) }).Count -gt 0
    } | Sort-Object Number)
    $plain = @($hits | Where-Object { -not $_.IsRev })
    if ($plain.Count -gt 0) { return $plain }
    return $hits
}

# "TL Labels" tab: one 3.5 x 1.125 in label per transmittal. Row 1 (33 pt): the TL big (italic R on a REV, with a small
# right-aligned "REVISES ..."); row 2 (30 pt): its sequences big, side by side; row 3 (18 pt): each sequence's due date small.
# Built on a 12-column grid of 21 pt so any number of sequences spreads evenly.
function Add-TlLabelSheet {
    param($Workbook, $TLs, $AllTLs)

    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "TL Labels"
    $sheet.Cells.NumberFormat = "@"
    $sheet.Cells.Font.Name = "Arial"
    $sheet.Cells.Font.Bold = $true
    for ($c = 1; $c -le 12; $c++) {
        $col = $sheet.Columns.Item($c)
        $w = 21 / 5.4
        for ($k = 0; $k -lt 8; $k++) {
            $col.ColumnWidth = $w
            $cur = [double]$col.Width
            if ([Math]::Abs($cur - 21) -lt 0.4) { break }
            $w = $w * (21 / $cur)
        }
    }

    # Column spans (12-column grid) per sequence count; beyond 6 sequences each gets 2 columns or less and shrinks to fit
    $spanPlans = @{ 1 = @(12); 2 = @(6, 6); 3 = @(4, 4, 4); 4 = @(3, 3, 3, 3); 5 = @(2, 2, 3, 2, 3); 6 = @(2, 2, 2, 2, 2, 2) }

    $i = 0
    foreach ($tl in $TLs) {
        $r1 = 4 * $i + 1; $r2 = $r1 + 1; $r3 = $r1 + 2; $r4 = $r1 + 3
        $sheet.Rows.Item($r1).RowHeight = 33
        $sheet.Rows.Item($r2).RowHeight = 30
        $sheet.Rows.Item($r3).RowHeight = 17.25
        $sheet.Rows.Item($r4).RowHeight = 0.75      # spacer: the page ends below the last border, so the bottom line is not clipped

        # Row 1: the transmittal (a REV gets the left 8 columns plus a small "REVISES ..." on the right 4)
        $revises = if ($tl.IsRev) { @(Get-RevisedTransmittals -Tl $tl -AllTLs $AllTLs -Sequences $tl.Sequences | ForEach-Object { $_.TL }) } else { @() }
        $tlSpan = if ($revises.Count -gt 0) { 8 } else { 12 }
        $tlCell = $sheet.Range($sheet.Cells.Item($r1, 1), $sheet.Cells.Item($r1, $tlSpan))
        $tlCell.Merge()
        $tlCell.Value2 = $tl.TL
        $tlCell.Font.Size = 26
        $tlCell.HorizontalAlignment = -4108
        $tlCell.VerticalAlignment = -4108
        $tlCell.ShrinkToFit = $true
        Set-LabelCellBorder $tlCell
        if ($tl.IsRev) { $tlCell.Characters(3, 1).Font.Italic = $true }   # the R of TLR####
        if ($revises.Count -gt 0) {
            $revCell = $sheet.Range($sheet.Cells.Item($r1, 9), $sheet.Cells.Item($r1, 12))
            $revCell.Merge()
            $revCell.Value2 = "REVISES $($revises -join ', ')"
            $revCell.Font.Size = 10
            $revCell.HorizontalAlignment = -4152
            $revCell.VerticalAlignment = -4108
            $revCell.ShrinkToFit = $true
            Set-LabelCellBorder $revCell
        }

        # Rows 2 and 3: each sequence big, its due date small underneath
        $seqs = @($tl.Sequences)
        $spans = if ($spanPlans.ContainsKey($seqs.Count)) { $spanPlans[$seqs.Count] } else { @(for ($n = 0; $n -lt $seqs.Count; $n++) { 12 / $seqs.Count }) }
        $col = 1
        for ($n = 0; $n -lt $seqs.Count; $n++) {
            $last = [Math]::Min(12, $col + [int][Math]::Ceiling($spans[$n]) - 1)
            $note = $tl.SeqNotes[$seqs[$n]]
            $due = if ($note -and $note.DueDate) { $note.DueDate } else { $tl.DueDate }

            $seqCell = $sheet.Range($sheet.Cells.Item($r2, $col), $sheet.Cells.Item($r2, $last))
            $seqCell.Merge()
            $seqCell.Value2 = $seqs[$n]
            $seqCell.Font.Size = 24
            $seqCell.HorizontalAlignment = -4108
            $seqCell.VerticalAlignment = -4108
            $seqCell.ShrinkToFit = $true
            Set-LabelCellBorder $seqCell

            $dueCell = $sheet.Range($sheet.Cells.Item($r3, $col), $sheet.Cells.Item($r3, $last))
            $dueCell.Merge()
            $dueCell.NumberFormat = "General"
            $dueCell.Formula = "=" + (Get-SeqNoteRef "I" "`"$($tl.TL)|$($seqs[$n])`"")   # follows the Sequence Notes tab
            $dueCell.Font.Size = 10
            $dueCell.HorizontalAlignment = -4108
            $dueCell.VerticalAlignment = -4108
            $dueCell.ShrinkToFit = $true
            Set-LabelCellBorder $dueCell
            $col = $last + 1
        }
        $i++
    }

    Set-LabelPrintSetup -Sheet $sheet -LastCol "L" -RowsPerLabel 4 -LabelCount $TLs.Count
    return $sheet
}

# --- SETTINGS TAB: values the setup prompts used to bake into cells now live here as named cells (Set_*) that the labels, work orders and Statistics follow ---
# Name, label, value, number format ("@" = text), where it is used. Shared by the Settings tab and the update (which compares the old tab with these).
function Get-SettingsList {
    param([string]$Title, [string]$CustomerJob, [string]$OurJob, [string]$Initials, [string]$ProjectManagers)
    $issued = [double][DateTime]::Today.ToOADate()
    return @(
        @("Set_Title", "Project title (label line 2)", $Title, "@", "Folder Labels, Statistics"),
        @("Set_CustomerJob", "Customer job #", $CustomerJob, "@", "Folder Labels, TERRA Work Order, Statistics"),
        @("Set_OurJob", "Our job #", $OurJob, "@", "Folder Labels (line 1), TERRA Work Order (ATF JOB / WO #)"),
        @("Set_IssuedLabel", "Issued prefix", "IS:", "@", "Folder Labels (printed before the initials)"),
        @("Set_Initials", "Initials (who made the folders)", $Initials, "@", "Folder Labels (line 2, right)"),
        @("Set_ProjectManagers", "Project managers (PM box)", $ProjectManagers, "@", "TERRA Work Order"),
        @("Set_IssuedDate", "Issued to shop date", $issued, "m/d/yy", "TERRA Work Order"),
        @("Set_DefaultDue", "Due date text when a sequence has none", "3 WKS AFTER MATERIAL ARRIVES", "@", "TERRA Work Order"),
        @("Set_DefaultNotes", "Default paint / notes (none listed for the sequence)", "-", "@", "TERRA Work Order (NOTES column)"),
        @("Set_CWI", "CWI required", "N/A", "@", "TERRA Work Order"),
        @("Set_WarnDays", "Flag a due date as DUE SOON within (days)", 7, "0", "Statistics (schedule)")
    )
}

function Add-SettingsSheet {
    param($Workbook, [string]$Title, [string]$CustomerJob, [string]$OurJob, [string]$Initials, [string]$ProjectManagers)

    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "Settings"
    $sheet.Cells.Font.Name = "Arial"
    $sheet.Cells.Font.Size = 10

    $sheet.Range("A1:C1").Merge()
    $sheet.Cells.Item(1, 1).Value2 = "SETTINGS - change the yellow cells; labels, work orders and Statistics follow"
    $sheet.Range("A1:C1").Font.Bold = $true
    $sheet.Range("A1:C1").Font.Size = 13
    $sheet.Range("A1:C1").Font.Color = Get-OleColor 255 255 255
    $sheet.Range("A1:C1").Interior.Color = Get-OleColor 27 54 93
    $sheet.Range("A1:C1").VerticalAlignment = -4108
    $sheet.Rows.Item(1).RowHeight = 26

    $sheet.Cells.Item(3, 1).Value2 = "Setting"
    $sheet.Cells.Item(3, 2).Value2 = "Value"
    $sheet.Cells.Item(3, 3).Value2 = "Used on"
    $sheet.Range("A3:C3").Font.Bold = $true
    $sheet.Range("A3:C3").Interior.Color = Get-OleColor 216 228 188

    $settings = Get-SettingsList -Title $Title -CustomerJob $CustomerJob -OurJob $OurJob -Initials $Initials -ProjectManagers $ProjectManagers
    $row = 4
    foreach ($s in $settings) {
        $sheet.Cells.Item($row, 1).Value2 = $s[1]
        $valueCell = $sheet.Cells.Item($row, 2)
        $valueCell.NumberFormat = $s[3]
        $valueCell.Value2 = $s[2]
        $valueCell.Interior.Color = Get-OleColor 255 242 204
        $valueCell.HorizontalAlignment = -4131
        $sheet.Cells.Item($row, 3).Value2 = $s[4]
        $sheet.Cells.Item($row, 3).Font.Color = Get-OleColor 110 110 110
        [void]$Workbook.Names.Add($s[0], "=Settings!`$B`$$row")
        $row++
    }
    $sheet.Range("A3:C$($row - 1)").Borders.LineStyle = 1
    $sheet.Range("A3:C$($row - 1)").Borders.Color = Get-OleColor 200 200 200
    $sheet.Columns.Item(1).ColumnWidth = 48
    $sheet.Columns.Item(2).ColumnWidth = 34
    $sheet.Columns.Item(3).ColumnWidth = 58
    $sheet.Cells.Item($row + 1, 1).Value2 = "Setup still asks for these once and fills them in here as the starting values."
    $sheet.Cells.Item($row + 1, 1).Font.Italic = $true
    return $sheet
}

# Writes a formula into a cell/merged range on a sheet whose cells are formatted as text (a "@" cell would show the formula as text)
function Set-CellFormulaText ($Range, [string]$Formula, [string]$NumberFormat = "General") {
    $Range.NumberFormat = $NumberFormat
    $Range.Cells.Item(1, 1).Formula = $Formula
}
# --- STATISTICS EXTRAS: schedule, sequence and plate-vs-profile tables plus headline ratios, all live formulas on the Master BOM table ---
# Folders = one row per (TL, sequence) from the Required Assm workbook ($Folders: TL, Sequence, DueDate); Due dates are cells you can edit.
function Add-DashboardExtras {
    param($Dashboard, $MasterData, $Folders)

    $wt = "MasterBOMTable[Total Wt (lbs)]"
    $qty = "MasterBOMTable[Qty]"
    $tlCol = "MasterBOMTable[TL '#]"   # # is a special character in structured references
    $asmCol = "MasterBOMTable[Assembly]"
    $catCol = "MasterBOMTable[Category]"
    $top = $Dashboard.UsedRange.Rows.Count + 3

    $header = {
        param($r, $c1, $c2, $title, $cols)
        $Dashboard.Range($Dashboard.Cells.Item($r, $c1), $Dashboard.Cells.Item($r, $c2)).Merge()
        $Dashboard.Cells.Item($r, $c1).Value2 = $title
        $bar = $Dashboard.Range($Dashboard.Cells.Item($r, $c1), $Dashboard.Cells.Item($r, $c2))
        $bar.Font.Bold = $true; $bar.Font.Size = 10; $bar.Font.Color = Get-OleColor 255 255 255
        $bar.Interior.Color = Get-OleColor 31 78 120; $bar.HorizontalAlignment = -4108
        for ($i = 0; $i -lt $cols.Count; $i++) { $Dashboard.Cells.Item($r + 1, $c1 + $i).Value2 = $cols[$i] }
        $sub = $Dashboard.Range($Dashboard.Cells.Item($r + 1, $c1), $Dashboard.Cells.Item($r + 1, $c2))
        $sub.Font.Bold = $true; $sub.Font.Size = 9; $sub.Interior.Color = Get-OleColor 216 228 188
        $sub.Borders.LineStyle = 1; $sub.Borders.Weight = 2; $sub.Borders.Color = Get-OleColor 166 166 166
    }
    $grid = {
        param($r1, $r2, $c1, $c2)
        $g = $Dashboard.Range($Dashboard.Cells.Item($r1, $c1), $Dashboard.Cells.Item($r2, $c2))
        $g.Borders.LineStyle = 1; $g.Borders.Weight = 2; $g.Borders.Color = Get-OleColor 220 225 230
        $g.HorizontalAlignment = -4131
    }

    # Distinct keys from the BOM, sorted; sequence = first 4 characters of the assembly
    $rows = @($MasterData | ForEach-Object { [PSCustomObject]@{ TL = "$($_.TL)"; Seq = if ("$($_.Assembly)".Length -ge 4) { "$($_.Assembly)".Substring(0, 4) } else { "$($_.Assembly)" } } } | Sort-Object TL, Seq -Unique)
    $tls = @($rows | ForEach-Object { $_.TL } | Sort-Object -Unique)
    $seqs = @($rows | ForEach-Object { $_.Seq } | Sort-Object -Unique)

    # --- TL / sequence detail (J:M) ---
    & $header $top 10 13 "TL / SEQUENCE DETAIL" @("TL #", "Sequence", "Assemblies", "Total Wt (lbs)")
    $r = $top + 2; $detailFirst = $r
    foreach ($k in $rows) {
        $Dashboard.Cells.Item($r, 10).NumberFormat = "@"; $Dashboard.Cells.Item($r, 10).Value2 = $k.TL
        $Dashboard.Cells.Item($r, 11).NumberFormat = "@"; $Dashboard.Cells.Item($r, 11).Value2 = $k.Seq
        $Dashboard.Cells.Item($r, 12).Formula2 = "=IFERROR(ROWS(UNIQUE(FILTER($asmCol,($tlCol=J$r)*(LEFT($asmCol,4)=K$r)))),0)"
        $Dashboard.Cells.Item($r, 13).Formula = "=SUMPRODUCT(($tlCol=J$r)*(LEFT($asmCol,4)=K$r)*$wt)"
        $Dashboard.Cells.Item($r, 13).NumberFormat = "#,##0.00"
        $r++
    }
    $detailLast = $r - 1
    & $grid $detailFirst $detailLast 10 13

    # --- sequence summary across TLs (E:H) ---
    & $header $top 5 8 "SEQUENCE SUMMARY (ALL TLs)" @("Sequence", "Assemblies", "Part Pieces", "Total Wt (lbs)")
    $r = $top + 2; $seqFirst = $r
    foreach ($s in $seqs) {
        $Dashboard.Cells.Item($r, 5).NumberFormat = "@"; $Dashboard.Cells.Item($r, 5).Value2 = $s
        $Dashboard.Cells.Item($r, 6).Formula2 = "=IFERROR(ROWS(UNIQUE(FILTER($asmCol,LEFT($asmCol,4)=E$r))),0)"
        $Dashboard.Cells.Item($r, 7).Formula = "=SUMPRODUCT((LEFT($asmCol,4)=E$r)*$qty)"
        $Dashboard.Cells.Item($r, 8).Formula = "=SUMPRODUCT((LEFT($asmCol,4)=E$r)*$wt)"
        $Dashboard.Cells.Item($r, 7).NumberFormat = "#,##0"
        $Dashboard.Cells.Item($r, 8).NumberFormat = "#,##0.00"
        $r++
    }
    $seqLast = $r - 1
    & $grid $seqFirst $seqLast 5 8

    # --- plate vs profile by TL (T:W) ---
    & $header $top 20 23 "PLATE vs PROFILE BY TL" @("TL #", "Plate (lbs)", "Profile (lbs)", "Plate % of both")
    $r = $top + 2; $ppFirst = $r
    foreach ($tl in $tls) {
        $Dashboard.Cells.Item($r, 20).NumberFormat = "@"; $Dashboard.Cells.Item($r, 20).Value2 = $tl
        $Dashboard.Cells.Item($r, 21).Formula = "=SUMIFS($wt,$tlCol,T$r,$catCol,`"Plate`")"
        $Dashboard.Cells.Item($r, 22).Formula = "=SUMIFS($wt,$tlCol,T$r,$catCol,`"Profile`")"
        $Dashboard.Cells.Item($r, 23).Formula = "=IFERROR(U$r/(U$r+V$r),0)"
        $Dashboard.Cells.Item($r, 21).NumberFormat = "#,##0"; $Dashboard.Cells.Item($r, 22).NumberFormat = "#,##0"; $Dashboard.Cells.Item($r, 23).NumberFormat = "0.0%"
        $r++
    }
    $ppLast = $r - 1
    $Dashboard.Cells.Item($r, 20).Value2 = "TOTAL"
    $Dashboard.Cells.Item($r, 21).Formula = "=SUM(U${ppFirst}:U$ppLast)"
    $Dashboard.Cells.Item($r, 22).Formula = "=SUM(V${ppFirst}:V$ppLast)"
    $Dashboard.Cells.Item($r, 23).Formula = "=IFERROR(U$r/(U$r+V$r),0)"
    $Dashboard.Cells.Item($r, 21).NumberFormat = "#,##0"; $Dashboard.Cells.Item($r, 22).NumberFormat = "#,##0"; $Dashboard.Cells.Item($r, 23).NumberFormat = "0.0%"
    $Dashboard.Range($Dashboard.Cells.Item($r, 20), $Dashboard.Cells.Item($r, 23)).Font.Bold = $true
    & $grid $ppFirst $r 20 23

    # --- delivery schedule (O:R): one row per folder (TL + sequence) with an editable due date ---
    $schedFirst = $null; $schedLast = $null
    & $header $top 15 18 "FOLDER SCHEDULE (edit the due dates)" @("TL / Sequence", "Due", "Days Left", "Status")
    $r = $top + 2; $schedFirst = $r
    foreach ($fo in @($Folders)) {
        if (-not $fo) { continue }
        $Dashboard.Cells.Item($r, 15).NumberFormat = "@"; $Dashboard.Cells.Item($r, 15).Value2 = "$($fo.TL) / $($fo.Sequence)"
        $d = [DateTime]::MinValue
        $Dashboard.Cells.Item($r, 16).NumberFormat = "m/d/yy"
        if ($fo.DueDate -and [DateTime]::TryParse("$($fo.DueDate)", [System.Globalization.CultureInfo]::GetCultureInfo("en-US"), [System.Globalization.DateTimeStyles]::None, [ref]$d)) { $Dashboard.Cells.Item($r, 16).Value2 = [double]$d.ToOADate() }
        $Dashboard.Cells.Item($r, 16).Interior.Color = Get-OleColor 255 242 204
        $Dashboard.Cells.Item($r, 17).Formula = "=IF(P$r=`"`",`"`",P$r-TODAY())"
        $Dashboard.Cells.Item($r, 17).NumberFormat = "0"
        $Dashboard.Cells.Item($r, 18).Formula = "=IF(P$r=`"`",`"NO DATE`",IF(P$r<TODAY(),`"OVERDUE`",IF(P$r-TODAY()<=Set_WarnDays,`"DUE SOON`",`"ON TRACK`")))"
        $r++
    }
    if ($r -gt $top + 2) {
        $schedLast = $r - 1
        & $grid $schedFirst $schedLast 15 18
        $status = $Dashboard.Range("R${schedFirst}:R$schedLast")
        foreach ($rule in @(@("OVERDUE", 192, 0, 0), @("DUE SOON", 191, 120, 0), @("ON TRACK", 0, 128, 0))) {
            $fc = $status.FormatConditions.Add(1, 3, "=`"$($rule[0])`"")
            $fc.Font.Bold = $true
            $fc.Font.Color = Get-OleColor $rule[1] $rule[2] $rule[3]
        }
    }
    else {
        $Dashboard.Cells.Item($r, 15).Value2 = "No Required Assm workbook this run"
        $Dashboard.Cells.Item($r, 15).Font.Italic = $true
    }

    # --- headline ratios (A:C) ---
    & $header $top 1 3 "HEADLINES & RATIOS" @("Metric", "Value", "Source / Scope")
    $dRange = "J${detailFirst}:J$detailLast"; $dSeq = "K${detailFirst}:K$detailLast"; $dWt = "M${detailFirst}:M$detailLast"
    $headlines = @(
        @("TL + sequence combinations", "=ROWS($dRange)", "#,##0", "TL / Sequence Detail"),
        @("Plate share of weight", "=IFERROR(SUMIFS($wt,$catCol,`"Plate`")/SUM($wt),0)", "0.0%", "Master BOM [Category]"),
        @("Profile share of weight", "=IFERROR(SUMIFS($wt,$catCol,`"Profile`")/SUM($wt),0)", "0.0%", "Master BOM [Category]"),
        @("Hardware share of weight", "=IFERROR(SUMIFS($wt,$catCol,`"Hardware`")/SUM($wt),0)", "0.0%", "Master BOM [Category]"),
        @("Heaviest TL / sequence", "=IFERROR(INDEX($dRange,MATCH(MAX($dWt),$dWt,0))&`" / `"&INDEX($dSeq,MATCH(MAX($dWt),$dWt,0)),`"`")", "General", "TL / Sequence Detail"),
        @("Heaviest TL / sequence (tons)", "=MAX($dWt)/2000", "#,##0.0", "TL / Sequence Detail"),
        @("Average tons per TL / sequence", "=IFERROR(SUM($dWt)/ROWS($dWt)/2000,0)", "#,##0.0", "TL / Sequence Detail"),
        @("Plate pieces", "=SUMIFS($qty,$catCol,`"Plate`")", "#,##0", "Master BOM [Category]"),
        @("Profile pieces", "=SUMIFS($qty,$catCol,`"Profile`")", "#,##0", "Master BOM [Category]")
    )
    if ($schedLast) {
        $sR = "R${schedFirst}:R$schedLast"; $sP = "P${schedFirst}:P$schedLast"; $sL = "O${schedFirst}:O$schedLast"
        $headlines += @(
            @("Folders (TL + sequence)", "=ROWS($sL)", "#,##0", "Folder Schedule"),
            @("REV folders", "=COUNTIF($sL,`"TLR*`")", "#,##0", "Folder Schedule"),
            @("Next due date", "=IFERROR(1/(1/MINIFS($sP,$sP,`">=`"&TODAY())),`"none upcoming`")", "m/d/yy", "Folder Schedule"),
            @("Days to next due", "=IFERROR({NEXT}-TODAY(),`"-`")", "0", "Folder Schedule"),
            @("Overdue folders", "=COUNTIF($sR,`"OVERDUE`")", "#,##0", "Folder Schedule"),
            @("Due soon (Settings: warn days)", "=COUNTIF($sR,`"DUE SOON`")", "#,##0", "Folder Schedule"),
            @("Folders with no date", "=COUNTIF($sR,`"NO DATE`")", "#,##0", "Folder Schedule")
        )
    }
    if ($script:ExceptionRows.Count -gt 0 -or $script:ExceptionNotInRun.Count -gt 0) {
        $headlines += @(
            @("Assemblies with exceptions", "=COUNTIF(Exc_Status,`"Exception set`")", "#,##0", "Assembly Exceptions"),
            @("Exception ASSMs not in this run", "=COUNTIF(Exc_Status,`"Not in this run*`")", "#,##0", "Assembly Exceptions")
        )
    }
    $r = $top + 2; $hFirst = $r
    foreach ($h in $headlines) {
        $Dashboard.Cells.Item($r, 1).Value2 = $h[0]
        $Dashboard.Cells.Item($r, 2).NumberFormat = $h[2]
        $Dashboard.Cells.Item($r, 2).Formula = ($h[1] -replace '\{NEXT\}', "B$($r - 1)")
        $Dashboard.Cells.Item($r, 3).Value2 = $h[3]
        $r++
    }
    & $grid $hFirst ($r - 1) 1 3
    $Dashboard.Range("B${hFirst}:B$($r - 1)").HorizontalAlignment = -4152
    $Dashboard.Cells.Item($top - 1, 1).Value2 = "SCHEDULE, SEQUENCE & MATERIAL MIX (live formulas; due dates and warn days are editable cells)"
    $Dashboard.Cells.Item($top - 1, 1).Font.Bold = $true
}
# --- TERRA WORK ORDER (one letter page per transmittal + sequence, laid out like the Tekla UDR_TERRA_Work_Order report) ---
# Assembly weight comes from the KSS parts (Tekla exports carry no weights): rolled shapes by lb/ft x length, plates and bars by volume.
function Get-KssPartWeight ($shape, $size, $lengthStr, $qty) {
    $lengthIn = 0.0
    $lenNum = 0.0
    if ([double]::TryParse(("$lengthStr").Trim(), [ref]$lenNum)) { $lengthIn = $lenNum / 25.4 }   # KISS lengths are millimeters
    $lengthFt = $lengthIn / 12
    $s = "$shape".Trim().ToUpper()
    $tok = @(("$size" -replace '"', '') -split '[xX]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $steel = 0.2836   # lb per cubic inch
    $each = 0.0
    switch -Regex ($s) {
        '^(PL|BPL|SHT|FB|BAR)$' {
            if ($tok.Count -ge 2) { $each = [double](Get-DecimalThickness $tok[0]) * [double](Get-DecimalThickness $tok[1]) * $lengthIn * $steel }
        }
        '^(HSS|TS|TUBE)$' {
            if ($tok.Count -ge 3) {
                $b = [double](Get-DecimalThickness $tok[0]); $h = [double](Get-DecimalThickness $tok[1]); $t = [double](Get-DecimalThickness $tok[2])
                $area = 2 * $t * ($b + $h - 2 * $t) - 3 * (4 - [Math]::PI) * $t * $t      # outside radius 2t, inside radius t
                $each = $area * 3.4026 * $lengthFt      # lb/ft per in2 that reproduces the AISC tube weights (HSS16X8X1/2 = 76.07 lb/ft)
            }
            elseif ($tok.Count -eq 2) {
                $d = [double](Get-DecimalThickness $tok[0]); $t = [double](Get-DecimalThickness $tok[1])
                $each = [Math]::PI / 4 * ($d * $d - [Math]::Pow($d - 2 * $t, 2)) * $steel * 12 * $lengthFt
            }
        }
        '^L$' {
            if ($tok.Count -ge 3) {
                $a = [double](Get-DecimalThickness $tok[0]); $b = [double](Get-DecimalThickness $tok[1]); $t = [double](Get-DecimalThickness $tok[2])
                $each = $t * ($a + $b - $t) * $steel * 12 * $lengthFt
            }
        }
        '^(W|M|S|HP|C|MC|WT|MT|ST)$' {
            if ($size -match '[Xx](\d+(?:\.\d+)?)\s*$') { $each = [double]$Matches[1] * $lengthFt }   # AISC designation ends in lb/ft
        }
    }
    return $each * $qty
}


# Formula text (no leading "=") that returns a column of the "Sequence Notes" tab for the row whose key is KeyExpr
# (key = "<TL>|<sequence>"); blank when there is no such row or the cell is empty. Goes through workbook names
# (Seq_*, defined by Add-SequenceNotesSheet): a direct 'Sequence Notes'!F:F reference inside the Master BOM table was saved as a broken external link.
function Get-SeqNoteRef ($Column, $KeyExpr) {
    $name = @{ E = "Seq_Revises"; F = "Seq_Due"; G = "Seq_Paint"; H = "Seq_LabelNotes"; I = "Seq_LabelDue"; J = "Seq_Expl"; M = "Seq_LabelPaint" }[$Column]
    'IFERROR(INDEX({0},MATCH({1},Seq_Key,0))&"","")' -f $name, $KeyExpr
}

# Same idea for the "Assembly Exceptions" tab (names Exc_*, defined by Add-AssemblyExceptionsSheet): the EFFECTIVE value
# (override if one is set, else the sequence default) for the assembly whose number is KeyExpr; blank when the assembly
# has no row there. Column: Paint | Expl | Due | Status.
function Get-ExcRef ($Column, $KeyExpr) {
    $name = @{ Paint = "Exc_Paint"; Expl = "Exc_Expl"; Due = "Exc_Due"; Status = "Exc_Status" }[$Column]
    'IFERROR(INDEX({0},MATCH({1},Exc_Key,0))&"","")' -f $name, $KeyExpr
}

# NOTES cell of one work-order item: the folder's paint / notes (Sequence Notes), or for an assembly with an exception its
# effective paint, plus " - <explanation>" and " - DUE <date>" when the exception carries them (the due only when it differs
# from the folder's due date, which is what the work order header shows).
function Get-WorkOrderNoteFormula ($Tl, $Sequence, $Assembly) {
    $key = "`"$Tl|$Sequence`""
    $paintRef = Get-SeqNoteRef "G" $key
    $base = "IF($paintRef=`"`",IFERROR(Set_DefaultNotes,`"-`"),$paintRef)"
    $digits = "$Assembly" -replace '\D', ''
    if (-not $script:AsmExceptions -or -not $script:AsmExceptions.ContainsKey($digits)) { return "=$base" }
    $a = "`"$digits`""
    $p = Get-ExcRef "Paint" $a; $d = Get-ExcRef "Due" $a
    $dueSeq = Get-SeqNoteRef "F" $key
    $text = "=IF($p<>`"`",$p,$base)"
    # the explanation is shown only when this ASSM's exception carries its own (not the sequence's general one)
    if ($script:AsmExceptions[$digits].Explanation) { $e = Get-ExcRef "Expl" $a; $text += "&IF($e<>`"`",`" - `"&$e,`"`")" }
    return $text + "&IF(AND($d<>`"`",$d<>$dueSeq),`" - DUE `"&$d,`"`")"
}

# "Sequence Notes" tab: the one place to edit due date and paint/notes. One row per transmittal + sequence (after the REV move).
# The Master BOM, Folder Labels, TL Labels and TERRA Work Order read these cells by formula (key in column A = "<TL>|<sequence>").
function Add-SequenceNotesSheet {
    param($Workbook, $Rows)

    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "Sequence Notes"
    $sheet.Cells.Font.Name = "Arial"
    $n = @($Rows).Count
    $sheet.Range("B2:G$($n + 2)").NumberFormat = "@"       # typed dates stay text, so the labels print them as typed
    $sheet.Range("H2:M$($n + 2)").NumberFormat = "General"  # formula columns (label notes, due, paint) are never Text
    # Everything goes down in ONE array write (header row, data, formulas, note): in Windows PowerShell a single-cell Value2
    # write followed by an array write on the same sheet throws InvalidCastException.
    # A:I as before; J = the sequence's Explanation (TL Notes tab); K:L are filled in by Add-SequenceNotesExceptionColumns when ASSM exceptions exist
    $data = New-Object "object[,]" ($n + 1), 14
    $headers = @("Key", "Transmittal", "Sequence", "Rev", "Revises", "Due Date", "Paint / Notes", "Label Notes", "Label Due", "Explanation")
    for ($c = 0; $c -lt $headers.Count; $c++) { $data[0, $c] = $headers[$c] }
    $data[0, 12] = "Label Paint"
    $data[0, 13] = "Edit the yellow cells (Due Date, Paint / Notes, Explanation). Master BOM, Folder Labels, TL Labels and TERRA Work Order follow. Blank paint prints the Settings default on work orders. ASSM-specific paint / due / explanation go on the Assembly Exceptions tab."
    for ($i = 0; $i -lt $n; $i++) {
        $r = $Rows[$i]; $x = $i + 2
        $data[($i + 1), 0] = "=B$x&`"|`"&C$x"
        $data[($i + 1), 1] = $r.TL
        $data[($i + 1), 2] = $r.Sequence
        $data[($i + 1), 3] = if ($r.IsRev) { "REV" } else { "" }
        $data[($i + 1), 4] = $r.Revises
        $data[($i + 1), 5] = $r.DueDate
        $data[($i + 1), 6] = $r.Paint
        $data[($i + 1), 7] = "=IF(E$x<>`"`",`"REVISES `"&E$x&IF(M$x<>`"`",`" - `"&M$x,`"`"),M$x)"
        $data[($i + 1), 12] = "=G$x"       # M: the paint / exception text WITHOUT the REVISES part (the folder label shows REVISES on its own line)
        $data[($i + 1), 8] = "=IF(F$x=`"`",`"DUE    /    /`",`"DUE `"&F$x)"
        $data[($i + 1), 9] = $r.Explanation
    }
    $sheet.Range("A1").Resize($n + 1, 14).Value2 = $data
    foreach ($nm in @(@("Seq_Key", "A"), @("Seq_Due", "F"), @("Seq_Paint", "G"), @("Seq_LabelNotes", "H"), @("Seq_LabelDue", "I"), @("Seq_Expl", "J"), @("Seq_Revises", "E"), @("Seq_LabelPaint", "M"))) {
        [void]$Workbook.Names.Add($nm[0], "='Sequence Notes'!`$$($nm[1]):`$$($nm[1])")
    }
    if ($n -gt 0) {
        $sheet.Range("F2:G$($n + 1)").Interior.Color = Get-OleColor 255 242 204     # inputs
        $sheet.Range("J2:J$($n + 1)").Interior.Color = Get-OleColor 255 242 204
        $sheet.Range("A2:A$($n + 1)").Font.Color = Get-OleColor 128 128 128
        $sheet.Range("H2:I$($n + 1)").Font.Color = Get-OleColor 128 128 128
    }
    $hdr = $sheet.Range("A1:J1")
    $hdr.Font.Bold = $true
    $hdr.Font.Color = Get-OleColor 255 255 255
    $hdr.Interior.Color = Get-OleColor 27 54 93
    $hdr = $sheet.Range("M1")
    $hdr.Font.Bold = $true
    $hdr.Font.Color = Get-OleColor 255 255 255
    $hdr.Interior.Color = Get-OleColor 27 54 93
    $sheet.Range("M2:M$($n + 1)").Font.Color = Get-OleColor 128 128 128
    $sheet.Columns.Item(13).ColumnWidth = [double]22
    $sheet.Range("N1").Font.Italic = $true
    $widths = [double[]]@(16, 14, 10, 7, 22, 14, 44, 44, 16, 30)   # doubles: the COM binder caches the first ColumnWidth argument type and later double writes would fail
    for ($c = 0; $c -lt $widths.Count; $c++) { $sheet.Columns.Item($c + 1).ColumnWidth = $widths[$c] }
    [void]$sheet.Range("A1:J$($n + 1)").AutoFilter()
    $sheet.Activate()
    $Workbook.Application.ActiveWindow.SplitRow = 1
    $Workbook.Application.ActiveWindow.FreezePanes = $true
    return $sheet
}

# Writes a 2-D array to a sheet starting at A1 in one go. If Excel refuses the bulk write (it has thrown "Insufficient memory"
# for formula arrays), falls back to cell by cell so a single bad cell is named in the log and skipped instead of stopping the run.
function Write-SheetArray {
    param($Sheet, $Data, [int]$RowCount, [int]$ColCount)
    try {
        $Sheet.Range("A1").Resize($RowCount, $ColCount).Value2 = $Data
        return
    }
    catch {
        Write-Log "Bulk write to '$($Sheet.Name)' failed ($($_.Exception.Message)); writing cell by cell." "Yellow"
    }
    $bad = 0
    for ($r = 0; $r -lt $RowCount; $r++) {
        for ($c = 0; $c -lt $ColCount; $c++) {
            $v = $Data[$r, $c]
            if ($null -eq $v -or "$v" -eq "") { continue }
            try {
                $cell = $Sheet.Cells.Item($r + 1, $c + 1)
                if ("$v".StartsWith("=")) { $cell.Formula = "$v" } else { $cell.Value2 = "$v" }
            }
            catch {
                $bad++
                if ($bad -le 5) { Write-Log "  Could not write row $($r + 1), column $($c + 1) ('$("$v".Substring(0, [Math]::Min(60, "$v".Length)))'): $($_.Exception.Message)" "Yellow" }
            }
        }
    }
    if ($bad -gt 0) { Write-Log "  $bad cell(s) on '$($Sheet.Name)' could not be written." "Yellow" }
}

# "Assembly Exceptions" tab: ASSM-specific overrides of the sequence's paint / explanation / due date, read from the Required
# Assm workbook's Exceptions tab. One row per exception ASSM that is in this run (Rows: Assembly, TL, Sequence, Paint,
# Explanation, DueDate = the overrides), then the ones that are not (NotInRun: Assembly, Paint, Explanation, DueDate, Why).
# Defaults and effective values are formulas (sequence default from Sequence Notes; override else default), so editing an
# override (yellow) updates the Master BOM, work orders, labels and statistics. Defines Exc_Key / Exc_FolderKey / Exc_Paint /
# Exc_Expl / Exc_Due / Exc_Status. Must be built after the Sequence Notes tab.
function Add-AssemblyExceptionsSheet {
    param($Workbook, $Rows, $NotInRun)

    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "Assembly Exceptions"
    $sheet.Cells.Font.Name = "Arial"
    $rows = @($Rows); $nIn = $rows.Count
    $outs = @($NotInRun); $nOut = $outs.Count
    $total = $nIn + $nOut
    $sheet.Range("A2:C$($total + 2)").NumberFormat = "@"       # not D: the Folder Key column holds a formula, which a Text cell would keep as text
    $sheet.Range("H2:J$($total + 2)").NumberFormat = "@"      # typed dates stay text, same as Sequence Notes
    # One array write for everything (see Add-SequenceNotesSheet)
    $data = New-Object "object[,]" ($total + 1), 16
    $headers = @("Assembly", "Transmittal", "Sequence", "Folder Key", "Default Paint / Notes", "Default Explanation", "Default Due Date",
                 "Override Paint / Notes", "Override Explanation", "Override Due Date", "Effective Paint / Notes", "Effective Explanation", "Effective Due Date", "Status")
    for ($c = 0; $c -lt $headers.Count; $c++) { $data[0, $c] = $headers[$c] }
    $data[0, 15] = "Overrides come from the Required Assm Exceptions tab (edit the yellow cells here to change them). A blank override uses the sequence default from the Sequence Notes tab. Master BOM, TERRA Work Order and Folder Labels follow."
    for ($i = 0; $i -lt $nIn; $i++) {
        $r = $rows[$i]; $x = $i + 2
        $data[($i + 1), 0] = $r.Assembly
        $data[($i + 1), 1] = $r.TL
        $data[($i + 1), 2] = $r.Sequence
        $data[($i + 1), 3] = "=B$x&`"|`"&C$x"
        $data[($i + 1), 4] = "=" + (Get-SeqNoteRef "G" "D$x")
        $data[($i + 1), 5] = "=" + (Get-SeqNoteRef "J" "D$x")
        $data[($i + 1), 6] = "=" + (Get-SeqNoteRef "F" "D$x")
        if ($r.Paint) { $data[($i + 1), 7] = $r.Paint }
        if ($r.Explanation) { $data[($i + 1), 8] = $r.Explanation }
        if ($r.DueDate) { $data[($i + 1), 9] = $r.DueDate }
        $data[($i + 1), 10] = "=IF(H$x<>`"`",H$x,E$x)"
        $data[($i + 1), 11] = "=IF(I$x<>`"`",I$x,F$x)"
        $data[($i + 1), 12] = "=IF(J$x<>`"`",J$x,G$x)"
        $data[($i + 1), 13] = "=IF(COUNTIF(H${x}:J$x,`"?*`")>0,`"Exception set`",`"Uses default`")"
    }
    for ($j = 0; $j -lt $nOut; $j++) {
        $o = $outs[$j]; $row = $nIn + 1 + $j
        $data[$row, 0] = $o.Assembly
        if ($o.Paint) { $data[$row, 7] = $o.Paint }
        if ($o.Explanation) { $data[$row, 8] = $o.Explanation }
        if ($o.DueDate) { $data[$row, 9] = $o.DueDate }
        $data[$row, 13] = "Not in this run: $($o.Why)"
    }
    Write-SheetArray -Sheet $sheet -Data $data -RowCount ($total + 1) -ColCount 16
    foreach ($nm in @(@("Exc_Key", "A"), @("Exc_FolderKey", "D"), @("Exc_Paint", "K"), @("Exc_Expl", "L"), @("Exc_Due", "M"), @("Exc_Status", "N"))) {
        [void]$Workbook.Names.Add($nm[0], "='Assembly Exceptions'!`$$($nm[1]):`$$($nm[1])")
    }
    if ($total -gt 0) {
        $sheet.Range("H2:J$($total + 1)").Interior.Color = Get-OleColor 255 242 204      # inputs
        $sheet.Range("D2:G$($total + 1)").Font.Color = Get-OleColor 128 128 128
        $sheet.Range("K2:M$($total + 1)").Font.Bold = $true
    }
    $hdr = $sheet.Range("A1:N1")
    $hdr.Font.Bold = $true
    $hdr.Font.Color = Get-OleColor 255 255 255
    $hdr.Interior.Color = Get-OleColor 27 54 93
    $sheet.Range("P1").Font.Italic = $true
    $widths = [double[]]@(12, 12, 10, 16, 22, 22, 14, 22, 22, 14, 22, 22, 14, 30)
    for ($c = 0; $c -lt $widths.Count; $c++) { $sheet.Columns.Item($c + 1).ColumnWidth = $widths[$c] }
    [void]$sheet.Range("A1:N$($total + 1)").AutoFilter()
    $sheet.Activate()
    $Workbook.Application.ActiveWindow.SplitRow = 1
    $Workbook.Application.ActiveWindow.FreezePanes = $true
    return $sheet
}

# Sequence Notes columns K:L and the label notes (H) once the Assembly Exceptions tab exists: how many exception ASSMs sit in the
# folder and what the first one's effective paint says, appended to the Label Notes ("GREY PAINT - 12 ASSM NO PAINT").
function Add-SequenceNotesExceptionColumns {
    param($Workbook, [int]$RowCount)

    $sheet = $Workbook.Sheets.Item("Sequence Notes")
    $last = $RowCount + 1
    if ($RowCount -lt 1) { return }
    $f = New-Object "object[,]" $RowCount, 4
    for ($i = 0; $i -lt $RowCount; $i++) {
        $x = $i + 2
        # M = paint plus the exception text; H = label notes = REVISES (if any) plus M
        $paintText = "G$x&IF(K$x>0,IF(G$x<>`"`",`" - `",`"`")&K$x&`" ASSM `"&IF(L$x<>`"`",L$x,`"EXCEPTION`"),`"`")"
        $f[$i, 0] = "=IF(E$x<>`"`",`"REVISES `"&E$x&IF(M$x<>`"`",`" - `"&M$x,`"`"),M$x)"
        $f[$i, 3] = "=$paintText"
        $f[$i, 1] = "=COUNTIFS(Exc_FolderKey,A$x,Exc_Status,`"Exception set`")"
        $f[$i, 2] = "=IFERROR(INDEX(Exc_Paint,MATCH(A$x,Exc_FolderKey,0))&`"`",`"`")"
    }
    # These columns must be General: a Text-formatted cell keeps a formula as plain text, and the folder label would then print the formula itself
    $sheet.Range("H2:M$last").NumberFormat = "General"
    # H is the label notes (rewritten with the exception text), K:L are new
    for ($i = 0; $i -lt $RowCount; $i++) { $sheet.Cells.Item($i + 2, 8).Formula = $f[$i, 0] }
    $sheet.Cells.Item(1, 11).Value2 = "Exception ASSMs"
    $sheet.Cells.Item(1, 12).Value2 = "Exception Paint"
    for ($i = 0; $i -lt $RowCount; $i++) {
        $sheet.Cells.Item($i + 2, 11).Formula = $f[$i, 1]
        $sheet.Cells.Item($i + 2, 12).Formula = $f[$i, 2]
        $sheet.Cells.Item($i + 2, 13).Formula = $f[$i, 3]
    }
    # Check: any cell that still holds a formula as text is written again as a real formula
    $fix = @(@(8, 0), @(11, 1), @(12, 2), @(13, 3))
    for ($i = 0; $i -lt $RowCount; $i++) {
        foreach ($c in $fix) {
            $cell = $sheet.Cells.Item($i + 2, $c[0])
            if (-not $cell.HasFormula) { $cell.NumberFormat = "General"; $cell.Formula = $f[$i, $c[1]] }
        }
    }
    $hdr = $sheet.Range("K1:L1")
    $hdr.Font.Bold = $true
    $hdr.Font.Color = Get-OleColor 255 255 255
    $hdr.Interior.Color = Get-OleColor 27 54 93
    $sheet.Range("K2:L$last").Font.Color = Get-OleColor 128 128 128
    $sheet.Columns.Item(11).ColumnWidth = [double]16
    $sheet.Columns.Item(12).ColumnWidth = [double]22
}

# "37'-7 13/16" from a KISS length in millimeters (nearest 1/16")
function Format-FeetInches16 ($lengthStr) {
    $mm = 0.0
    if (-not [double]::TryParse(("$lengthStr").Trim(), [ref]$mm)) { return "$lengthStr" }
    $sixteenths = [int][Math]::Round($mm / 25.4 * 16, [MidpointRounding]::AwayFromZero)
    $feet = [Math]::Floor($sixteenths / 192)
    $rem = $sixteenths - $feet * 192
    $inches = [Math]::Floor($rem / 16)
    $frac = $rem - $inches * 16
    $out = "$feet'-$inches"
    if ($frac -gt 0) {
        $g = $frac; $d = 16
        while ($g % 2 -eq 0 -and $d % 2 -eq 0 -and $g -gt 0) { $g /= 2; $d /= 2 }
        $out += " $g/$d"
    }
    return $out
}

# One row per (source TL, assembly): the assembly-level facts shared by the work orders and the Master BOM columns.
# Part qtys in a KISS D block are totals for the assembly qty. An assembly listed in a REV transmittal (Required Assm)
# is placed on that REV (TlNumber) instead of its original TL.
function Get-AssemblyRows {
    param($MasterData, $RevAssemblies = $null)

    $rows = foreach ($g in ($MasterData | Group-Object -Property TL, Assembly)) {
        $parts = @($g.Group)
        $main = @($parts | Where-Object { $_.PartMark -eq $_.MainMark -and $_.Assembly -eq $_.MainMark })[0]
        if (-not $main) { $main = $parts[0] }
        $asmQty = [int]$main.Qty; if ($asmQty -lt 1) { $asmQty = 1 }
        $wt = 0.0
        foreach ($p in $parts) { $wt += Get-KssPartWeight $p.Shape $p.Size $p.Length $p.Qty }
        $size = "$($main.Size)"
        $desc = if ($size -match '^[A-Za-z]') { $size } else { "$($main.Shape)$size" }
        # "HSS16X8X1/2" -> "HSS16 x 8 x 1/2", "SHT10GAX5-1/2" -> "SHT10ga x 5 1/2"
        $desc = $desc -replace '(?<=[\d/]|GA)[xX](?=[\d.])', ' x '
        $desc = ($desc -replace '(\d)-(\d+/\d+)', '$1 $2') -creplace '(?<=\d)GA', 'ga'
        $assm = "$($main.Assembly)"
        $seq = if ($assm.Length -ge 4) { $assm.Substring(0, 4) } else { $assm }
        $tl = "$($main.TL)"
        [PSCustomObject]@{
            TL = $tl; Sequence = $seq; Assembly = $assm; Qty = $asmQty
            Description = $desc; Remark = "$($main.Description)"; Length = Format-FeetInches16 $main.Length
            WeightTotal = $wt; WeightEach = $wt / $asmQty
        }
    }
    foreach ($row in $rows) {
        $own = [int]($row.TL -replace '\D', '0')
        $tlNumber = if ($RevAssemblies -and $RevAssemblies.ContainsKey($row.Assembly)) { [int]$RevAssemblies[$row.Assembly] } else { $own }
        $row | Add-Member -NotePropertyName TlNumber -NotePropertyValue $tlNumber -PassThru -Force
    }
}

# One work order per transmittal + sequence (a folder per TL, a folder per sequence inside it).
# Revises = "TlNumber|Sequence" -> text of the transmittals a REV replaces (blank for non-REV).
function Get-WorkOrderPages {
    param($AssemblyRows, $SeqNotes, $TlNames = $null, $DueDates = $null, $Revises = $null)

    foreach ($g in ($AssemblyRows | Group-Object -Property TlNumber, Sequence)) {
        $first = $g.Group[0]
        $key = "$($first.TlNumber)|$($first.Sequence)"
        $items = @($g.Group | Group-Object -Property Assembly | ForEach-Object { $_.Group | Select-Object -Last 1 } | Sort-Object Assembly)
        $note = if ($SeqNotes) { $SeqNotes[$first.Sequence] } else { $null }
        $tlName = if ($TlNames -and $TlNames.ContainsKey($first.TlNumber)) { $TlNames[$first.TlNumber] } else { "TL$($first.TlNumber)" }
        [PSCustomObject]@{
            Sequence = $first.Sequence
            TL = $tlName
            TlNumber = $first.TlNumber
            IsRev = ($tlName -like 'TLR*')
            Revises = if ($Revises -and $Revises.ContainsKey($key)) { $Revises[$key] } else { "" }
            DueDate = if ($DueDates -and $DueDates.ContainsKey($key)) { $DueDates[$key] } else { "" }
            Paint = if ($note -and $note.Coating) { $note.Coating } else { "" }
            Explanation = if ($note -and $note.Explanation) { $note.Explanation } else { "" }
            Items = $items
        }
    }
}
# Splits the Exceptions-tab overrides into the ASSMs that are in this run (with the transmittal they are placed on and their
# sequence) and those that are not (with the reason). One row per ASSM number.
function Get-ExceptionRows {
    param($ExceptionMap, $AsmRows, $TlNames, $WorkbookTLs)

    $inRun = [System.Collections.Generic.List[object]]::new()
    $notInRun = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($row in @($AsmRows | Sort-Object { [int]$_.TlNumber }, Sequence, Assembly)) {
        $digits = "$($row.Assembly)" -replace '\D', ''
        if (-not $ExceptionMap -or -not $ExceptionMap.ContainsKey($digits) -or $seen.ContainsKey($digits)) { continue }
        $seen[$digits] = $true
        $ex = $ExceptionMap[$digits]
        $tlName = if ($TlNames -and $TlNames.ContainsKey([int]$row.TlNumber)) { $TlNames[[int]$row.TlNumber] } else { "TL$($row.TlNumber)" }
        $inRun.Add([PSCustomObject]@{ Assembly = $digits; TL = $tlName; Sequence = $row.Sequence; Paint = $ex.Paint; Explanation = $ex.Explanation; DueDate = $ex.DueDate })
    }
    if ($ExceptionMap) {
        foreach ($k in @($ExceptionMap.Keys | Sort-Object)) {
            if ($seen.ContainsKey($k)) { continue }
            $ex = $ExceptionMap[$k]
            $inSheet = @($WorkbookTLs | Where-Object { $_.SeqAssemblies.Values | Where-Object { $_.Contains($k) } }).Count -gt 0
            $why = if ($inSheet) { "its transmittal's .kss is not in the selected files" } else { "not listed in Sheet1 of the Required Assm workbook" }
            $notInRun.Add([PSCustomObject]@{ Assembly = $k; Paint = $ex.Paint; Explanation = $ex.Explanation; DueDate = $ex.DueDate; Why = $why })
        }
    }
    return [PSCustomObject]@{ InRun = @($inRun); NotInRun = @($notInRun) }
}

function Add-WorkOrderSheet {
    param($Workbook, $Pages, [string]$LogoBase64)   # job, customer job, PMs, issued date and CWI come from the Settings tab (Set_*)

    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "TERRA Work Order"
    $sheet.Cells.NumberFormat = "@"
    $sheet.Cells.Font.Name = "Arial"
    $sheet.Cells.Font.Size = 9.5

    # 17 columns cut at every x position used by the report's two header grids and the item grid (points, left margin = 0)
    $edges = @(0, 28.8, 50.4, 100.8, 108, 118.4, 194.4, 201.6, 234, 280.8, 327.6, 345.6, 378, 381.6, 429.8, 453.6, 481.9, 554.4)
    $edges = @(0, 28.8, 50.4, 100.8, 108, 118.4, 194.4, 201.6, 234, 280.8, 327.6, 345.6, 378, 381.6, 429.8, 453.6, 482, 554.4)
    for ($c = 1; $c -le 17; $c++) {
        $target = $edges[$c] - $edges[$c - 1]
        $col = $sheet.Columns.Item($c)
        $w = $target / 5.4
        for ($k = 0; $k -lt 10; $k++) {
            $col.ColumnWidth = $w
            $cur = [double]$col.Width
            if ([Math]::Abs($cur - $target) -lt 0.4) { break }
            $w = $w * ($target / $cur)
        }
    }

    $dark = 38 + 38 * 256 + 38 * 65536    # RGB(38,38,38)
    $white = 16777215
    # Alignment: xlCenter -4108, xlLeft -4131, xlRight -4152. Borders: xlThick 4, xlMedium -4138, xlContinuous 1, xlDash -4115.
    $box = {
        param($rng, $thickness)
        foreach ($edge in 7, 8, 9, 10) { $bd = $rng.Borders.Item($edge); $bd.LineStyle = 1; $bd.Weight = $thickness; $bd.Color = 0 }
    }
    $labelCell = {
        param($r, $c1, $c2, $text)
        $rng = $sheet.Range($sheet.Cells.Item($r, $c1), $sheet.Cells.Item($r, $c2))
        $rng.Merge()
        $rng.Value2 = $text
        $rng.Interior.Color = $dark
        $rng.Font.Color = $white
        $rng.Font.Underline = 2
        $rng.HorizontalAlignment = -4108
        $rng.VerticalAlignment = -4108
        & $box $rng 4
    }
    $valueCell = {
        param($r, $c1, $c2, $text, $fmt = "General")
        $rng = $sheet.Range($sheet.Cells.Item($r, $c1), $sheet.Cells.Item($r, $c2))
        $rng.Merge()
        if ("$text".StartsWith("=")) { Set-CellFormulaText $rng $text $fmt } else { $rng.Value2 = $text }
        $rng.HorizontalAlignment = -4108
        $rng.VerticalAlignment = -4108
        $rng.WrapText = $true
        & $box $rng 4
    }

    $logoPath = $null
    if ($LogoBase64) {
        $logoPath = Join-Path ([System.IO.Path]::GetTempPath()) "terra_wo_logo.jpg"
        [System.IO.File]::WriteAllBytes($logoPath, [Convert]::FromBase64String($LogoBase64))
    }

    $createdText = "Report Created: $((Get-Date).ToString('MM/dd/yyyy h:mm tt', [System.Globalization.CultureInfo]::InvariantCulture))"
    $firstCap = 16    # items that fit under the full header; continuation pages (table header only) fit 23
    $contCap = 23
    $grey = 242 + 242 * 256 + 242 * 65536
    $row = 1
    $pageStarts = @()

    foreach ($wo in $Pages) {
        # Split the work order's items into pages
        $chunkList = [System.Collections.Generic.List[object]]::new()
        $queue = @($wo.Items)
        $cap = $firstCap
        do {
            $take = [Math]::Min($cap, $queue.Count)
            $chunkList.Add(@($queue | Select-Object -First $take))
            $queue = @($queue | Select-Object -Skip $take)
            $cap = $contCap
        } while ($queue.Count -gt 0)
        $chunks = $chunkList.Count
        $itemNo = 0
        $sumRanges = @()

        for ($pg = 1; $pg -le $chunks; $pg++) {
            $start = $row
            $pageStarts += $start
            $pageItems = @($chunkList[$pg - 1])
            $isFirst = ($pg -eq 1)
            $isLast = ($pg -eq $chunks)

            # Row heights follow the report: 79.2 address, 14.4 / 28.8 header grids, 21.6 gap, then 14.4 table header, 7.2 gap
            $h = if ($isFirst) { @(79.2, 14.4, 28.8, 14.4, 28.8, 21.6, 14.4, 7.2) } else { @(14.4, 7.2) }
            for ($i = 0; $i -lt $h.Count; $i++) { $sheet.Rows.Item($start + $i).RowHeight = $h[$i] }

            if ($isFirst) {
                $addr = $sheet.Range($sheet.Cells.Item($start, 1), $sheet.Cells.Item($start, 10))
                $addr.Merge()
                $addr.Value2 = "ADVANTAGE TERRAFAB`n400 S. 45TH STREET EAST`nMUSKOGEE, OK 74403`nPHONE: (918) 681-4496"
                $addr.Font.Bold = $true
                $addr.Font.Size = 11.5
                $addr.WrapText = $true
                $addr.HorizontalAlignment = -4131
                $addr.VerticalAlignment = -4108
                $addr.IndentLevel = 1
                if ($logoPath) {
                    $top = $sheet.Cells.Item($start, 1).Top
                    [void]$sheet.Shapes.AddPicture($logoPath, 0, -1, 338.4, $top, 217.44, 72)
                }

                $r1 = $start + 1; $r2 = $start + 2; $r3 = $start + 3; $r4 = $start + 4
                & $labelCell $r1 1 4 "PM"
                & $labelCell $r1 5 7 "ATF JOB / WO #"
                & $labelCell $r1 8 11 "ISSUED TO SHOP DATE"
                & $labelCell $r1 12 15 "DATE DUE"
                & $labelCell $r1 16 17 "CWI REQ'D"
                & $valueCell $r2 1 4 "=Set_ProjectManagers"
                & $valueCell $r2 5 7 "=TEXTJOIN("" / "",TRUE,Set_OurJob,""$($wo.Sequence)"")"
                & $valueCell $r2 8 11 "=Set_IssuedDate" "m/d/yy"
                & $valueCell $r2 12 15 ""
                $dueCellWo = $sheet.Cells.Item($r2, 12)
                $dueCellWo.NumberFormat = "General"
                $dueRef = Get-SeqNoteRef "F" "`"$($wo.TL)|$($wo.Sequence)`""
                $dueCellWo.Formula = "=IF($dueRef=`"`",IFERROR(Set_DefaultDue,`"3 WKS AFTER MATERIAL ARRIVES`"),$dueRef)"
                & $valueCell $r2 16 17 "=Set_CWI"
                & $labelCell $r3 1 3 "DWGS APPROVED"
                & $labelCell $r3 4 6 "CUSTOMER JOB#"
                & $labelCell $r3 7 9 "SEQUENCE"
                & $labelCell $r3 10 13 "NOTES / REMARKS"
                & $labelCell $r3 14 16 "DATE COMPLETED"
                & $labelCell $r3 17 17 "QC DATE"
                & $valueCell $r4 1 3 ""
                & $valueCell $r4 4 6 "=Set_CustomerJob"
                & $valueCell $r4 7 9 $wo.Sequence
                & $valueCell $r4 10 13 $wo.TL
                & $valueCell $r4 14 16 ""
                & $valueCell $r4 17 17 ""
            }

            # Item table header
            $hr = $start + $h.Count - 2
            & $labelCell $hr 1 1 "ITEM"
            & $labelCell $hr 2 2 "Qty"
            & $labelCell $hr 3 5 "PART #"
            & $labelCell $hr 6 8 "DESCRIPTION"
            & $labelCell $hr 9 10 "REMARK"
            & $labelCell $hr 11 12 "LENGTH"
            & $labelCell $hr 13 14 "W# EA."
            & $labelCell $hr 15 16 "TOTAL W#"
            & $labelCell $hr 17 17 "NOTES"

            # Items: even rows bold on light grey, odd rows plain
            $first = $start + $h.Count
            $n = 0
            foreach ($it in $pageItems) {
                $r = $first + $n
                $itemNo++
                $sheet.Rows.Item($r).RowHeight = 28.8
                $spans = @(@(1, 1), @(2, 2), @(3, 5), @(6, 8), @(9, 10), @(11, 12), @(13, 14), @(15, 16), @(17, 17))
                foreach ($sp in $spans) {
                    $rng = $sheet.Range($sheet.Cells.Item($r, $sp[0]), $sheet.Cells.Item($r, $sp[1]))
                    $rng.Merge()
                    $rng.HorizontalAlignment = -4108
                    $rng.VerticalAlignment = -4108
                }
                $sheet.Cells.Item($r, 1).NumberFormat = "0"; $sheet.Cells.Item($r, 1).Value2 = $itemNo
                $sheet.Cells.Item($r, 2).NumberFormat = "0"; $sheet.Cells.Item($r, 2).Value2 = [int]$it.Qty
                $sheet.Cells.Item($r, 3).Value2 = $it.Assembly
                $sheet.Cells.Item($r, 6).Value2 = $it.Description
                $sheet.Cells.Item($r, 9).Value2 = $it.Remark
                $sheet.Cells.Item($r, 11).Value2 = $it.Length
                $sheet.Cells.Item($r, 15).NumberFormat = "#,##0"; $sheet.Cells.Item($r, 15).Value2 = [double]$it.WeightTotal
                $sheet.Cells.Item($r, 13).NumberFormat = "#,##0"; $sheet.Cells.Item($r, 13).Formula = "=O$r/B$r"
                $sheet.Cells.Item($r, 17).NumberFormat = "General"
                # paint/notes live on the Sequence Notes; an assembly with an exception shows its own paint (Assembly Exceptions tab)
                $sheet.Cells.Item($r, 17).Formula = Get-WorkOrderNoteFormula $wo.TL $wo.Sequence $it.Assembly
                $sheet.Cells.Item($r, 17).WrapText = $true
                if ($itemNo % 2 -eq 0) {
                    $line = $sheet.Range($sheet.Cells.Item($r, 1), $sheet.Cells.Item($r, 17))
                    $line.Interior.Color = $grey
                    $line.Font.Bold = $true
                }
                # dotted divider between W# EA. and TOTAL W#
                $bd = $sheet.Cells.Item($r, 15).Borders.Item(7)
                $bd.LineStyle = -4118; $bd.Weight = 2; $bd.Color = 0
                $n++
            }
            if ($pageItems.Count -gt 0) { $sumRanges += "O$($first):O$($first + $pageItems.Count - 1)" }

            # Total row (after the last item of the work order), then the closing rule
            $tr = $first + $pageItems.Count
            $sheet.Rows.Item($tr).RowHeight = 21.6
            if ($isLast) {
                $tl = $sheet.Range($sheet.Cells.Item($tr, 11), $sheet.Cells.Item($tr, 12))
                $tl.Merge(); $tl.Value2 = "Total:"; $tl.Font.Bold = $true
                $tl.HorizontalAlignment = -4152; $tl.VerticalAlignment = -4107
                $tv = $sheet.Range($sheet.Cells.Item($tr, 15), $sheet.Cells.Item($tr, 16))
                $tv.Merge(); $tv.NumberFormat = "#,##0"
                $tv.HorizontalAlignment = -4108; $tv.VerticalAlignment = -4107
                $sheet.Cells.Item($tr, 15).Formula = "=SUM($($sumRanges -join ','))"
            }
            if ($isLast) {
                $rule = $sheet.Range($sheet.Cells.Item($tr, 1), $sheet.Cells.Item($tr, 17)).Borders.Item(9)
                $rule.LineStyle = 1; $rule.Weight = -4138; $rule.Color = 0
            }

            # Filler so the footer lands at the bottom of the page, then the footer (rule above and below)
            $used = ($h | Measure-Object -Sum).Sum + 28.8 * $pageItems.Count + 21.6
            $ft = $tr + 1
            $remaining = [double](714.2 - $used)
            while ($remaining -gt 0.5) {               # Excel caps a row at 409.5 pt
                $sheet.Rows.Item($ft).RowHeight = [Math]::Min(400.0, $remaining)
                $remaining -= 400.0
                $ft++
            }
            $sheet.Rows.Item($ft).RowHeight = 14.4
            $fl = $sheet.Range($sheet.Cells.Item($ft, 1), $sheet.Cells.Item($ft, 8))
            $fl.Merge(); $fl.Value2 = "Page $pg of $chunks"
            $fr2 = $sheet.Range($sheet.Cells.Item($ft, 9), $sheet.Cells.Item($ft, 17))
            $fr2.Merge(); $fr2.Value2 = $createdText
            $foot = $sheet.Range($sheet.Cells.Item($ft, 1), $sheet.Cells.Item($ft, 17))
            $foot.Font.Name = "Microsoft Sans Serif"; $foot.Font.Size = 7
            $fl.HorizontalAlignment = -4131; $fr2.HorizontalAlignment = -4152
            $foot.VerticalAlignment = -4108
            foreach ($seg in $fl, $fr2) {
                foreach ($edge in 8, 9) { $bd = $seg.Borders.Item($edge); $bd.LineStyle = 1; $bd.Weight = -4138; $bd.Color = 0 }
            }

            $row = $ft + 1
        }
    }

    try {
        $ps = $sheet.PageSetup
        $ps.PrintArea = "`$A`$1:`$Q`$$($row - 1)"
        $ps.Orientation = 1
        $ps.PaperSize = 1                         # Letter
        $ps.LeftMargin = 28.08; $ps.RightMargin = 28.08
        $ps.TopMargin = 28.08; $ps.BottomMargin = 28.08
        $ps.HeaderMargin = 0; $ps.FooterMargin = 0
        $ps.CenterHorizontally = $true
        $sheet.ResetAllPageBreaks()
        foreach ($ps0 in ($pageStarts | Select-Object -Skip 1)) { [void]$sheet.HPageBreaks.Add($sheet.Cells.Item($ps0, 1)) }
        $ps.Zoom = $false
        $ps.FitToPagesWide = 1
        $ps.FitToPagesTall = $false
    }
    catch {
        Write-Log "TERRA Work Order tab built, but page setup failed (is a printer installed?): $($_.Exception.Message)" "Yellow"
    }
    return $sheet
}

$script:TerraLogoBase64 = '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'

# Blank Required Assm workbook (Sheet1, Table1 Transposed, Sequences in Each TL, Sequences by TL, TL Notes, Exceptions): all formulas,
# tables and formatting, no data. Written out by menu choice 3. To change it: write a copy, edit it in Excel, re-embed the saved .xlsx as base64.
$script:RequiredAssmTemplateBase64 = (@'
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
Shi0Z2EA+R57MXQxRDMXErlgwplhnM3EM8PEzmb+qkbUXCujCvS5EkcQCxCuaRj8YmjtWlByAOjnk+5KX81cxcpZRTDpCjnkPQGynCEbiMRRcYEsZ25l6YK5m7YUpiF/QW2fjfF6Hcsn
GHr2cqRfUEsDBBQAAAAIAA+MZF2ajMJMpWEAABweAwAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1srN3JkmTndWjpeT0FjBOORODvz6GRuuYe3vdtKCJmEAhKMBEEC4C6Kqt3r0RD
ME+s5ffqyEQTu49/RiZiYyNpy7ek3/2v//j6T5/825fffvfVN3/+/a/Dbz779Sdf/vmLb/7w1Z//6fe/vt8Wf9f9+pPvvv/8z3/4/E/f/PnL3//6P7/87tf/6+//r9/9+zff/st3//zl
l99/8uEL/Pm73//qn7///i+//fTT77745y+//vy733zzly///OE/+eM33379+fcf/u23//Tpd3/59svP//DjD/r6T5/Gzz6rn379+Vd//tVPX+G33/5XvsY3f/zjV198Ofvmi3/9+ss/
f//TF/n2yz99/v2HX/53//zVX77761f7+ov/ypf7+vNv/+Vf//J3X3zz9V8+fIl//OpPX33/nz9+0b9+mf8I+XN+pa+/+uLbb7775o/f/+bDj/z518Q/vP7T/tPPv/jlK3373/syIX/4
A/y3r36Yz9++VPxvfq3yy9eKf/ti6b/5xeovXyz96pOvv/jt+p/+/M23n//jn778/a9+/LZ98h/ffvif+OHvH/7z//j2t//61R9+/6v/97Of//Z3H/4e/u6v/+qnf/jr3/6/X/397/7w
1YcJ//DFP/n2yz/+/leT8NvJP6RUf/Xp3//ux1/D81df/vt3H/3rT77//B+vX/7pyy++//LDTxN+9cn/8803X1+/+PyHX0777KN/e/hh/n/6CX/4s/gfv/nmX374AusPP+yzDz/xdz9+
kR9+4s8//NO/ffn05Z8+vJ6l1P3qk+/+7x9/MT/+mw+/kE9/+dk//td//VUtfvwT7fTtJ3/48o+f/+ufvr988++rL7/6p3/+/sMvr3z4jvzwPfrtH/5z9uV3X3z4c/nDT/6bWH74ql98
86fvfvzHT77+6s8//qF8/fl//PjP//7VH77/59//KubftPBZn9qHr/LFv373/Tdf/8NP/0H4+Yf/9APjzz8w/fIDQ/lNaZ+lEP/3PzD//APz337G9JvYlVDq/+FHlp9/ZPnbT1n/az+y
/vwjY/gv/mo//enb9ON3evb595///e++/ebfP/n2x+/Td3/5/Ie/JIXf5v5Xn/zw/U7xN+GvX+SXGTwawYdf1Q9fZ/LDF/rwh/LhZ//wn3z3gf/t7+vvPv23H37qn59Mf3ny6c/yBJlB
5pAFZAlZQdaQDWQL2UH2kAPkCDlBzpAL5Aq5Qe4fy6cfpvrLaKOMNobf9GX0bCNnG/p3w40YLmQGmUMWkCVkBVlDNpAtZAfZQw6QI+QEOUMukCvkBrnHR8NN/2PDTRxu/OzdcBOGC5lB
5pAFZAlZQdaQDWQL2UH2kAPkCDlBzpAL5Aq5Qe7p0XDz/9hwswz3/V+WM4YLmUHmkAVkCVlB1pANZAvZQfaQA+QIOUHOkAvkCrlB7vnRcMv/2HCLDDe8G27BcCEzyByygCwhK8gasoFs
ITvIHnKAHCEnyBlygVwhN8i9PBpu/R8bbuVwU3o33IrhQmaQOWQBWUJWkDVkA9lCdpA95AA5Qk6QM+QCuUJukHt9NNz2Pzbc9uNPEdPHm9veDffnNx//9+k4fPL005OPv0p996fITJ7k
4ZO5PCnDJwt58u53kaU8efdHtJIn3fDJ+qcn+eMn7/6L5uanJx/vRXv3X1e28uTdXxd38uTdd3cvT959dw/y5N139yhP3n13T/Lk3Xf3LE/efXcv8uTdd/cqT959d2980r377t7lybvv
7vNPT8IvO/UP7+HlPby+h7f3MJlAppAnyAwyhywgS8gKsoZsIFvIDrKHHCBHyAlyhlwgV8gNcodgnpOPBzr4a2Rnf4387Dfh42D12fi60MlfMN/92T396U3922+GP0H72++F72H+Hhbv
YfkeVu9h/RN0f/td8D1sf4L+b78Hvof9ezi8h+N7OL2H83u4vIfre7i9h/t7eO7eL/J7eHkPr+/hrcMiQ6aQJ8gMMocsIEvICrKGbCBbyA6yhxwgR8gJcoZcIFfIDXKHYJ6Tjwc6WOT+
3SL/b/dz2r/fvf797r2H+XtYvIfle1i9h3X/fvfev9i+h9172L+Hw3s4vofTezi/h8t7uL6H23u4v4fn/v3uvYeX9/D6Ht567B5kCnmCzCBzyAKyhKwga8gGsoXsIHvIAXKEnCBnyAVy
hdwgdwjmOfl4oIPdC5+NWr4fnw+272f5aP0gc8gCsoSsIOuf5aMlxJstZAfZQw6QI+QEOUMukCvkBrlDnn+WjzYS8gJ5hbxBJhPSlPREmpHmpAVpSVqR1qQNaUvakfakA+lIOpHOpAvp
SrqR7iROejIY9XBZw7hlDVjWgGUNWNaAZQ1Y1oBlDVjWgGUNWNaAZQ1Y1oBlDVjWgGUNWNaAZQ1Y1oBlDVjWgGUNWNaAZQ1Y1oBlDVzWwGUNXNbAZQ1c1sBlDVzWwGUNXNbAZQ1c1sBl
DVzWwGUNXNbAZQ1c1sBlDVzWwGUNXNbAZQ1c1vBoWeO4ZY1Y1ohljVjWiGWNWNaIZY1Y1ohljVjWiGWNWNaIZY1Y1ohljVjWiGWNWNaIZY1Y1ohljVjWiGWNWNaIZY1c1shljVzWyGWN
XNbIZY1c1shljVzWyGWNXNbIZY1c1shljVzWyGWNXNbIZY1c1shljVzWyGWNXNb4aFnTuGVNWNaEZU1Y1oRlTVjWhGVNWNaEZU1Y1oRlTVjWhGVNWNaEZU1Y1oRlTVjWhGVNWNaEZU1Y
1oRlTVjWhGVNXNbEZU1c1sRlTVzWxGVNXNbEZU1c1sRlTVzWxGVNXNbEZU1c1sRlTVzWxGVNXNbEZU1c1sRlTVzW9GhZ87hlzVjWjGXNWNaMZc1Y1oxlzVjWjGXNWNaMZc1Y1oxlzVjW
jGXNWNaMZc1Y1oxlzVjWjGXNWNaMZc1Y1oxlzVzWzGXNXNbMZc1c1sxlzVzWzGXNXNbMZc1c1sxlzVzWzGXNXNbMZc1c1sxlzVzWzGXNXNbMZc1c1vxoWcu4ZS1Y1oJlLVjWgmUtWNaC
ZS1Y1oJlLVjWgmUtWNaCZS1Y1oJlLVjWgmUtWNaCZS1Y1oJlLVjWgmUtWNaCZS1c1sJlLVzWwmUtXNbCZS1c1sJlLVzWwmUtXNbCZS1c1sJlLVzWwmUtXNbCZS1c1sJlLVzWwmUtXNby
aFnruGWtWNaKZa1Y1oplrVjWimWtWNaKZa1Y1oplrVjWimWtWNaKZa1Y1oplrVjWimWtWNaKZa1Y1oplrVjWimWtXNbKZa1c1splrVzWymWtXNbKZa1c1splrVzWymWtXNbKZa1c1spl
rVzWymWtXNbKZa1c1splrVzW+mhZ/3YkFv8ry9qwrA3L2rCsDcvasKwNy9qwrA3L2rCsDcvasKwNy9qwrA3L2rCsDcvasKwNy9qwrA3LioskyAvkFfIWeJVEmpKeSDPSnLQgLUkr0pq0
IW1JO9KedCAdSSfSmXQhXUk30p3ESU8Gox4uazduWXFhFHBiBJlDFpAlZAVZBxwa4c0WsoPsIQfIEXKCnCEXyBVyg9whzwFXR5AXyCvkLfDyiDQlPZFmpDlpQVqSVqQ1aUPaknakPelA
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
UnHy44pTRlacIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlSccrjilNGVpwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadI
xSlScYpUnCIVp0jFKVJxyuOKU0ZWnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnHK44pTRlacIhWnSMUpUnGKVJwiFadIxSlScYpUnCIV
p0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlSccrjilNGVpwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxyuOKU0ZWnCIVp0jFKVJxilSc
IhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnHK44pTRlacIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlSccrjilNG
VpwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxyuOKU0ZWnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxilSc
IhWnSMUpUnHK44pTRlacIhWnSMUpUnGKVJwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJwiFadIxSlSccrjilNHVpwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnGq
VJwqFadKxalScapUnCoVp0rFqVJx6uOKU0dWnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnHq44pTR1acKhWnSsWpUnGqVJwqFadKxalS
capUnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalScerjilNHVpwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJx6uOKU0dWnCoVp0rF
qVJxqlScKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnHq44pTR1acKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalS
cerjilNHVpwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJx6uOKU0dWnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rF
qVJxqlScKhWnSsWpUnHq44pTR1acKhWnSsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJxqlScKhWnSsWpUnGqVJwqFadKxalScerjilNHVpwqFadKxalScapUnCoVp0rFqVJxqlScKhWn
SsWpUnGqVJwqFadKxalScapUnCoVp0rFqVJx6uOK00ZWnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnHa44rTRlacJhWnScVpUnGaVJwm
FadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScdrjitNGVpwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJx2uOK00ZW
nCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnHa44rTRlacJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnGaVJwm
FadJxWlScdrjitNGVpwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJx2uOK00ZWnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScZpU
nCYVp0nFaVJxmlScJhWnScVpUnHa44rTRlacJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScdrjitNGVpwmFadJxWlScZpUnCYVp0nFaVJx
mlScJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJx2uOK00ZWnCYVp0nFaVJxmlScJhWnScVpUnGaVJwmFadJxWlScZpUnCYVp0nFaVJxmlScJhWnScVpUnHa44rTjaw4nVScTipO
JxWnk4rTScXppOJ0UnE6qTidVJxOKk4nFaeTitNJxemk4nRScTqpOJ1UnE4qTicVp3tccbqRFaeTitNJxemk4nRScTqpOJ1UnE4qTicVp5OK00nF6aTidFJxOqk4nVScTipOJxWnk4rT
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
UnHy44qTR1acLBUnS8XJUnGyVJwsFSdLxclScbJUnCwVJ0vFyVJxslScLBUnS8XJUnGyVJwsFSdLxclScfLjipNHVpwsFSdLxclScbJUnCwVJ0vFyVJxslScLBUnS8XJUnGyVJwsFSdL
xclScbJUnCwVJ0vFyVJx8uOKk0dWnCwVJ0vFyVJxslScLBUnS8XJUnGyVJwsFSdLxclScbJUnCwVJ0vFyVJxslScLBUnS8XJUnHy44qTR1acLBUnS8XJUnGyVJwsFSdLxclScbJUnCwV
J0vFyVJxslScLBUnS8XJUnGyVJwsFSdLxclScfLjilNGVpwiFadIxSlScYpUnCIVp0jFKVJxilScIhWnSMUpUnGKVJzy/xdjt7t2Y8eBhm9FEIIgBgx0qtZZVbUUtwHbGc8YSWzHsS5A
SZ9uC5GlHvXpfNz9qOVxR0keYsI/k3+bdUguki/3xsEDxdlQnA3F2VCcDcXZUJwNxdnXirNvKs6G4mwozobibCjOhuJsKM6G4mwozobibCjOhuJsKM6G4mwozobibCjOhuJsKM6G4uxr
xdk3FWdDcTYUZ0NxNhRnQ3E2FGdDcTYUZ0NxNhRnQ3E2FGdDcTYUZ0NxNhRnQ3E2FGdDcfa14uybirOhOBuKs6E4G4qzoTgbirOhOBuKs6E4G4qzoTgbirOhOBuKs6E4G4qzoTgbirOh
OPtacfZNxdlQnA3F2VCcDcXZUJwNxdlQnA3F2VCcDcXZUJwNxdlQnA3F2VCcDcXZUJwNxdlQnH2tOPum4mwozobibCjOhuJsKM6G4mwozobibCjOhuJsKM6G4mwozobibCjOhuJsKM6G
4mwozr5WnH1TcTYUZ0NxNhRnQ3E2FGdDcTYUZ0NxNhRnQ3E2FGdDcTYUZ0NxNhRnQ3E2FGdDcTYUZ18rzr6pOBuKs6E4G4qzoTgbirOhOBuKs6E4G4qzoTgbirOhOBuKs6E4G4qzoTgb
irOhOBuKs68VZ99UnA3F2VCcDcXZUJwNxdlQnA3F2VCcDcXZUJwNxdlQnA3F2VCcDcXZUJwNxdlQnA3F2deKs28qzobibCjOhuJsKM6G4mwozobibCjOhuJsKM6G4mwozobibCjOhuJs
KM6G4mwozobi7GvFqZuKU1CcguIUFKegOAXFKShOQXEKilNQnILiFBSnoDgFxSkoTkFxCopTUJyC4hQUp64Vp24qTkFxCopTUJyC4hQUp6A4BcUpKE5BcQqKU1CcguIUFKegOAXFKShO
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
KM6C4iwozoLiLCjOguIsKM6C4iwozoLiLCjOulacdVNxFhRnQXEWFGdBcRYUZ0FxFhRnQXEWFGdBcRYUZ0FxFhRnQXEWFGdBcRYUZ0FxFhRn/WfF+eyb3z0+Pv3lq6dXP/7R7x/ff/X4
s8c3b7559g/vvv0uXT3/ZPrs/eOXnz//yX7x8mO7/zivFy9L83jxMjTPFy9T8/Xi5dL84cXLh4/X/W+X+eMfff27d28fn17/w6/fP/vy3dunX3zx4e6eP3v6168fP3/+9t3P3r39p8f3
37x+9/bjQ3j39ovXTx82Xr35+bv3v3/19PT67VfPvvnfH0//037xsr97S7/8zbdvHv/vGR7/5ev3j998PP7ZF//y5cez9/NnX79//e7966d//fx5fDjiyw/n+vbNqx//6jd/9tf/4+e/
/bOf/kn/sH7w4bV+9+bb3799/sOf/erlL3/7kw/jefHTP1nnB5//+Q9+9NkfD/rRZ39Y8L+wcH26cH6y8E9++ZffLfqnb57+4k+/evqL589/+Ml1fD+9upzv/ugr+owP7P/1HOfFx/fs
v3BD+9MbWv8fn+TDpws//Dc+ya9fffX4N6/ef/X67TfP3jx++fG38sP79f71V7/74+end19//PThYf39u6end7//49bvHl998fj+u60P/yl9+e7d0x83PvvDef/u8enbr599uMkPv8Gv
vlv68+dfv3v/9P7V66fvdnl69fdvHn/96v3T99/1797l76fP3r94/eFRvf/FFx+/o5/92+4fNv753ft//Piz8eP/A1BLAwQUAAAACAAPjGRdRi09IPNDAADA9AEAGAAAAHhsL3dvcmtz
aGVldHMvc2hlZXQyLnhtbK29+69kd51l+fv8FVaOVBQtwI6MiPuggNZ2nPc58X7HaDTKstNgle10ZyYFaDT/+9y002ZHrM1M3RJSPw4LO23W2nwarr6d9bv/+ddvv/nkP1+/fff1m+9+
/4vRbz77xSevv/vizZdff/fH3/9iv6t+/fCLT969f/Xdl6++efPd69//4m+v3/3if/7hf/vdX968/Y93f3r9+v0nT7/Ad+9+/+JP799//9tPP333xZ9ef/vq3W/efP/6u6d/5qs3b799
9f7pH77946fvvn/7+tWXP/xJ337z6cvPPrv79NtXX3/34sdf4bdv/yu/xpuvvvr6i9fFmy/+/O3r797/+Iu8ff3Nq/dPf/vv/vT19+9++tW+/eK/8st9++rtf/z5+19/8ebb759+iX//
+puv3//th1/0p1/mr6PJK/5K3379xds379589f43T3/mx78n/st7/PTx01df/Pwrvf3v/TKjydO/wP/8+kOfv/9SL/+bv9b051/r5d9/sfF/8xe7+/kXG7/45Nsvftv+8bs3b1/9+zev
f//iB22f/PXt0/96+fS/n/75v7797Z+//vL3L/7vu/LzYjS5n/16Nh69/PVkUt39+mE6nvxaL+8fp5/dTatpeff/vPjD7778+qnwh1/8k7evv/r9C41+O2xGn3324tM//O6Hv4fD16//
8s6+P/mwyH9/8+Y/PvyD9umv9NnTL/Lu9Tevv/iwjU9ePf1f//l69vqbb37/on759K/+3f/64df98P30S37686/j3z/9+tUPk1m9/eTL11+9+vM37zdv/tK8/vqPf3r/+xej6dO/tg//
an/75d+K1+++eFrl01/6Ny+nH37VL9588+6H//OTb7/+7umPfdL06q8//N9/+frL9396+pr8ZvJyev8wevrjP/niz+/ev/n2+PGf+fjn//hnvvz4Z05f/v1PHf/mfvTZ4/g+/pmf/viX
/uHvvnj1/tUffvf2zV8+efvDX/vd968+/Bt29Nvx+LMXn3z4FzF++ZvRT7/Kz//C4r+uyZPULz78OvrwCz39DT39O+Xpn3j3hP/zDy9/9+l/fvhLf/xDPv/5D/n0I5mBFCAlSAVSgzQg
LUgH0v9IRuOfyQAyB1mALEFWIGuQDcgWZAeyBzmAHEFOIGeQC4hE9DnRjKggKokqoproY9vR5O+oJeqIeqKBaE60IFoSrYjWRBuiLdGOaE90IDoSnYjORBdHnz5dh59PxMt0Il6OfvM4
/a/diKdb9tNf5OWP/467tyMxvjkSP/8hPx8JkAKkBKlAapAGpAXpQPofyWj69yMBMgdZgCxBViBrkA3IFmQHsgc5gBxBTiBnkAuIRPQ50YyoICqJKqKa6GPbqyNB1BH1RAPRnGhBtCRa
Ea2JNkRboh3RnuhAdCQ6EZ2JLi//0ZEY3x6J/+9DMP7x1/n7/9P7OcgMpAApQSqQGqQBaUE6kB5kAJmDLECWICuQNcgGZAuyA9mDHECOICeQM8gFRIwqVhWzil3FsGJZMa1+auv/vifq
iHqigWhOtCBaEq2I1kQboi3RjmhPdCA6Ep2IzkSX8T/69/3kn/cfDiY//jWe/sPBF99++MP/8LuvPvzHhFdv377624uP/61v8vTf+h5f/KGtys1mufnX/+uv33z13W+Gcvfh6/tvf/On
X+0+/BfO0f/xvzdP/9309dt3/+evfvwnvvz5n/jw335+on9+9/pX/zqU1d//9Je/fPq30fDil//j4y/9+Xm2HP71p1/j419O888LfYRf/Gq23C92P//DX/7yl//yx/f/9tkvP/4lvvrT
xz/r6f/8y7vfVO2wKzc//+V+/tv4+Q//8h//4V/yD//rr3YbLbar5bb8+Dd82O406//V/+I//LK//OWv2uojfvovji9+9fS/Pv6jp3/qxYtf/u7Tr/w/hv0Yw/5T2C0obkF5C6pbUN+C
5ha0t6C7Bf0tGG7B/BYsbsHyFqxuwfoWbG7B9hbsbsH+FhxuwfEWnG7B+RZcboEEgnZCPKGekE/oJwTUx4KjBzu3RB1RTzQQzYkWREuiFdGaaEO0JdoR7YkOREeiE9GZ6AL0uf5+jT//
3L5n9l3Yd2nflX3X9t3Yd2vfnX339j3Y99y+F/a9tO+Vfa/te2PfW/ve2ffevg/2fbTvk32f7fvy9++ZeZuZt5l5m5m3mXmbmbeZeZuZt5l5m5m3mXmbmbeZeZuZt5l5m5m3mXmbmbeZ
eZuZt5l5m5m3mXmbmbeZeZuZt8K8FeatMG+FeSvMW2HeCvNWmLfCvBXmrTBvhXkrzFth3grzVpi3wrwV5q0wb4V5K8xbYd4K81aYt8K8FeatNG+leSvNW2neSvNWmrfSvJXmrTRvpXkr
zVtp3krzVpq30ryV5q00b6V5K81bad5K81aat9K8leatNG+leavMW2XeKvNWmbfKvFXmrTJvlXmrzFtl3irzVpm3yrxV5q0yb5V5q8xbZd4q81aZt8q8VeatMm+VeavMW2XeavNWm7fa
vNXmrTZvtXmrzVtt3mrzVpu32rzV5q02b7V5q81bbd5q81abt9q81eatNm+1eavNW23eavNWm7fGvDXmrTFvjXlrzFtj3hrz1pi3xrw15q0xb415a8xbY94a89aYt8a8NeatMW+NeWvM
W2PeGvPWmLfGvDXmrTVvrXlrzVtr3lrz1pq31ry15q01b615a81ba95a89aat9a8teatNW+teWvNW2veWvPWmrfWvLXmrTVvrXnrzFtn3jrz1pm3zrx15q0zb51568xbZ94689aZt868
deatM2+deevMW2feOvPWmbfOvHXmrTNvnXnrzFtn3nrz1pu33rz15q03b715681bb95689abt9689eatN2+9eevNW2/eevPWm7fevPXmrTdvvXnrzVtv3nrz1pu3wbwN5m0wb4N5G8zb
YN4G8zaYt8G8DeZtMG+DeRvM22DeBvM2mLfBvA0b/thp+s/7sdP0px87/fzfH6a3P/u4BcUtKG9BdQvqW9DcgvYWdLegvwXDLZjfgsUtWN6C1S1Y34LNLdjegt0t2N+Cwy043oLTLTjf
gssUP/sAQTshnlBPyCf0EwLqY8HRo/3sg6gj6okGojnRgmhJtCJaE22ItkQ7oj3RgehIdCI6E12APtfUfvZh3zP7Luy7tO/Kvmv7buy7te/Ovnv7Hux7bt8L+17a98q+1/a9se+tfe/s
e2/fB/s+2vfJvs/2fZnazz7s27zNzNvMvM3M28y8zczbzLzNzNvMvM3M28y8zczbzLzNzNvMvM3M28y8zczbzLzNzNvMvM3M28y8zczbzLwV5q0wb4V5K8xbYd4K81aYt8K8FeatMG+F
eSvMW2HeCvNWmLfCvBXmrTBvhXkrzFth3grzVpi3wrwV5q0wb6V5K81bad5K81aat9K8leatNG+leSvNW2neSvNWmrfSvJXmrTRvpXkrzVtp3krzVpq30ryV5q00b6V5K81bZd4q81aZ
t8q8VeatMm+VeavMW2XeKvNWmbfKvFXmrTJvlXmrzFtl3irzVpm3yrxV5q0yb5V5q8xbZd4q81abt9q81eatNm+1eavNW23eavNWm7favNXmrTZvtXmrzVtt3mrzVpu32rzV5q02b7V5
q81bbd5q81abt9q8NeatMW+NeWvMW2PeGvPWmLfGvDXmrTFvjXlrzFtj3hrz1pi3xrw15q0xb415a8xbY94a89aYt8a8NeatMW+teWvNW2veWvPWmrfWvLXmrTVvrXlrzVtr3lrz1pq3
1ry15q01b615a81ba95a89aat9a8teatNW+teWvNW2feOvPWmbfOvHXmrTNvnXnrzFtn3jrz1pm3zrx15q0zb51568xbZ94689aZt868deatM2+deevMW2feOvPWm7fevPXmrTdvvXnr
zVtv3nrz1pu33rz15q03b715681bb95689abt9689eatN2+9eevNW2/eevPWm7fevA3mbTBvg3kbzNtg3gbzNpi3wbwN5m0wb4N5G8zbYN4G8zaYt8G8DeZt+NHb1c8+7v55P/u4w88+
7m5/9nELiltQ3oLqFtS3oLkF7S3obkF/C4ZbML8Fi1uwvAWrW7C+BZtbsL0Fu1uwvwWHW3C8BadbcL4Flzv87AME7YR4Qj0hn9BPCKiPBa9+9kHUEfVEA9GcaEG0JFoRrYk2RFuiHdGe
6EB0JDoRnYkud/zZhz2R/dy+Z/Zd2Hdp35V91/bd2Hdr35199/Y92Pfcvhf2vbTvlX2v7Xtj31v73tn33r4P9n2075N9n+37Yo+HzdvMvM3M28y8zczbzLzNzNvMvM3M28y8zczbzLzN
zNvMvM3M28y8zczbzLzNzNvMvM3M28y8zczbzLzNzNvMvBXmrTBvhXkrzFth3grzVpi3wrwV5q0wb4V5K8xbYd4K81aYt8K8FeatMG+FeSvMW2HeCvNWmLfCvBXmrTBvpXkrzVtp3krz
Vpq30ryV5q00b6V5K81bad5K81aat9K8leatNG+leSvNW2neSvNWmrfSvJXmrTRvpXkrzVtl3irzVpm3yrxV5q0yb5V5q8xbZd4q81aZt8q8VeatMm+VeavMW2XeKvNWmbfKvFXmrTJv
lXmrzFtl3irzVpu32rzV5q02b7V5q81bbd5q81abt9q81eatNm+1eavNW23eavNWm7favNXmrTZvtXmrzVtt3mrzVpu32rw15q0xb415a8xbY94a89aYt8a8NeatMW+NeWvMW2PeGvPW
mLfGvDXmrTFvjXlrzFtj3hrz1pi3xrw15q0xb615a81ba95a89aat9a8teatNW+teWvNW2veWvPWmrfWvLXmrTVvrXlrzVtr3lrz1pq31ry15q01b615a81bZ94689aZt868deatM2+d
eevMW2feOvPWmbfOvHXmrTNvnXnrzFtn3jrz1pm3zrx15q0zb51568xbZ94689abt9689eatN2+9eevNW2/eevPWm7fevPXmrTdvvXnrzVtv3nrz1pu33rz15q03b715681bb95689ab
t968DeZtMG+DeRvM22DeBvM2mLfBvA3mbTBvg3kbzNtg3gbzNpi3wbwN5m340dvVzz7u/3k/+7jHzz7ub3/2cQuKW1DeguoW1LeguQXtLehuQX8LhlswvwWLW7C8BatbsL4Fm1uwvQW7
W7C/BYdbcLwFp1twvgWXe/zsAwTthHhCPSGf0E8IqI8Fr372QdQR9UQD0ZxoQbQkWhGtiTZEW6Id0Z7oQHQkOhGdiS73/NmH/Xvzc/ue2Xdh36V9V/Zd23dj3619d/bd2/dg33P7Xtj3
0r5X9r227419b+17Z997+z7Y99G+T/Z9tu+LnS3zNjNvM/M2M28z8zYzbzPzNjNvM/M2M28z8zYzbzPzNjNvM/M2M28z8zYzbzPzNjNvM/M2M28z8zYzbzPzNjNvhXkrzFth3grzVpi3
wrwV5q0wb4V5K8xbYd4K81aYt8K8FeatMG+FeSvMW2HeCvNWmLfCvBXmrTBvhXkrzFtp3krzVpq30ryV5q00b6V5K81bad5K81aat9K8leatNG+leSvNW2neSvNWmrfSvJXmrTRvpXkr
zVtp3krzVpm3yrxV5q0yb5V5q8xbZd4q81aZt8q8VeatMm+VeavMW2XeKvNWmbfKvFXmrTJvlXmrzFtl3irzVpm3yrzV5q02b7V5q81bbd5q81abt9q81eatNm+1eavNW23eavNWm7fa
vNXmrTZvtXmrzVtt3mrzVpu32rzV5q02b415a8xbY94a89aYt8a8NeatMW+NeWvMW2PeGvPWmLfGvDXmrTFvjXlrzFtj3hrz1pi3xrw15q0xb415a8xba95a89aat9a8teatNW+teWvN
W2veWvPWmrfWvLXmrTVvrXlrzVtr3lrz1pq31ry15q01b615a81ba95a89aZt868deatM2+deevMW2feOvPWmbfOvHXmrTNvnXnrzFtn3jrz1pm3zrx15q0zb51568xbZ94689aZt868
9eatN2+9eevNW2/eevPWm7fevPXmrTdvvXnrzVtv3nrz1pu33rz15q03b715681bb95689abt9689eatN2+DeRvM22DeBvM2mLfBvA3mbTBvg3kbzNtg3gbzNpi3wbwN5m0wb4N5G370
dvWzj4d/3s8+HvCzj4fbn33cguIWlLegugX1LWhuQXsLulvQ34LhFsxvweIWLG/B6hasb8HmFmxvwe4W7G/B4RYcb8HpFpxvweUBP/sAQTshnlBPyCf0EwLqY8Grn30QdUQ90UA0J1oQ
LYlWRGuiDdGWaEe0JzoQHYlORGeiywN/9mG/98fn9j2z78K+S/uu7Lu278a+W/vu7Lu378G+5/a9sO+lfa/se23fG/ve2vfOvvf2fbDvo32f7Pts35cH+9mHfZu3mXmbmbeZeZuZt5l5
m5m3mXmbmbeZeZuZt5l5m5m3mXmbmbeZeZuZt5l5m5m3mXmbmbeZeZuZt5l5m5m3wrwV5q0wb4V5K8xbYd4K81aYt8K8FeatMG+FeSvMW2HeCvNWmLfCvBXmrTBvhXkrzFth3grzVpi3
wrwV5q00b6V5K81bad5K81aat9K8leatNG+leSvNW2neSvNWmrfSvJXmrTRvpXkrzVtp3krzVpq30ryV5q00b6V5q8xbZd4q81aZt8q8VeatMm+VeavMW2XeKvNWmbfKvFXmrTJvlXmr
zFtl3irzVpm3yrxV5q0yb5V5q8xbZd5q81abt9q81eatNm+1eavNW23eavNWm7favNXmrTZvtXmrzVtt3mrzVpu32rzV5q02b7V5q81bbd5q81abt8a8NeatMW+NeWvMW2PeGvPWmLfG
vDXmrTFvjXlrzFtj3hrz1pi3xrw15q0xb415a8xbY94a89aYt8a8NeatNW+teWvNW2veWvPWmrfWvLXmrTVvrXlrzVtr3lrz1pq31ry15q01b615a81ba95a89aat9a8teatNW+teevM
W2feOvPWmbfOvHXmrTNvnXnrzFtn3jrz1pm3zrx15q0zb51568xbZ94689aZt868deatM2+deevMW2feevPWm7fevPXmrTdvvXnrzVtv3nrz1pu33rz15q03b715681bb95689abt968
9eatN2+9eevNW2/eevPWm7fBvA3mbTBvg3kbzNtg3gbzNpi3wbwN5m0wb4N5G8zbYN4G8zaYt8G8DT96u/rZx+M/72cfj/jZx+Ptzz5uQXELyltQ3YL6FjS3oL0F3S3ob8FwC+a3YHEL
lrdgdQvWt2BzC7a3YHcL9rfgcAuOt+B0C8634PKIn32AoJ0QT6gn5BP6CQH1seDVzz6IOqKeaCCaEy2IlkQrojXRhmhLtCPaEx2IjkQnojPR5ZE/+7Dvz+17Zt+FfZf2Xdl3bd+Nfbf2
3dl3b9+Dfc/te2HfS/te2ffavjf2vbXvnX3v7ftg30f7Ptn32b4vj/azD/s2bzPzNjNvM/M2M28z8zYzbzPzNjNvM/M2M28z8zYzbzPzNjNvM/M2M28z8zYzbzPzNjNvM/M2M28z8zYz
b4V5K8xbYd4K81aYt8K8FeatMG+FeSvMW2HeCvNWmLfCvBXmrTBvhXkrzFth3grzVpi3wrwV5q0wb4V5K8xbad5K81aat9K8leatNG+leSvNW2neSvNWmrfSvJXmrTRvpXkrzVtp3krz
Vpq30ryV5q00b6V5K81bad5K81aZt8q8VeatMm+VeavMW2XeKvNWmbfKvFXmrTJvlXmrzFtl3irzVpm3yrxV5q0yb5V5q8xbZd4q81aZt8q81eatNm+1eavNW23eavNWm7favNXmrTZv
tXmrzVtt3mrzVpu32rzV5q02b7V5q81bbd5q81abt9q81eatNm+NeWvMW2PeGvPWmLfGvDXmrTFvjXlrzFtj3hrz1pi3xrw15q0xb415a8xbY94a89aYt8a8NeatMW+NeWvMW2veWvPW
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
Cekp6RmpIJUklg3Thm3DuGHdMG/YNwwcFg4Th43DyGHlMHPYOQwdlg5Th63D2KnVrp2Dkz/tHJzgHJzcPwf3obgP5X1on+AcQLqQHqQPGUCGkBFkDJlAppAZpILMIQvIErKCrCEbyBay
O+E5ID0lPSMVpJLEsmHasG0YN6wb5g37hoHDwmHisHEYOawcZg47h6HD0mHqsHUYO7XatXNw+qedg1Ocg9P75+A+FPehvA/tU5wDSBfSg/QhA8gQMoKMIRPIFDKDVJA5ZAFZQlaQNWQD
2UJ2pzwHpKekZ6SCVJJYNkwbtg3jhnXDvGHfMHBYOEwcNg4jh5XDzGHnMHRYOkwdtg5jp1a7dg7O/rRzcIZzcHb/HNyH4j6U96F9hnMA6UJ6kD5kABlCRpAxZAKZQmaQCjKHLCBLyAqy
hmwgW8jujOeA9JT0jFSQShLLhmnDtmHcsG6YN+wbBg4Lh4nDxmHksHKYOewchg5Lh6nD1mHs1GrXzsH5n3YOznEOzu+fg/tQ3IfyPrTPcQ4gXUgP0ocMIEPICDKGTCBTyAxSQeaQBWQJ
WUHWkA1kC9md8xyQnpKekQpSSWLZMG3YNowb1g3zhn3DwGHhMHHYOIwcVg4zh53D0GHpMHXYOoydWu3aOWg8+dPuwfuPqh+Ej/TFRYAUkBLS/vTRX1wFUpfUI/VJA9KQNCKNSRPSlDQj
VaQ5aUFaklakNWlD2pJ2pETsqdgzsUKsFJPmkeiR6pHske6R8JHykfSR9pH4kfqR/JH+kQFEFhCZQGQDkRFEVhCZQeo7qN+Qxp93Qxq8IQ3ckAZuSAM3pIEb0uANafCGNHhDGrwhDd6Q
Bm9IgzekwRvS4A1p8IY0eEMavCEN3pAGb0iDN6TBG9LgDWnwhjR4Qxq8IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25IQ25
IQ25IQ25IQ25IQ25IQ25IY1/eUP+vBfMjSZvSBM3pIkb0sQNaeKGNHlDmrwhTd6QJm9IkzekyRvS5A1p8oY0eUOavCFN3pAmb0iTN6TJG9LkDWnyhjR5Q5q8IU3ekCZvSFNuSFNuSFNu
SFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSFNuSPNf3pDDXjg3ju4/uXj6ic6/OBSgglSS2qQOqUvqkfqkAWlI
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
gsgMUt9B/VYc9m7znO82z/luk1SQSlKb1CF1ST1SnzQgDUkj0pg0IU1JM1JFmpMWpCVpRVqTNqQtaXcu7zbFpHckeKR4JHmkeSR6pHoke6R7JHykfCR9pH0kfqR+JH+kf2QAkQVEJhDZ
QGQEkRVEZpD6Dmq3ovHksIebH358/Vp8ti/OhVghVoq1xTpiXbGeWF9sIDYUG4mNxSZiU7GZWCU2F1uILcVWYmuxjdhWbCeWGNoSYlOIbSE2htgaYnOI7SE2iNgiYpOIbSI2itgqYrOI
7SI2jNgyYtOIbSM2jtg6YvOI7SM2kNxbyIcD8/jm58vL2+Li9uLH799cXv90+ezy9eubB8+vfrk7Ks2HX+iD68uXd7flu17j7uvve/O7XvP95/7xMfuf+erti1e3r67eXrxuX12/ubi9
ffX2pwc3//7hS1rfffim9j/uZfXL68sHt/98d/nDw8vf311f3tzsv+rhgxe/v+y92H8j+5P37vrV1fWr23/ufwH7r3i5/7hfXl/8mHHxVVp/eX37t7/8dPu3hw+/nlRfVZP13h790Pr6
2WS4HI3f/0nj0aPvH3/6qu8ff/gZ7/7AvsUfv3938dPl6OL6p1dvbx68vnz5/sLu/73v+tVPP3/649urd+//6Pjhg79f3d5evfn0Zz9fXry4vL77s6OHD15eXd1++pO7357frq7/8f43
/cf/D1BLAwQUAAAACAAPjGRdhDUEpN0xAACOqgEAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQ0LnhtbK3daY9c57le4e/5FQQD2FZgiaypB1nSwd29d1XtmucpCAKabEqEOeh0tyw5Qf57
mqO4ey0nLoPAObZ0iV1sat16ZAEv7O/+47dXLx/8/er65sWb19//sfHN4z8+uHr99M2zF69//P6P61X367M/Pri5ffL62ZOXb15fff/Hf1zd/PE/fvgv3/365vpvNz9dXd0+uPuA1zff
P/zp9vbnbx89unn609WrJzffvPn56vXdH3n+5vrVk9u7373+8dHNz9dXT569+6JXLx81Hz8+efTqyYvXD99/wrfX/8pnvHn+/MXTq+LN019eXb2+ff8h11cvn9zeffs3P734+ebjp716
+q983Ksn13/75eevn7559fPdR/z1xcsXt/9496EfP+a3RvsJP+nVi6fXb27ePL/95u4rP3xP/OWdPzp/9OTpp0+6/vc+ptG++wX+/cXbPr9/VPPf/KzOp89q/v5hrX/zw04+fVjr4YNX
T7+tfnz95vrJX19eff/w3Z+2B79d3/1f8+7/7/74b9ff/vLi2fcP//d547RxmvPO1+cnF82v283T1tdn7dN8fV487pxeXJ6nyOn/efjDd89e3BV+++EPrq+ef/8wjW9zaDx+/PDRD9+9
+x42L65+vfnstx+8XeRf37z529vfqe5+psd3H3Jz9fLq6dttPHhy929/v7q8evny+4cXd3lv/vPdx9795t0HPvr0KZ//9sdP774bzOz6wbOr509+eXm7ePNr/+rFjz/dfv+w0bn7lb39
tX777B/F1c3Tu03e/cTfNDtvP/Xpm5c37/71wasXr+9+7N2fpCe/vfv3X188u/3p+4fNs286p49bjbsf/uDpLze3b15t3/+Bxocvf/+FzQ9f2Gl++spG65vTxuPz1ql+5aP3P/O7b754
cvvkh++u3/z64PrdT33z85O3f7U2vn37aW9/Ca3mN42PH/Lpl6W/qvbdn9Cnbz8mbz/n7vu5+9N49wdu7vjvPzz+7tHf3/7MH37Ixacf8uiDXEIKSAnpQnqQPqSCDN5Lo/1JhpARZAyZ
QKaQGWQOWUCWkBVkDdlAtpAdZA85QBLSBemSVJBKUpfUI/VJFYl1w7xh3zBwWDhMHDYOI4eVw8xh5zB0WDpMHbYOY6dW+9Hdcfh0IZpyIZqNb847/9qJuLtkH3+O5vu/BE8/uxGdezfi
0w/5dCMgBaSEdCE9SB9SQQZN3AjICDKGTCBTyAwyhywgS8gKsoZsIFvIDrKHHJq8EaQL0iWpIJWkLqlH6pMqEuuGecO+YeCwcJg4bBxGDiuHmcPOYeiwdJg6bB3GTq127Ua07t2I//cd
aL3/mN//VnwBuYQUkBLShfQgfUgFGbTwlz1kBBlDJpApZAaZQxaQJWQFWUM2kC1kB9lDDi3+ZU+6IF2SClJJ6pJ6pD6pIrFumDfsGwYOC4eJw8Zh5LBymDnsHIYOS4epw9Zh7NRq1/6y
b3+x/2jQfv9T3P1Hg6ev3v7wH757/vY/JDy5vn7yj4cf/nGv/W2vcfdHqm65WEwXf/qfv718/vqbUbl6+1s/v/rmpz+v3v6TZuO//9f+3T+UXl3f/I8/v/8Dzz79gbf/5PNRf7m5+vOf
RmX39y9vfnX3V9Ho4Vf/7cNHX+wvp6M/ffyMDz9dxhdFPuDTP19O15PVp9/96quv/vDj7V8ef/Xhp3j+04evuvvXX2++6VajVbn49NN9+jY+/fBn//yHP+MP/+3Dj15Nf/827z6i8fGP
39Q+bTldrD78utaTar4uP/+1//aHJ69+/svDh39u3/0SPnz56w9fPs6wzGKR/Z8W0+3ywxfcfHX3Sx+tx5Plp5/4K/szdP3nj3+mluvxn77++tNPevcDL/vT6bK8+5jfP+Pjj/7q9+/n
+2pSlLs/ffZLevuxX3314Ru9+5zNcpXL4YcP7b//nYfFi5vbF6+f3j64ufrPX65eP726efhnVn1t3/P/+vi9fvjd90k/+wk//hzLDx/94E/PX1zf3D5oP3j24scXtzdfPfzU/96X1H8V
r99+5MOHX3336Pnn/xH4018Kn/5eCCkgJaQL6UH6kAoy+CBnv/+9EDKCjCETyBQyg8whC8gSsoKsIRvIFrKD7CEHSEK6IF2SClJJ6pJ6pD6pIrFumDfsGwYOC4eJw8Zh5LBymDnsHIYO
S4epw9Zh7NRq1/5e2Plify/s3P9L8AJyCSkgJaQL6UH6kAoy6OAkQEaQMWQCmUJmkDlkAVlCVpA1ZAPZQnaQPeTQ4UkgXZAuSQWpJHVJPVKfVJFYN8wb9g0Dh4XDxGHjMHJYOcwcdg5D
h6XD1GHrMHZqtWsn4eSLnYQTnIT38tlFuA/FfSjvQ/c+9O5D/z5U92Hw4Ts7//0UQEaQMWQCmUJmkDlkAVlCVpA1ZAPZQnaQPeQASUgXpEtSQSpJXVKP1CdVJNYN84Z9w8Bh4TBx2DiM
HFYOM4edw9Bh6TB12DqMnVrt2ik4/WKn4BSn4PT+KbgPxX0o70P3PvTuQ/8+VPdhcIpTABlBxpAJZAqZQeaQBWQJWUHWkA1kC9lB9pDDKU8B6YJ0SSpIJalL6pH6pIrEumHesG8YOCwc
Jg4bh5HDymHmsHMYOiwdpg5bh7FTq107BWdf7BSc4RSc3T8F96G4D+V96N6H3n3o34fqPgzOcAogI8gYMoFMITPIHLKALCEryBqygWwhO8gecjjjKSBdkC5JBakkdUk9Up9UkVg3zBv2
DQOHhcPEYeMwclg5zBx2DkOHpcPUYeswdmq1a6fg/IudgnOcgvP7p+A+FPehvA/d+9C7D/37UN2HwTlOAWQEGUMmkClkBplDFpAlZAVZQzaQLWQH2UMO5zwFpAvSJakglaQuqUfqkyoS
64Z5w75h4LBwmDhsHEYOK4eZw85h6LB0mDpsHcZOrXbtFDQef7Fb8O6j6sfgA312DSAFpIR0IT1IH1JBBh+/xc+uAmlEGpMmpClpRpqTFqQlaUVakzakLWlH2pMOpETsQuxSrBArxbpi
PbG+WCUm3SPhI+Uj6SPtI/Ej9SP5I/0jA4gsIDKByAYiI4isIDKD1HdQvx+NL3c/GrwfDdyPBu5HA/ejgfvRwP1o4H40cD8auB8N3o8G70eD96PB+9Hg/WjwfjR4Pxq8Hw3ejwbvR4P3
o8H70eD9aPB+NHg/GrwfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfDbkfjX96P77cM99Gk/ej
ifvRxP1o4n40cT+auB9N3I8m7kcT96PJ+9Hk/WjyfjR5P5q8H03ejybvR5P3o8n70eT9aPJ+NHk/mrwfTd6PJu9Hk/ejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfej
KfejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfejKfej+U/vR+vL3Y8W70cL96OF+9HC/WjhfrRwP1q4Hy3cjxbuR4v3o8X70eL9aPF+tHg/WrwfLd6PFu9Hi/ejxfvR
4v1o8X60eD9avB8t3o8W70dL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kdL7kfrn96PL/eWuNHm
/WjjfrRxP9q4H23cjzbuRxv3o4370cb9aPN+tHk/2rwfbd6PNu9Hm/ejzfvR5v1o8360eT/avB9t3o8270eb96PN+9Hm/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL
/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/WjL/Wj/0/vx5d5fNvgA8wN9fj86uB8d3I8O7kcH96OD+9HB/ejgfnR4Pzq8Hx3ejw7vR4f3o8P70eH96PB+dHg/
OrwfHd6PDu9Hh/ejw/vR4f3o8H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H505H50/un9+HKP
NRt8rdnAc01IASkhXUgP0odUkEGDzzZJI9KYNCFNSTPSnLQgLUkr0pq0IW1JO9KedGjIM06xC7FLsUKsFOuK9cT6YpWYdI+Ej5SPpI+0j8SP1I/kj/SPDCCygMgEIhuIjCCygsgMUt9B
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
v3v0Ub979P4T3/6GfQP/v++rc/d9df6V7+vx599X8/Pvq/PvfV8/P/nxavzk+scXr28evLx6/u6u3/3T5vWLH3/6+Nu3b35+91udhw/++ub29s2rj7/309WTZ1fXb3+v9fDB8zdvbj/+
ztsov765/tu71D/8X1BLAwQUAAAACAAPjGRdXkVIgD0XAAAcngAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQ1LnhtbJ2dbW8bV5KFv++vIDhAYg9kS+e+9IvH8kBxSx0nN3ZWlmdmMRgs
GKllC5FEhaRiZxf737dJNSnK59yMK0Fik0d9b9cttqoeGzml53/9dHU5+rWbzS+m1/tf4+ne16Pu+nR6dnH9fv/rdydHT6qvR/PF5Ppscjm97va//q2bf/3XF//x/ON09vP8Q9ctRv0G
1/P98YfF4ubZ7u789EN3NZk/nd501/1Xzqezq8mifzt7vzu/mXWTs9Wiq8tdt7dX7F5NLq7Hdzs8m33JHtPz84vTrpme3l5114u7TWbd5WTRhz//cHEzX+92dfol211NZj/f3jw5nV7d
9Fv8dHF5sfhttel6m08IE97p6uJ0Np1PzxdP+5VDTHy8erfenZxudpr9sW0Q+gP+erH8fO63cn9wr7jZy91v5v/gZsVmMz8eXZ0+e/X+ejqb/HTZ7Y9XaRt9mvX/uv6//uufZs9uL872
x/97VB0cNN8c4MlhbIonIRZHTw58XT8Jh3uHe3uubHBU/t/4xfOzi/4TXm4+mnXn++MDPEt9Tse7L56vQvjbRfdxvvV6tJjepO588bK7vNwfNxiPlk/oT9Ppz8uvvurvvNdvOu8uu9Pl
szKa9L/92t1d/QP2xqP5L6v7LF/399jdbLz9en3Do9Uj9ONsdNadT24vF8fTj992F+8/LPbHiP1Zl6d/dvZb081P+6e0v/VTF5e7nk4v56tfR1cX1/21fdomn1a/f7w4W3zoX/UP3unt
fDG9+vsgDMvuFrhhgdss8OVTV0XEor/B7630w0q/WRn9l60Mw8rwpUHGYUG8X1A9LbFX+/L371QMCwvz6cphZblZ6fZ+d0E1LKg2C8LvXl8P19dfmoTlE3X32e6ZT4PNcwHzp4X1IwL3
b0LdvXsYV89zM1lMXjyfTT+OZquncX4zWZZ0PFvusnyqXfm03tx486jLJ93H/vvsdLnRwXKnZejjUf+VeS//+sKH57u/Lm8+XPPN5prdQXlJSkPK4bay28e9Cd6p4PeeYm/7H3zZSdz9
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
wBQZmCIDUyRLKWQpMhj/L3KmyMAUyVITmCIDUyQvbANTZGCKDEyRLKWQpchgo8jAFBmYIllqAlNkYIrkhW1gigxMkYEpkqUUshQZbBQZmCIDUyRLTWCKDEyRvLANTJGBKTIwRbKUQpYi
g40iA1NkYIpkqQlMkYEpkhe2gSkyMEUGpkiWUshSZLBRZGCKDEyRLDWBKTIwRfLCNjBFBqbIwBTJUgpZigw2igxMkYEpkqUmMEUGpkhe2AamyMAUGZgiWUohS5HRRpGRKTIyRbLURKbI
yBTJC9vIFBmZIiNTJEspZiky2igyMkVGpkiWmsgUGZkieWEbmSIjU2RkimQpxSxFRhtFRqbIyBTJUhOZIiNTJC9sI1NkZIqMTJEspZilyGijyMgUGZkiWWoiU2RkiuSFbWSKjEyRkSmS
pRSzFBltFBmZIiNTJEtNZIqMTJG8sI1MkZEpMjJFspRiliKj0SLDFBmZIllqIlNkZIrkhW1kioxMkZEpkqUUsxQZbRQZmSIjUyRLTWSKjEyRvLCNTJGRKTIyRbKUYpYio40iI1NkZIpk
qYlMkZEpkhe2kSkyMkVGpkiWUsxSZLRRZGSKjEyRLDWRKTIyRfLCNjJFRqbIyBTJUopZiow2ioxMkZEpkqUmMkVGpkhe2EamyMgUGZkiWUoxS5GFjSKLgem2bIRractHyFKzlrachGtp
y0rIC9u1tGUmLJgiC6ZIllKRpcjCRJFNccdhrt4KipmuYKZjKRVZpiucLSjHQTFhFUxYLKUiS1iFtwXlOSjmnYJ5h6VUZHmnCLagAgfF9FEwfbCUiix9FPf0Eb4kqMhBMQsUzAIspSLL
AkVhC6rgoLgzF9yZWUpFtjMXpS2okoPiPllwn2QpFdk+WVS2oCoOirtWwV2LpVRku1ZR24KqOSjuIQX3EJZSke0h5Z4pqHKPgiq5opdc0VlKZbail7AFxRW95IpeckVnKZXZil46W1Bc
0Uuu6CVXdJZSma3opbcFxRW95IpeckVnKZXZil4GW1Bc0Uuu6CVXdJZSma3opa2il1zRS67oJVd0llKZreilraKXXNFLruglV3SWUpmt6KWtopdc0Uuu6CVXdJZSma3opa2il1zRS67o
JVd0llKZreilraKXXNFLruglV3SWUpmt6JWtoldc0Suu6BVXdJZSla3ola2iV1zRK67oFVd0llKVreiVraJXXNErrugVV3SWUpWt6JWtoldc0Suu6BVXdJZSla3ola2iV1zRK67oFVd0
llKVreiVraJXXNErrugVV3SWUpWt6JWtoldc0Suu6BVXdJZSla3ola2iV1zRK67oFVd0llKVreiVraJXXNErrugVV3SWUpWt6JWtoldc0Suu6BVXdJZSla3ota2i11zRa67oNVd0llKd
rei1raLXXNFrrug1V3SWUp2t6LWtotdc0Wuu6DVXdJZSna3ota2i11zRa67oNVd0llKdrei1raLXXNFrrug1V3SWUp2t6LWtotdc0Wuu6DVXdJZSna3ota2i11zRa67oNVd0llKdrei1
raLXXNFrrug1V3SWUp2t6LWtotdc0Wuu6DVXdJZSna3ota2i11zRa67oNVd0llKdrejYs5X0u+sfhrXWHkxa2OOqLrS00VRoMIYGERpEaBChQYSGfGjOGJoToTkRmhOhORGay4fmjaF5
EZoXoXkRmheh+XxowRhaEKEFEVoQoQURWsiHFo2hRRFaFKFFEVoUocV8aIUxtEKEVojQChFaIUIr8qGVxtBKEVopQitFaKUIrcyHVhlDq0RolQitEqFVIrQqH1ptDK0WoYnhNntius2e
GG+zl+8GMHYDiG4A0Q0gugFEN0C+G8DYDSC6AUQ3gOgGEN0A+W4AYzeA6AYQ3QCiG0B0A+S7AYzdAKIbQHQDiG4A0Q2Q7wYwdgOIbgDRDSC6AUQ3QL4bwNgNILoBRDeA6AYQ3QD5bgBj
N4DoBhDdAKIbQHQD5LsBjN0AohtAdAOIbgDRDZDvBjB2A4huANENILoBRDdAvhvA2A0guoEadaZmnalhZ/lpZ3DGbuBENxDjxyDmjwktIT+BDM7YDZzoBmIkGMRMMKEl5KeCwRm7gRPd
QIzpgpjTJbSE/KQuOGM3cKIbiNFZELOzhJaQn54FZ+wGTnQDMc4KYp6V0BLyE63gjN3AiW4gRkxBzJgSWkJ+yhScsRs40Q3E2CeIuU9CS8hPfoIzdgMnuoEYxQQxi0loCflpTHDGbuBE
NxDjkSDmIwktIT8hCc7YDZzoBmJkEcTMIqEl5KcWwRu7gRfdQIwRgpgjJLSE/CQheGM38KIbiNE+ELN9hJaQn+4Db+wGXnQDMW4HYt6O0BLyE3fgjd3Ai24gRuBAzMARWkJ+Cg68sRt4
0Q3EWBqIuTRCS8hPpoE3dgMvuoEYFQMxK0ZoCflpMfDGbuBFNxDjWyDmtwgtIT/BBd7YDbzoBmKkCsRMFaEl5KeqwBu7gRfdQIw5gZhzIrSE/KQTeGM38KIbiNEjELNHhJaQnz6CYOwG
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
/P/e/ulPPrt/eYevF7/1+Obz927/LL37ieY7t9/56U+OW/g/79/8+ckrH9/67w8ffvW7e3evX73duf3ND391fbu/fDa8PrR/ePjw368nnC83c+fydR7dfXBz66+/e3R569+7XW7fevrw
0cc3nz99/+bLy0k/03H71t17T+//6eY3l+Peu/2Hh0+fPvzqesDt67fF08vo88cP//vNg+PGbr68uXd9Em89mg5+dpFnV/1FvXydJ/9xvK7rx5dT33n5Ql79+MUL/PB4ZH/z+NZnN5/f
/frLp799+Odf3Nz/4xeXO06XSx3bfvezv35w8+Te5bvi8rreluOq9x5++eT4562v7j+4HHt5m+7+5fj3n+9/9vSLy0ft7dLuaLocfuve108ud/t/Pf/E89OfnSjPT5RvThzfeYI+P0G/
OUHezlJa/599qfz8zPzyTKlvt3RnaPvuE8vzE8s3J3boxdXnJ9bVrbTnJ7bVW+3PT+yrtzqenzhWbzXdefHW31m92fTyqUmrt5tePDdJwRt+59kTezz0H9x9evenP3n88M+3Hh+P7JPL
N9Xl15n07vVq10df77y4xMtvBvN7IV++y+9dL/Kz61Uub/TlX5dPPLmM//TT3H7yzp+uX/j5Mf/y8ph3nk/enyYfTJPTNPlwmvx8mvximpynyS+nyUfT5ONp8smrk3cuO3y5SLEWKW8n
bJWXd+3FLuXZV5BXd9m/tcuXx7zc5TT5YJqcpsmH0+Tn0+QX0+Q8TX45TT6aJh9Pk09enby2S/3WLr97YfrsMndeXvhfpsn70+SDaXKaJh9Ok59Pk19Mk/M0+eU0+WiafDxNPnl18tp6
8tJ68rSeafL+NPlgmpymyYfT5OfT5BfT5DxNfjlNPpomH0+TT/LfWk9ZWk95cZlXvt3Gt77djGPKndePed84Rr/1U+AH1nXS68ecrGPk9WM+tI7R14/5uXVMfv2YX1jHlNePOVvH1NeP
+aV1zLde+0fWMd/6ae1j65hvvRefGMfUb96L1x6EuvQg1GcX7t98nzyfXM6899X1J+6f/uTz6xe9+/jx3b/eflbPl2N+ev7wh6e/3Lt5dPxW6dNrp//rv/6XT7+4/+TWJSf/7a1//dnv
fvfJrV99/dUfbh7/27+9d/v2W5f//D9/+fLzB29/fPr0h5ePHn319oO38Ev8092vHv3zs2tczvzireP49K//5ReX30Vcfr/5b88/8dnLT1x/wf/2WfefPnl+E//yf7//649/+OMfP7+T
z957fkc/enGTP/vkXz742fPP3nvrd7//5MXHP/rRj946f3j67W9//dsfPjv209N/+/SXvz7/6oe337p1+61Pf/v70/OLXP755ydvf3j++NPTb3/44sZ/+PHpw09f/kh+9N7tTz++/aP/
44ff3OI//fHpP9+5fpnbv3r49NbnD79+8Nntyw9/8s7nr37rvXyfLm/Q93o3jvtYfwPyt+/kg+d3Mr7jifkAeGJ+d/MfX988uHfzNx+XPz2f/LePf/3rj37/G/RqL278xeSy1Jsn33ze
/vRv7t5/8PTWO7c+uHl69/6Xl8fr9u3r2/7iTt6789bLp+pP03tzAjZy+kfbyOkvj768++D4o5G1bXz4fBvlO7bx4T/aNj74+ubW5WeYm7VV/PzZKtorqTBNzs/XVb9JhefHpO/9rf4/
O+fXf7p5/Pj+Zze3vv3Y/9s/ffn0n68/H10u8j2v8N3nffDsjyDm0769uo/+fkt49UlfX8BrZ2Mv/rVTvv3CP37+PLT/9S/85UO9/qq/ORV7yd8c/+3X+8nz15v/l73eT+4+vffFzWe3
Pv34h09+dD3rm19m37p9XPDBN4PLxd7/9e9/9enPvu930LurC/zR8ev/W7dfnnfryc3Ty539/snNkxd/ZPdaD7yWf20p/9qUfw3Ivxb598bzr7nJvwbkX6PKP2Ajp0aTfw3Iv8aRf23K
v2lyblP+tci/2x811vxrZPnXIv828q8v5V+f8q8D+dcj/954/nU3+deB/OtU+Qds5NRp8q8D+dc58q9P+TdNzn3Kvx75d/ujzpp/nSz/euTfRv6NpfwbU/4NIP9G5N8bz7/hJv8GkH+D
Kv+AjZwGTf4NIP8GR/6NKf+myXlM+Tci/25/NFjzb5Dl34j828i/6/8ef6H/jsNfD8AXo+8swMtBkYBvOgG/eafeeAO+uJXvjEDkoflPVIHITk7/cDv5/h34Yh/fGYL/cPv4niX4fBmv
puA8Or9Y2Ssx+OIo7hr8O27BWQ6+eCRoevDFC44g/H5BmNaCMM1BmJAgTBGEbz4Ik58gTEgQJq4gTEgQJp4gTEgQJpIgTHMQpjkI0xyEKYLwEoSJNggTWxCmCMKdIJS1IJQ5CAUJQokg
fPNBKH6CUJAgFK4gFCQIhScIBQlCIQlCmYNQ5iCUOQglgvAShEIbhMIWhBJBuBOEa3/tVNI5CBUJQo0gfPNBqH6CUJEgVK4gVCQIlScIFQlCJQlCnYNQ5yDUOQg1gvAShEobhMoWhBpB
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
SFcRSFcRMl1FIF1FiHQVgXQVYdFVxNBVxNBVxNBVJHSVayIKr64idLqKhK6yl4iLuooYuopAuoqEruIhER3pKgLpKkKmqwikqwiRriKQriIsuooYuooYuooYuoqErnIkIq+uInS6ioSu
speIi7qKGLqKQLqKhK7iIREd6SoC6SpCpqsIpKsIka4ikK4iLLqKGLqKGLqKGLqKhK5yJCKvriJ0uoqErrKXiIu6ihi6ikC6ioSu4iERHekqAukqQqarCKSrCJGuIpCuIiy6ihi6ihi6
ihi6ioSuciQir64idLqKhK6yl4iLuooYuopAuoqEruIhER3pKgLpKkKmqwikqwiRriKQriIsuooYuooYuooYuoqErnIkIq+uInS6ioSuspeIi7qKGLqKQLqKhK7iIREd6SoC6SpCpqsI
pKsIka4ikK4iLLqKGLqKGLqKGLqKhK5yJCKvriJ0uoqErrKXiIu6ihi6ikC6ioSu4iERHekqAukqQqarCKSrCJGuIpCuIiy6ihi6ihi6ihi6ioSuciQir64idLqKhK6yl4iLuooYuopA
uoqEruIhER3pKgLpKkKmqwikqwiRriKQriIsuooYuooYuooYuoqErnIkIq+uInS6ioSuspeIi7qKGLqKQLqKhK7iIREd6SoC6SpCpqsIpKsIka4ikK4iLLqKGLqKGLqKGLqKhK5yJCKv
riJ0uoqErrKXiIu6ihi6ikC6ioSu4iERHekqAukqQqarCKSrCJGuIpCuIiy6ihi6ihi6ihi6ioSuciQir64idLqKhK6ylYi6qKuooasopKto6CoOElEd6SoK6SpKpqsopKsoka6ikK6i
LLqKGrqKGrqKGrqKhq5yTUTl1VWUTlfR0FX2EnFRV1FDV1FIV9HQVTwkoiNdRSFdRcl0FYV0FSXSVRTSVZRFV1FDV1FDV1FDV9HQVY5E5NVVlE5X0dBV9hJxUVdRQ1dRSFfR0FU8JKIj
XUUhXUXJdBWFdBUl0lUU0lWURVdRQ1dRQ1dRQ1fR0FWOROTVVZROV9HQVfYScVFXUUNXUUhX0dBVPCSiI11FIV1FyXQVhXQVJdJVFNJVlEVXUUNXUUNXUUNX0dBVjkTk1VWUTlfR0FX2
EnFRV1FDV1FIV9HQVTwkoiNdRSFdRcl0FYV0FSXSVRTSVZRFV1FDV1FDV1FDV9HQVY5E5NVVlE5X0dBV9hJxUVdRQ1dRSFfR0FU8JKIjXUUhXUXJdBWFdBUl0lUU0lWURVdRQ1dRQ1dR
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
ukqCdJVEpKskSFdJLLpKMnSVZOgqydBVUugqRyLy6iqJTldJoavsJeKirpIMXSVBukoKXcVDIjrSVRKkqyQyXSVBukoi0lUSpKskFl0lGbpKMnSVZOgqKXSVIxF5dZVEp6uk0FX2EnFR
V0mGrpIgXSWFruIhER3pKgnSVRKZrpIgXSUR6SoJ0lUSi66SDF0lGbpKMnSVFLrKkYi8ukqi01VS6Cp7ibioqyRDV0mQrpJCV/GQiI50lQTpKolMV0mQrpKIdJUE6SqJRVdJhq6SDF0l
GbpKCl3lSEReXSXR6SopdJWtRJRFXUUMXUUgXUVCV3GQiOJIVxFIVxEyXUUgXUWIdBWBdBVh0VXE0FXE0FXE0FUkdJVrIgqvriJ0uoqErrKXiIu6ihi6ikC6ioSu4iERHekqAukqQqar
CKSrCJGuIpCuIiy6ihi6ihi6ihi6ioSuciQir64idLqKhK6yl4iLuooYuopAuoqEruIhER3pKgLpKkKmqwikqwiRriKQriIsuooYuooYuooYuoqErnIkIq+uInS6ioSuspeIi7qKGLqK
QLqKhK7iIREd6SoC6SpCpqsIpKsIka4ikK4iLLqKGLqKGLqKGLqKhK5yJCKvriJ0uoqErrKXiIu6ihi6ikC6ioSu4iERHekqAukqQqarCKSrCJGuIpCuIiy6ihi6ihi6ihi6ioSuciQi
r64idLqKhK6yl4iLuooYuopAuoqEruIhER3pKgLpKkKmqwikqwiRriKQriIsuooYuooYuooYuoqErnIkIq+uInS6ioSuspeIi7qKGLqKQLqKhK7iIREd6SoC6SpCpqsIpKsIka4ikK4i
LLqKGLqKGLqKGLqKhK5yJCKvriJ0uoqErrKXiIu6ihi6ikC6ioSu4iERHekqAukqQqarCKSrCJGuIpCuIiy6ihi6ihi6ihi6ioSuciQir64idLqKhK6yl4iLuooYuopAuoqEruIhER3p
KgLpKkKmqwikqwiRriKQriIsuooYuooYuooYuoqErnIkIq+uInS6ioSuspeIi7qKGLqKQLqKhK7iIREd6SoC6SpCpqsIpKsIka4ikK4iLLqKGLqKGLqKGLqKhK5yJCKvriJ0uoqErrKV
iLqoq6ihqyikq2joKg4SUR3pKgrpKkqmqyikqyiRrqKQrqIsuooauooauooauoqGrnJNROXVVZROV9HQVfYScVFXUUNXUUhX0dBVPCSiI11FIV1FyXQVhXQVJdJVFNJVlEVXUUNXUUNX
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
'@) -replace '\s', ''

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
                $text = " $pointer $box $($e.Name)   ->  $tlText$(if ($e.Note) { "   [$($e.Note)]" })"
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
        Write-Host (" [{0}] {1}   ->  {2}{3}" -f ($i + 1), $Entries[$i].Name, $tlText, $(if ($Entries[$i].Note) { "   [$($Entries[$i].Note)]" } else { '' }))
    }
    Write-Host ""
    $ticked = @($Entries | Where-Object { $_.Checked })
    $tickedNumbers = @(for ($i = 0; $i -lt $Entries.Count; $i++) { if ($Entries[$i].Checked) { $i + 1 } })
    $dflt = if ($ticked.Count -gt 0) { " (Enter = $($tickedNumbers -join ','))" } else { "" }
    $choice = Read-Host "Numbers separated by commas (e.g. 1,3), or A for all$dflt"
    if (-not $choice -and $ticked.Count -gt 0) { return $ticked }
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
    param([hashtable]$Notes = @{}, [string[]]$PreTick = @())
    $files = @(Get-ChildItem -Filter "*.kss" -File |
        Where-Object {
            $_.Name -notmatch ' - Filtered Assemblies' -and
            $_.Name -notmatch ' \(Generated\)' -and
            $_.Name -notmatch ' - Filtered Assemblies - VALID'
        } |
        Sort-Object Name)
    if ($files.Count -eq 0) { return @() }

    $entries = @(foreach ($f in $files) {
        [PSCustomObject]@{ FullName = $f.FullName; Name = $f.Name; TL = (Get-TransmittalFromName $f.Name); Checked = ($PreTick -contains $f.Name); Note = $(if ($Notes.ContainsKey($f.Name)) { $Notes[$f.Name] } else { '' }) }
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
    Add-OutputFile $path
}

# --- SELECT-AWARE SHOP SCRIPTS ---
# One set of two scripts per run (Sort-PlatesByThickness, Type-Plates). When a script starts it lists every
# transmittal + sequence found in the data and asks which ones are being nested NOW. Only plates in those are
# sorted / typed, and a plate's quantity counts only the selected transmittals + sequences. It then shows a
# preview (first 5, last 5 and the 5 highest quantities) to check against other documents before anything happens.
# $Keys: Index, TL, Seq, Parts, Qty, Text.  $Plates: Mark, Folder, Grade, Dets (KeyIndex, Qty per transmittal + sequence).
function New-ShopSelectionLines ($Title, $ProjectName, $JobNumber, $Keys, $Plates, $SelFile, [bool]$IsSort) {
    $n = @($Keys).Count
    $badChars = '[%!"]'
    $safeProjectHeader = Escape-BatchText $ProjectName
    $safeProject = Escape-BatchText ($ProjectName -replace '[&|<>^%]', '')
    $L = [System.Collections.Generic.List[string]]::new()
    $L.Add('@echo off')
    $L.Add(':: ========================================================')
    $L.Add(":: $Title")
    $L.Add(":: Project: $safeProjectHeader (Job $JobNumber)")
    $L.Add(':: Asks which transmittals / sequences are being nested now, previews the result, then only handles those.')
    $L.Add(':: ========================================================')
    $L.Add('setlocal enabledelayedexpansion')
    $L.Add('cd /d "%~dp0"')
    $L.Add(('set "SELFILE=%~dp0{0}"' -f ($SelFile -replace $badChars, '')))
    $L.Add('for /f %%a in (''echo prompt $E^| cmd'') do set "ESC=%%a"')
    $L.Add('set "CY=%ESC%[96m"')
    $L.Add('set "GR=%ESC%[92m"')
    $L.Add('set "YE=%ESC%[93m"')
    $L.Add('set "DG=%ESC%[90m"')
    $L.Add('set "RS=%ESC%[0m"')
    if ($IsSort) {
        $L.Add(":: --- Use the 'dxf' subfolder next to this script when there is one ---")
        $L.Add('set "WORK_DIR=%~dp0"')
        $L.Add('if exist "%~dp0dxf\" set "WORK_DIR=%~dp0dxf\"')
        $L.Add('cd /d "!WORK_DIR!"')
        $L.Add('set "SUFFIX=.dxf"')
    }
    else { $L.Add('set "SUFFIX="') }
    $L.Add('')
    $L.Add(":: --- Every transmittal + sequence in this run: name, sequence, plate marks, plate quantity ---")
    $L.Add("set ""KEY_COUNT=$n""")
    $L.Add("set ""TOTAL_MARKS=$(@($Plates).Count)""")
    foreach ($k in $Keys) {
        $i = $k.Index
        $kTl = $k.TL -replace $badChars, ''
        $kSeq = $k.Seq -replace $badChars, ''
        $kText = $k.Text -replace $badChars, ''
        $L.Add("set ""KEY_${i}_TL=$kTl""")
        $L.Add("set ""KEY_${i}_SEQ=$kSeq""")
        $L.Add("set ""KEY_${i}_QTY=$($k.Qty)""")
        $L.Add("set ""KEY_${i}_TEXT=$kText""")
    }
    $L.Add('')
    $L.Add(':select')
    $L.Add('echo.')
    $L.Add('echo %CY%========================================================%RS%')
    $L.Add("echo   %YE%$(Escape-BatchText $Title)%RS%")
    $L.Add("echo   Project: $safeProject (Job $JobNumber)")
    $L.Add('echo %CY%========================================================%RS%')
    $L.Add('echo.')
    $L.Add('echo   Which transmittals / sequences are you nesting NOW?')
    $L.Add('echo   Only these are handled, and a part''s quantity counts only these.')
    $L.Add('echo.')
    foreach ($k in $Keys) {
        $kText = Escape-BatchText ($k.Text -replace $badChars, '')
        $L.Add("echo   %GR%[$($k.Index)]%RS% $kText")
    }
    $L.Add('echo   %GR%[A]%RS% All of them')
    $L.Add('echo.')
    $L.Add('echo   Pick by number, TL name or sequence. Separate with spaces or commas.')
    $L.Add('echo   %DG%Example: 1 3     or     TL1296     or     5131%RS%')
    $L.Add('set "LAST="')
    $L.Add('if exist "%SELFILE%" set /p LAST=<"%SELFILE%"')
    $L.Add('if defined LAST echo   %DG%Last time: !LAST!  - press Enter to use it again%RS%')
    $L.Add('set "ANS="')
    $L.Add('set /p "ANS=%GR%Selection:%RS% "')
    $L.Add('if not defined ANS if defined LAST set "ANS=!LAST!"')
    $L.Add('if not defined ANS goto :select')
    $L.Add('set "ANS=!ANS:,= !"')
    $L.Add("for /l %%k in (1,1,$n) do set ""SEL_%%k=""")
    $L.Add('for %%t in (!ANS!) do call :pick "%%~t"')
    $L.Add('set "NSEL=0"')
    $L.Add('set "QSEL=0"')
    $L.Add("for /l %%k in (1,1,$n) do if defined SEL_%%k set /a NSEL+=1 & set /a QSEL+=!KEY_%%k_QTY!")
    $L.Add('if !NSEL!==0 (')
    $L.Add('    echo   %YE%Nothing matched. Try again.%RS%')
    $L.Add('    goto :select')
    $L.Add(')')
    $L.Add('>"%SELFILE%" echo !ANS!')
    $L.Add('echo.')
    $L.Add('echo   %YE%Nesting now:%RS%')
    $L.Add("for /l %%k in (1,1,$n) do if defined SEL_%%k echo     - !KEY_%%k_TEXT!")

    # Collect the plates that fall in the selection (A-Z), each with its quantity from the selected keys only
    $L.Add('set "PARTS=0"')
    foreach ($p in $Plates) {
        $mark = $p.Mark -replace $badChars, ''
        $grade = $p.Grade -replace $badChars, ''
        $keysText = ((ConvertTo-SendKeysText $p.Grade) -replace '!', '' -replace '"', '') -replace '%', '%%'
        $L.Add('set "HIT="')
        $L.Add('set /a Q=0')
        $L.Add('set "DETAIL="')
        foreach ($d in $p.Dets) {
            $L.Add("if defined SEL_$($d.KeyIndex) (")
            $L.Add('    set "HIT=1"')
            $L.Add("    set /a Q+=$($d.Qty)")
            $L.Add("    set ""DETAIL=!DETAIL! !KEY_$($d.KeyIndex)_TL!/!KEY_$($d.KeyIndex)_SEQ!:$($d.Qty)""")
            $L.Add(')')
        }
        $L.Add('if defined HIT (')
        $L.Add('    set /a PARTS+=1')
        $L.Add("    set ""P_!PARTS!_MARK=$mark""")
        $L.Add("    set ""P_!PARTS!_THK=$($p.Folder)""")
        $L.Add("    set ""P_!PARTS!_GRD=$grade""")
        $L.Add("    set ""P_!PARTS!_KEYS=$keysText""")
        $L.Add('    set "P_!PARTS!_QTY=!Q!"')
        $L.Add('    set "P_!PARTS!_DET=!DETAIL!"')
        $L.Add(')')
    }
    $L.Add('set /a OTHERS=TOTAL_MARKS-PARTS')
    $L.Add('echo.')
    $L.Add('echo   Plate marks in this selection: !PARTS!      Plate quantity: !QSEL!      Left alone: !OTHERS!')
    $L.Add('call :preview')
    $L.Add('echo.')
    $L.Add('set "OK="')
    $L.Add('set /p "OK=%GR%[Enter]%RS% continue   %GR%[R]%RS% pick again: "')
    $L.Add('if /i "!OK!"=="R" goto :select')
    $L.Add('')
    return $L
}

# Subroutines shared by both scripts. Placed after the main flow.
function New-ShopRoutineLines ($KeyCount, [bool]$IsSort) {
    $L = [System.Collections.Generic.List[string]]::new()
    $L.Add('exit /b 0')
    $L.Add('')
    $L.Add(':: --- Marks one answer (number, TL name, sequence or A) as selected ---')
    $L.Add(':pick')
    $L.Add('set "TOK=%~1"')
    $L.Add('if /i "!TOK!"=="A" (')
    $L.Add("    for /l %%k in (1,1,$KeyCount) do set ""SEL_%%k=1""")
    $L.Add('    exit /b 0')
    $L.Add(')')
    $L.Add('set "FOUND="')
    $L.Add("for /l %%k in (1,1,$KeyCount) do (")
    $L.Add('    if "!TOK!"=="%%k" (')
    $L.Add('        set "SEL_%%k=1"')
    $L.Add('        set "FOUND=1"')
    $L.Add('    )')
    $L.Add('    if /i "!TOK!"=="!KEY_%%k_TL!" (')
    $L.Add('        set "SEL_%%k=1"')
    $L.Add('        set "FOUND=1"')
    $L.Add('    )')
    $L.Add('    if "!TOK!"=="!KEY_%%k_SEQ!" (')
    $L.Add('        set "SEL_%%k=1"')
    $L.Add('        set "FOUND=1"')
    $L.Add('    )')
    $L.Add(')')
    $L.Add('if not defined FOUND echo   %YE%Not recognised, skipped: !TOK!%RS%')
    $L.Add('exit /b 0')
    $L.Add('')
    $L.Add(':: --- Preview: first 5, last 5 and the 5 highest quantities of the selection ---')
    $L.Add(':preview')
    $L.Add('echo.')
    $L.Add('echo   %YE%Preview - check these against your other documents%RS%')
    $L.Add('echo.')
    $L.Add('echo   %YE%First 5%RS%')
    $L.Add('call :rowhead')
    $L.Add('set /a E=5')
    $L.Add('if !PARTS! LSS 5 set /a E=PARTS')
    $L.Add('for /l %%i in (1,1,!E!) do call :row %%i')
    $L.Add('if !PARTS! GTR 5 (')
    $L.Add('    echo.')
    $L.Add('    echo   %YE%Last 5%RS%')
    $L.Add('    call :rowhead')
    $L.Add('    set /a S=PARTS-4')
    $L.Add('    if !S! LSS 6 set /a S=6')
    $L.Add('    for /l %%i in (!S!,1,!PARTS!) do call :row %%i')
    $L.Add(')')
    $L.Add('echo.')
    $L.Add('echo   %YE%5 highest quantities%RS%')
    $L.Add('call :rowhead')
    $L.Add('for /l %%i in (1,1,!PARTS!) do set "U_%%i="')
    $L.Add('set "N=0"')
    $L.Add(':topnext')
    $L.Add('set "BEST=0"')
    $L.Add('set "BQ=-1"')
    $L.Add('for /l %%i in (1,1,!PARTS!) do if not defined U_%%i if !P_%%i_QTY! GTR !BQ! (set "BQ=!P_%%i_QTY!" & set "BEST=%%i")')
    $L.Add('if !BEST!==0 exit /b 0')
    $L.Add('set "U_!BEST!=1"')
    $L.Add('call :row !BEST!')
    $L.Add('set /a N+=1')
    $L.Add('if !N! LSS 5 goto :topnext')
    $L.Add('exit /b 0')
    $L.Add('')
    $L.Add(':rowhead')
    $L.Add('echo   %DG%Mark                 Thick.        Qty  Grade%RS%')
    $L.Add('exit /b 0')
    $L.Add('')
    $L.Add(':row')
    $L.Add('set "R1=!P_%~1_MARK!!SUFFIX!                         "')
    $L.Add('set "R2=!P_%~1_THK!          "')
    $L.Add('set "R3=         !P_%~1_QTY!"')
    $L.Add('set "R5="')
    if ($IsSort) {
        $L.Add('if exist "!P_%~1_MARK!.dxf" (set "R5=%GR%found%RS%") else (set "R5=%YE%NOT FOUND%RS%")')
    }
    $L.Add('echo   !R1:~0,20! !R2:~0,8! !R3:~-8!  !P_%~1_GRD!  !R5!')
    $L.Add('exit /b 0')
    return $L
}

function New-TypeMacroLines ($Keys, $Plates, $ProjectName, $JobNumber, $SelFile) {
    $safeProject = Escape-BatchText ($ProjectName -replace '[&|<>^%]', '')
    $lines = [System.Collections.Generic.List[string]]@(New-ShopSelectionLines "Auto-Type Plate Quantities / Grades" $ProjectName $JobNumber $Keys $Plates $SelFile $false)
    $lines.Add(':choose')
    $lines.Add('echo.')
    $lines.Add('echo   What do you want to type?')
    $lines.Add('echo   %GR%[Q]%RS% Quantities')
    $lines.Add('echo   %GR%[G]%RS% Grades')
    $lines.Add('echo   %GR%[X]%RS% Exit')
    $lines.Add('set "WHAT="')
    $lines.Add('set /p "WHAT=%GR%Choice:%RS% "')
    $lines.Add('if /i "!WHAT!"=="X" exit /b 0')
    $lines.Add('if /i "!WHAT!"=="Q" (')
    $lines.Add('    set "KIND=Quantities"')
    $lines.Add('    set "NOUN=quantity"')
    $lines.Add('    goto :go')
    $lines.Add(')')
    $lines.Add('if /i "!WHAT!"=="G" (')
    $lines.Add('    set "KIND=Grades"')
    $lines.Add('    set "NOUN=grade"')
    $lines.Add('    goto :go')
    $lines.Add(')')
    $lines.Add('goto :choose')
    $lines.Add('')
    $lines.Add(':go')
    $lines.Add(":: ========================================================")
    $lines.Add(":: Instructions:")
    $lines.Add("::   1. Open your target software window.")
    $lines.Add("::   2. Click and select the FIRST cell at the top.")
    $lines.Add("::   3. Press any key here to start the 5-second countdown.")
    $lines.Add(":: ========================================================")
    $lines.Add('echo.')
    $lines.Add("echo ========================================================")
    $lines.Add('echo   Auto-Type !KIND! (A-Z order, selected only)')
    $lines.Add("echo   Project: $safeProject (Job $JobNumber)")
    $lines.Add('echo   Total items to type: !PARTS!')
    $lines.Add("echo ========================================================")
    $lines.Add("echo.")
    $lines.Add('echo WARNING: Position your cursor on the first !NOUN! cell NOW.')
    $lines.Add("pause")
    $lines.Add("")
    $lines.Add("echo Starting countdown...")
    $lines.Add("for /l %%i in (5,-1,1) do (")
    $lines.Add("    echo   %%i seconds remaining...")
    $lines.Add("    timeout /t 1 >nul")
    $lines.Add(")")
    $lines.Add('echo GO - typing !NOUN!s...')
    $lines.Add("powershell -NoProfile -Command `"`$wshell = New-Object -ComObject WScript.Shell; Start-Sleep -Milliseconds 300`"")
    $lines.Add('if "!KIND!"=="Quantities" (')
    $lines.Add('    for /l %%i in (1,1,!PARTS!) do call :typeqty %%i')
    $lines.Add(') else (')
    $lines.Add('    for /l %%i in (1,1,!PARTS!) do call :typegrade %%i')
    $lines.Add(')')
    $lines.Add('echo.')
    $lines.Add('echo ========================================================')
    $lines.Add('echo   Auto-Type !KIND! Complete!')
    $lines.Add('echo ========================================================')
    $lines.Add('pause')
    $lines.Add('goto :choose')
    $lines.Add('')
    foreach ($l in (New-ShopRoutineLines @($Keys).Count $false)) { $lines.Add($l) }
    $lines.Add('')
    $lines.Add(':typeqty')
    $lines.Add('echo %GR%[TYPING]%RS% Part !P_%~1_MARK! -^> Qty !P_%~1_QTY!   %DG%!P_%~1_DET!%RS%')
    $lines.Add("powershell -NoProfile -Command `"`$wshell = New-Object -ComObject WScript.Shell; `$wshell.SendKeys('!P_%~1_QTY!'); Start-Sleep -Milliseconds 250; `$wshell.SendKeys('{DOWN}'); Start-Sleep -Milliseconds 200`"")
    $lines.Add('exit /b 0')
    $lines.Add('')
    $lines.Add(':typegrade')
    $lines.Add('echo %GR%[TYPING]%RS% Part !P_%~1_MARK! -^> Grade !P_%~1_GRD!')
    $lines.Add("powershell -NoProfile -Command `"`$wshell = New-Object -ComObject WScript.Shell; `$wshell.SendKeys('!P_%~1_KEYS!'); Start-Sleep -Milliseconds 250; `$wshell.SendKeys('{DOWN}'); Start-Sleep -Milliseconds 200`"")
    $lines.Add('exit /b 0')
    return $lines
}

function New-SortScriptLines ($Keys, $Plates, $ProjectName, $JobNumber, $SelFile) {
    $lines = [System.Collections.Generic.List[string]]@(New-ShopSelectionLines "DXF Plate & Sheet Sorting Utility" $ProjectName $JobNumber $Keys $Plates $SelFile $true)
    $lines.Add('echo.')
    $lines.Add('echo   Working in: !WORK_DIR!')
    $lines.Add('echo.')
    $lines.Add('set "MOVED_COUNT=0"')
    $lines.Add('set "MISSING_COUNT=0"')
    $lines.Add('for /l %%i in (1,1,!PARTS!) do call :moveone %%i')
    $lines.Add('echo.')
    $lines.Add('echo ========================================================')
    $lines.Add('echo   Sorting Complete!')
    $lines.Add('echo   Successfully Moved: !MOVED_COUNT! file(s)')
    $lines.Add('echo   Missing / Not Found: !MISSING_COUNT! file(s)')
    $lines.Add('echo   Left alone (not in this selection): !OTHERS! mark(s)')
    $lines.Add('echo ========================================================')
    $lines.Add('echo.')
    $lines.Add('pause')
    foreach ($l in (New-ShopRoutineLines @($Keys).Count $true)) { $lines.Add($l) }
    $lines.Add('')
    $lines.Add(':moveone')
    $lines.Add('set "M=!P_%~1_MARK!"')
    $lines.Add('set "T=!P_%~1_THK!"')
    $lines.Add('if not exist "!T!\" mkdir "!T!" >nul 2>&1')
    $lines.Add('if exist "!M!.dxf" (')
    $lines.Add('    move /y "!M!.dxf" "!T!\" >nul 2>&1')
    $lines.Add('    echo %GR%[SUCCESS]%RS% !M!.dxf -^> !T!')
    $lines.Add('    set /a MOVED_COUNT+=1')
    $lines.Add(') else (')
    $lines.Add('    echo %YE%[NOT FOUND]%RS% !M!.dxf')
    $lines.Add('    set /a MISSING_COUNT+=1')
    $lines.Add(')')
    $lines.Add('exit /b 0')
    return $lines
}

# $PlateRows: Mark, Size, Grade, TL, Seq, Qty (one per plate line, TL = the transmittal it is placed under)
function Write-ShopScripts ($PlateRows, $ProjectName, $JobNumber, $Label) {
    if (@($PlateRows).Count -eq 0) { return $null }

    $keyGroups = @($PlateRows | Group-Object { "$($_.TL)|$($_.Seq)" } |
        Sort-Object { [double]('0' + ($_.Group[0].TL -replace '\D', '')) }, { $_.Group[0].TL }, { $_.Group[0].Seq })
    $keys = [System.Collections.Generic.List[PSCustomObject]]::new()
    $keyIndex = @{}
    $i = 0
    foreach ($g in $keyGroups) {
        $i++
        $first = $g.Group[0]
        $qty = [int](($g.Group | Measure-Object -Property Qty -Sum).Sum)
        $parts = @($g.Group | Select-Object -ExpandProperty Mark -Unique).Count
        $text = "{0,-9} seq {1,-6} {2,4} plate mark(s)   qty {3}" -f $first.TL, $first.Seq, $parts, $qty
        $keys.Add([PSCustomObject]@{ Index = $i; TL = $first.TL; Seq = $first.Seq; Parts = $parts; Qty = $qty; Text = $text })
        $keyIndex[$g.Name] = $i
    }

    $plates = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($g in ($PlateRows | Group-Object Mark | Sort-Object Name)) {
        $first = $g.Group[0]
        $grades = @($g.Group | ForEach-Object { "$($_.Grade)".Trim() } | Sort-Object -Unique)
        $sizes = @($g.Group | ForEach-Object { "$($_.Size)".Trim() } | Sort-Object -Unique)
        if ($grades.Count -gt 1 -or $sizes.Count -gt 1) {
            Write-Log "WARNING: plate $($first.Mark) has different sizes/grades in different transmittals ($($sizes -join ' / ') ; $($grades -join ' / ')); the first one is used for thickness and grade." "Yellow"
        }
        $dets = @($g.Group | Group-Object { "$($_.TL)|$($_.Seq)" } | ForEach-Object {
            [PSCustomObject]@{ KeyIndex = $keyIndex[$_.Name]; Qty = [int](($_.Group | Measure-Object -Property Qty -Sum).Sum) }
        } | Sort-Object KeyIndex)
        $plates.Add([PSCustomObject]@{ Mark = $first.Mark; Folder = (Format-ThicknessFolder (Get-DecimalThickness $first.Size)); Grade = $first.Grade; Dets = $dets })
    }

    Write-Host ""
    Write-Host "Plates per transmittal + sequence (the shop scripts ask which of these you are nesting):"
    foreach ($k in $keys) { Write-Host ("  [{0}] {1}" -f $k.Index, $k.Text) }

    $sortName  = "Sort-PlatesByThickness - $Label.bat"
    $typeName  = "Type-Plates - $Label.bat"
    $selFile   = "Shop-Selection - $Label.txt"
    Save-BatFile $sortName (New-SortScriptLines $keys $plates $ProjectName $JobNumber $selFile)
    Save-BatFile $typeName (New-TypeMacroLines $keys $plates $ProjectName $JobNumber $selFile)

    return [PSCustomObject]@{ Label = $Label; Sort = $sortName; Type = $typeName; Keys = $keys.Count }
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


# =========================================================================
# UPDATE A JOB THAT WAS ALREADY BUILT
# Reads the previous BOM workbook (and the small "last build" record saved next to it), compares it with what this run
# produces from the .kss files and the Required Assm workbook, shows what is new / changed / removed, and keeps or
# replaces the values that were edited by hand in the old workbook. Everything downstream (folder labels, work orders,
# QPR, Exceptions, Statistics, shop scripts) is then rebuilt by the normal run from the merged values.
# =========================================================================

# Text used to compare two cell values: trimmed, one space between words, no case, numbers written the same way.
function ConvertTo-CmpText ($v) {
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    if ($null -eq $v) { return "" }
    if ($v -is [double] -or $v -is [decimal] -or $v -is [int] -or $v -is [long]) { return ([double]$v).ToString("0.######", $inv) }
    $t = (("$v").Trim()) -replace '\s+', ' '
    $d = 0.0
    if ($t -and [double]::TryParse($t, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d)) { return $d.ToString("0.######", $inv) }
    return $t.ToLowerInvariant()
}

# A cell's value as plain text (numbers written the same way on every PC; a date number is NOT converted).
function Get-CellText ($v) {
    if ($null -eq $v) { return "" }
    if ($v -is [double]) { return $v.ToString("0.######", [System.Globalization.CultureInfo]::InvariantCulture) }
    return ("$v").Trim()
}

# Which of the three states an editable value is in, given the value in the OLD workbook ($Old), the value the inputs
# (Required Assm) give now ($New) and, when the "last build" record exists, the value the inputs gave then ($Seed).
#   Same      nothing to decide
#   TakeNew   the inputs changed and nobody edited the old workbook here -> use the new value
#   KeepOld   somebody edited the old workbook and the inputs did not change -> keep the edit
#   Conflict  both changed, or there is no record to tell which one changed -> ask
function Get-EditState {
    param($Old, $New, $Seed, [bool]$HasSeed)
    $o = ConvertTo-CmpText $Old; $n = ConvertTo-CmpText $New
    if ($o -eq $n) { return 'Same' }
    if ($HasSeed) {
        $s = ConvertTo-CmpText $Seed
        if ($o -eq $s) { return 'TakeNew' }
        if ($n -eq $s) { return 'KeepOld' }
        return 'Conflict'
    }
    if ($o -eq '') { return 'TakeNew' }      # the old workbook had nothing there, so there is no edit to lose
    return 'Conflict'
}

# Saved as a hidden file with its own extension, so it is never mistaken for the KISS (.kss) file.
function Get-KissStatePath ([string]$WorkbookPath) { return [System.IO.Path]::ChangeExtension($WorkbookPath, '.kiss-state') }

# The "last build" record saved next to every workbook: which .kss files were used and what the inputs said then.
function Read-KissState ([string]$WorkbookPath) {
    $p = Get-KissStatePath $WorkbookPath
    if (-not (Test-Path -LiteralPath $p -ErrorAction SilentlyContinue)) {
        $p = [System.IO.Path]::ChangeExtension($WorkbookPath, '.kiss.json')      # the first version of the record used this name
        if (-not (Test-Path -LiteralPath $p)) { return $null }
    }
    try { $j = [System.IO.File]::ReadAllText($p) | ConvertFrom-Json } catch { return $null }
    if (-not $j -or -not $j.Version) { return $null }
    $seq = @{}; $exc = @{}; $set = @{}; $sched = @{}; $files = @{}
    foreach ($s in @($j.Seq)) { if ($s) { $seq[[string]$s.K] = @{ Due = [string]$s.Due; Paint = [string]$s.Paint; Expl = [string]$s.Expl } } }
    foreach ($s in @($j.Exc)) { if ($s) { $exc[[string]$s.K] = @{ Due = [string]$s.Due; Paint = [string]$s.Paint; Expl = [string]$s.Expl } } }
    foreach ($s in @($j.Settings)) { if ($s) { $set[[string]$s.K] = [string]$s.V } }
    foreach ($s in @($j.Sched)) { if ($s) { $sched[[string]$s.K] = [string]$s.V } }
    foreach ($s in @($j.Files)) { if ($s) { $files[[string]$s.N] = [string]$s.H } }
    return [PSCustomObject]@{
        Built = [string]$j.Built; Merge = $j.MergeDuplicates; Qpr = [bool]$j.Qpr; QprBy = [string]$j.QprBy
        Seq = $seq; Exc = $exc; Settings = $set; Sched = $sched; Files = $files
    }
}

# Writes the record next to the workbook. $Fresh = what the inputs said this run (before any value from an old workbook was kept).
function Save-KissState {
    param([string]$WorkbookPath, $Fresh, $KssFiles, [bool]$Merge, [bool]$Qpr, [string]$QprBy)
    $files = @(foreach ($f in @($KssFiles)) {
        $h = ""
        try { $h = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash } catch { }
        @{ N = "$($f.Name)"; H = $h }
    })
    $state = [ordered]@{
        Version = 1
        Built = (Get-Date).ToString('s')
        Workbook = [System.IO.Path]::GetFileName($WorkbookPath)
        MergeDuplicates = $Merge
        Qpr = $Qpr
        QprBy = $QprBy
        Files = $files
        Settings = @(foreach ($k in $Fresh.Settings.Keys) { @{ K = $k; V = "$($Fresh.Settings[$k])" } })
        Seq = @(foreach ($k in $Fresh.Seq.Keys) { @{ K = $k; Due = $Fresh.Seq[$k].Due; Paint = $Fresh.Seq[$k].Paint; Expl = $Fresh.Seq[$k].Expl } })
        Exc = @(foreach ($k in $Fresh.Exc.Keys) { @{ K = $k; Due = $Fresh.Exc[$k].Due; Paint = $Fresh.Exc[$k].Paint; Expl = $Fresh.Exc[$k].Expl } })
        Sched = @(foreach ($k in $Fresh.Sched.Keys) { @{ K = $k; V = "$($Fresh.Sched[$k])" } })
    }
    $json = $state | ConvertTo-Json -Depth 5
    $statePath = Get-KissStatePath $WorkbookPath
    if (Test-Path -LiteralPath $statePath) { (Get-Item -LiteralPath $statePath -Force).Attributes = 'Normal' }
    [System.IO.File]::WriteAllText($statePath, $json, (New-Object System.Text.UTF8Encoding($false)))
    try { (Get-Item -LiteralPath $statePath -Force).Attributes = 'Hidden' } catch { }
    $legacy = [System.IO.Path]::ChangeExtension($WorkbookPath, '.kiss.json')
    if (Test-Path -LiteralPath $legacy) { Remove-Item -LiteralPath $legacy -Force -ErrorAction SilentlyContinue }
}

# What the inputs say right now for every value that can be edited by hand in the workbook.
function Get-FreshSeeds {
    param($SettingsList, $SequenceRows, $ExceptionRows)
    $seq = [ordered]@{}; $exc = [ordered]@{}; $set = [ordered]@{}; $sched = [ordered]@{}
    foreach ($s in @($SettingsList)) { $set[$s[0]] = Get-CellText $s[2] }
    foreach ($r in @($SequenceRows)) {
        $seq["$($r.TL)|$($r.Sequence)"] = @{ Due = "$($r.DueDate)"; Paint = "$($r.Paint)"; Expl = "$($r.Explanation)" }
        $sched["$($r.TL) / $($r.Sequence)"] = "$($r.DueDate)"
    }
    foreach ($e in @($ExceptionRows)) { $exc["$($e.Assembly)"] = @{ Due = "$($e.DueDate)"; Paint = "$($e.Paint)"; Expl = "$($e.Explanation)" } }
    return @{ Settings = $set; Seq = $seq; Exc = $exc; Sched = $sched }
}

function Get-SheetValues ($Sheet) {
    try { $last = $Sheet.Cells.SpecialCells(11) } catch { return $null }
    $v = $Sheet.Range($Sheet.Cells.Item(1, 1), $last).Value2
    if ($v -is [System.Array] -and $v.Rank -eq 2) { return $v }
    return $null
}

function Get-HeaderColumns ($Vals) {
    $map = @{}
    $r0 = $Vals.GetLowerBound(0)
    for ($c = $Vals.GetLowerBound(1); $c -le $Vals.GetUpperBound(1); $c++) {
        $h = ("$($Vals[$r0, $c])").Trim().ToLowerInvariant()
        if ($h -and -not $map.ContainsKey($h)) { $map[$h] = $c }
    }
    return $map
}

# Opens the previous BOM workbook read-only in an Excel of its own and returns what the update needs from it:
# the part rows of the Master BOM, the values typed on Settings / Sequence Notes / Assembly Exceptions / the Statistics schedule.
function Read-PreviousBom {
    param([string]$Path, $SettingNames)

    $res = [PSCustomObject]@{
        Path = $Path; Built = (Get-Item -LiteralPath $Path).LastWriteTime
        Settings = @{}; SeqNotes = @{}; Exceptions = @{}; Schedule = @{}
        Bom = [System.Collections.Generic.List[object]]::new(); SourceFiles = @(); Sheets = @()
    }
    $xl = $null; $wb = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($Path, 0, $true)
        $res.Sheets = @(foreach ($ws in $wb.Worksheets) { $ws.Name })

        foreach ($n in @($SettingNames)) {
            try { $res.Settings[$n] = $wb.Names.Item($n).RefersToRange.Value2 } catch { }
        }

        if ($res.Sheets -contains 'Master BOM') {
            $v = Get-SheetValues $wb.Worksheets.Item('Master BOM')
            if ($v) {
                $h = Get-HeaderColumns $v
                $cA = $h['assembly']; $cP = $h['partmark']
                if ($cA -and $cP) {
                    $get = { param($r, $name) if ($h.ContainsKey($name)) { Get-CellText ($v[$r, $h[$name]]) } else { "" } }
                    $src = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    for ($r = $v.GetLowerBound(0) + 1; $r -le $v.GetUpperBound(0); $r++) {
                        $asm = Get-CellText ($v[$r, $cA])
                        if (-not $asm) { continue }          # blank row or the totals row
                        $q = 0.0
                        [void][double]::TryParse((& $get $r 'qty'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$q)
                        $tl = & $get $r 'transmittal'; if (-not $tl) { $tl = & $get $r 'tl #' }
                        $res.Bom.Add([PSCustomObject]@{
                            Asm = $asm; Part = (Get-CellText ($v[$r, $cP])); Qty = $q
                            Shape = (& $get $r 'shape'); Size = (& $get $r 'size'); Len = (& $get $r 'length'); Grade = (& $get $r 'grade')
                            Tl = $tl; Seq = (& $get $r 'sequence'); Status = (& $get $r 'status')
                        })
                        $s = & $get $r 'sourcefile'
                        if ($s) { [void]$src.Add($s) }
                    }
                    $res.SourceFiles = @($src)
                }
            }
        }

        if ($res.Sheets -contains 'Sequence Notes') {
            $v = Get-SheetValues $wb.Worksheets.Item('Sequence Notes')
            if ($v) {
                $h = Get-HeaderColumns $v
                if ($h['transmittal'] -and $h['sequence']) {
                    for ($r = $v.GetLowerBound(0) + 1; $r -le $v.GetUpperBound(0); $r++) {
                        $tl = Get-CellText ($v[$r, $h['transmittal']]); $sq = Get-CellText ($v[$r, $h['sequence']])
                        if (-not $tl -or -not $sq) { continue }
                        $res.SeqNotes["$tl|$sq"] = @{
                            Due = $(if ($h['due date']) { ConvertTo-DueText ($v[$r, $h['due date']]) } else { "" })
                            Paint = $(if ($h['paint / notes']) { Get-CellText ($v[$r, $h['paint / notes']]) } else { "" })
                            Expl = $(if ($h['explanation']) { Get-CellText ($v[$r, $h['explanation']]) } else { "" })
                        }
                    }
                }
            }
        }

        if ($res.Sheets -contains 'Assembly Exceptions') {
            $v = Get-SheetValues $wb.Worksheets.Item('Assembly Exceptions')
            if ($v) {
                $h = Get-HeaderColumns $v
                if ($h['assembly']) {
                    for ($r = $v.GetLowerBound(0) + 1; $r -le $v.GetUpperBound(0); $r++) {
                        $asm = (Get-CellText ($v[$r, $h['assembly']])) -replace '\D', ''
                        if ($asm.Length -lt 4) { continue }
                        $res.Exceptions[$asm] = @{
                            Due = $(if ($h['override due date']) { ConvertTo-DueText ($v[$r, $h['override due date']]) } else { "" })
                            Paint = $(if ($h['override paint / notes']) { Get-CellText ($v[$r, $h['override paint / notes']]) } else { "" })
                            Expl = $(if ($h['override explanation']) { Get-CellText ($v[$r, $h['override explanation']]) } else { "" })
                        }
                    }
                }
            }
        }

        if ($res.Sheets -contains 'Statistics') {
            $v = Get-SheetValues $wb.Worksheets.Item('Statistics')
            if ($v -and $v.GetUpperBound(1) -ge 16) {
                for ($r = $v.GetLowerBound(0); $r -le $v.GetUpperBound(0); $r++) {
                    $label = (Get-CellText ($v[$r, 15]))
                    if ($label -match '^(TLR?\d+)\s*/\s*(\d{4})$') { $res.Schedule["$($Matches[1]) / $($Matches[2])"] = ConvertTo-DueText ($v[$r, 16]) }
                }
            }
        }
    }
    finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($xl) {
            try { $xl.Quit() } catch { }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        }
    }
    return $res
}

# One row per part of the new Master BOM, in the same shape Read-PreviousBom gives for the old one.
function Get-NewBomRows {
    param($MasterData, $AsmByKey, $TlNames)
    foreach ($item in @($MasterData)) {
        $a = $AsmByKey["$($item.TL)|$($item.Assembly)"]
        $tl = if ($a -and $TlNames.ContainsKey($a.TlNumber)) { $TlNames[$a.TlNumber] } elseif ($a) { "TL$($a.TlNumber)" } else { "$($item.TL)" }
        $seq = if ($a) { $a.Sequence } else { "$($item.Assembly)".Substring(0, [Math]::Min(4, "$($item.Assembly)".Length)) }
        [PSCustomObject]@{
            Asm = "$($item.Assembly)".Trim(); Part = "$($item.PartMark)"; Qty = [double]$item.Qty
            Shape = "$($item.Shape)"; Size = "$($item.Size)"; Len = "$($item.Length)"; Grade = "$($item.Grade)"
            Tl = $tl; Seq = "$seq"; Status = ""
        }
    }
}

# Compares the old Master BOM rows with the new ones, assembly by assembly.
# $StillRequired = assemblies the Required Assm workbook still lists (an old assembly missing from the new data but still listed
# there means its .kss was not selected, not that it was dropped).
function Get-BomChanges {
    param($OldRows, $NewRows, $StillRequired)

    $group = {
        param($rows)
        $d = [ordered]@{}
        foreach ($r in $rows) {
            $k = "$($r.Asm)".Trim()
            if (-not $d.Contains($k)) { $d[$k] = @{ Tl = "$($r.Tl)"; Seq = "$($r.Seq)"; Parts = [ordered]@{} } }
            $pk = ConvertTo-CmpText $r.Part
            $p = $d[$k].Parts
            if (-not $p.Contains($pk)) { $p[$pk] = @{ Name = "$($r.Part)"; Qty = 0.0; Spec = (@($r.Shape, $r.Size, $r.Len, $r.Grade | ForEach-Object { ConvertTo-CmpText $_ }) -join '|'); Show = "$($r.Shape) $($r.Size) x $($r.Len) $($r.Grade)".Trim() } }
            $p[$pk].Qty += [double]$r.Qty
        }
        return $d
    }
    $old = & $group @($OldRows)
    $new = & $group @($NewRows)

    $added = [System.Collections.Generic.List[object]]::new()
    $removed = [System.Collections.Generic.List[object]]::new()
    $moved = [System.Collections.Generic.List[object]]::new()
    $changed = [System.Collections.Generic.List[object]]::new()
    $changedParts = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $unchanged = 0
    $fmt = { param($q) ([double]$q).ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture) }

    foreach ($k in $new.Keys) {
        $n = $new[$k]
        if (-not $old.Contains($k)) { $added.Add([PSCustomObject]@{ Asm = $k; Tl = $n.Tl; Seq = $n.Seq; Parts = $n.Parts.Count }); continue }
        $o = $old[$k]
        $isMoved = (ConvertTo-CmpText $o.Tl) -ne (ConvertTo-CmpText $n.Tl)
        if ($isMoved) { $moved.Add([PSCustomObject]@{ Asm = $k; FromTl = $o.Tl; FromSeq = $o.Seq; Tl = $n.Tl; Seq = $n.Seq }) }
        $details = [System.Collections.Generic.List[string]]::new()
        foreach ($pk in $n.Parts.Keys) {
            $np = $n.Parts[$pk]
            if (-not $o.Parts.Contains($pk)) { $details.Add("part $($np.Name) added (qty $(& $fmt $np.Qty), $($np.Show))"); continue }
            $op = $o.Parts[$pk]
            $bits = @()
            if ([Math]::Abs($op.Qty - $np.Qty) -gt 0.0001) { $bits += "qty $(& $fmt $op.Qty) -> $(& $fmt $np.Qty)" }
            if ($op.Spec -ne $np.Spec) { $bits += "$($op.Show) -> $($np.Show)" }
            if ($bits.Count -gt 0) { $details.Add("part $($np.Name): $($bits -join '; ')"); [void]$changedParts.Add("$k|$($np.Name)") }
        }
        foreach ($pk in $o.Parts.Keys) { if (-not $n.Parts.Contains($pk)) { $details.Add("part $($o.Parts[$pk].Name) removed (was qty $(& $fmt $o.Parts[$pk].Qty))") } }
        if ($details.Count -gt 0) { $changed.Add([PSCustomObject]@{ Asm = $k; Tl = $n.Tl; Seq = $n.Seq; Details = @($details) }) }
        elseif (-not $isMoved) { $unchanged++ }
    }
    foreach ($k in $old.Keys) {
        if ($new.Contains($k)) { continue }
        $o = $old[$k]
        $why = if ($StillRequired -and $StillRequired.Contains($k)) { 'NoData' } else { 'Dropped' }
        $removed.Add([PSCustomObject]@{ Asm = $k; Tl = $o.Tl; Seq = $o.Seq; Why = $why })
    }

    # Folders (one per TL + sequence) that need new or reprinted paperwork
    $oldFolders = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $newFolders = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($k in $old.Keys) { [void]$oldFolders.Add("$($old[$k].Tl) / $($old[$k].Seq)") }
    foreach ($k in $new.Keys) { [void]$newFolders.Add("$($new[$k].Tl) / $($new[$k].Seq)") }
    $folders = [ordered]@{}
    $mark = {
        param($f, $isNew)
        if (-not $folders.Contains($f)) { $folders[$f] = if ($isNew) { 'NEW' } else { 'CHANGED' } }
    }
    foreach ($a in $added) { $f = "$($a.Tl) / $($a.Seq)"; & $mark $f (-not $oldFolders.Contains($f)) }
    foreach ($a in $changed) { $f = "$($a.Tl) / $($a.Seq)"; & $mark $f (-not $oldFolders.Contains($f)) }
    foreach ($a in $moved) {
        $f = "$($a.Tl) / $($a.Seq)"; & $mark $f (-not $oldFolders.Contains($f))
        $from = "$($a.FromTl) / $($a.FromSeq)"
        if ($newFolders.Contains($from)) { & $mark $from $false } elseif (-not $folders.Contains($from)) { $folders[$from] = 'GONE' }
    }
    foreach ($a in $removed) {
        $from = "$($a.Tl) / $($a.Seq)"
        if ($newFolders.Contains($from)) { & $mark $from $false } elseif (-not $folders.Contains($from)) { $folders[$from] = 'GONE' }
    }
    $oldTls = @($old.Keys | ForEach-Object { $old[$_].Tl } | Select-Object -Unique)
    $newTls = @($new.Keys | ForEach-Object { $new[$_].Tl } | Select-Object -Unique)

    return [PSCustomObject]@{
        Added = @($added); Removed = @($removed); Moved = @($moved); Changed = @($changed); Unchanged = $unchanged
        ChangedParts = $changedParts; Folders = $folders
        NewTls = @($newTls | Where-Object { $oldTls -notcontains $_ }); GoneTls = @($oldTls | Where-Object { $newTls -notcontains $_ })
        OldAssemblies = $old.Count; NewAssemblies = $new.Count
    }
}

# Every editable value, old vs new, with its state (see Get-EditState). Settings, Sequence Notes, Assembly Exceptions, Statistics schedule.
function Get-EditItems {
    param($Prev, $PrevState, $SettingsList, $SequenceRows, $ExceptionRows, $AsmDigits)

    $items = [System.Collections.Generic.List[object]]::new()
    $hasState = [bool]$PrevState
    $add = {
        param($Group, $Key, $Field, $Label, $Old, $New, $Seed, $HasSeed, $Force)
        $state = Get-EditState -Old $Old -New $New -Seed $Seed -HasSeed $HasSeed
        if ($Force -and $state -ne 'Same') { $state = 'Conflict' }
        $items.Add([PSCustomObject]@{ Group = $Group; Key = $Key; Field = $Field; Label = $Label; Old = $Old; New = $New; State = $state; Use = $(if ($state -eq 'TakeNew') { 'New' } else { 'Old' }) })
    }

    foreach ($s in @($SettingsList)) {
        $n = $s[0]
        if (-not $Prev.Settings.ContainsKey($n)) { continue }
        $seed = if ($hasState -and $PrevState.Settings.ContainsKey($n)) { $PrevState.Settings[$n] } else { $null }
        $hasSeed = $hasState -and $PrevState.Settings.ContainsKey($n)
        # The issue date always moves with the day the update is made, so it is always the user's call
        & $add 'Settings' $n $n $s[1] $Prev.Settings[$n] $s[2] $seed $hasSeed ($n -eq 'Set_IssuedDate')
    }
    foreach ($r in @($SequenceRows)) {
        $key = "$($r.TL)|$($r.Sequence)"
        if (-not $Prev.SeqNotes.ContainsKey($key)) { continue }      # a new folder: nothing to keep
        $o = $Prev.SeqNotes[$key]
        $sd = if ($hasState -and $PrevState.Seq.ContainsKey($key)) { $PrevState.Seq[$key] } else { $null }
        $hs = [bool]$sd
        $where = "$($r.TL) / $($r.Sequence)"
        & $add 'Sequence Notes' $key 'Due' "$where  Due date" $o.Due $r.DueDate $(if ($hs) { $sd.Due } else { $null }) $hs $false
        & $add 'Sequence Notes' $key 'Paint' "$where  Paint / Notes" $o.Paint $r.Paint $(if ($hs) { $sd.Paint } else { $null }) $hs $false
        & $add 'Sequence Notes' $key 'Expl' "$where  Explanation" $o.Expl $r.Explanation $(if ($hs) { $sd.Expl } else { $null }) $hs $false
        $sk = "$($r.TL) / $($r.Sequence)"
        # The Statistics schedule due date is asked about on its own only when it was edited apart from the Sequence Notes one
        if ($Prev.Schedule.ContainsKey($sk) -and (ConvertTo-CmpText $Prev.Schedule[$sk]) -ne (ConvertTo-CmpText $o.Due)) {
            $ss = if ($hasState -and $PrevState.Sched.ContainsKey($sk)) { $PrevState.Sched[$sk] } else { $null }
            & $add 'Statistics schedule' $sk 'Due' "$where  Due date (Statistics schedule)" $Prev.Schedule[$sk] $r.DueDate $ss ($hasState -and $PrevState.Sched.ContainsKey($sk)) $false
        }
    }
    # Exceptions: the new ones plus any ASSM the old workbook gave an override that the Exceptions tab no longer lists (still in this run)
    $freshByAsm = @{}
    foreach ($e in @($ExceptionRows)) { $freshByAsm["$($e.Assembly)"] = $e }
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($k in $freshByAsm.Keys) { $keys.Add($k) }
    foreach ($k in $Prev.Exceptions.Keys) { if (-not $freshByAsm.ContainsKey($k) -and $AsmDigits.Contains($k)) { $keys.Add($k) } }
    foreach ($k in ($keys | Sort-Object)) {
        $o = if ($Prev.Exceptions.ContainsKey($k)) { $Prev.Exceptions[$k] } else { @{ Due = ""; Paint = ""; Expl = "" } }
        $e = $freshByAsm[$k]
        $nv = if ($e) { @{ Due = "$($e.DueDate)"; Paint = "$($e.Paint)"; Expl = "$($e.Explanation)" } } else { @{ Due = ""; Paint = ""; Expl = "" } }
        $sd = if ($hasState -and $PrevState.Exc.ContainsKey($k)) { $PrevState.Exc[$k] } elseif ($hasState) { @{ Due = ""; Paint = ""; Expl = "" } } else { $null }
        $hs = [bool]$sd
        & $add 'Assembly Exceptions' $k 'Due' "ASSM $k  Override due date" $o.Due $nv.Due $(if ($hs) { $sd.Due } else { $null }) $hs $false
        & $add 'Assembly Exceptions' $k 'Paint' "ASSM $k  Override paint / notes" $o.Paint $nv.Paint $(if ($hs) { $sd.Paint } else { $null }) $hs $false
        & $add 'Assembly Exceptions' $k 'Expl' "ASSM $k  Override explanation" $o.Expl $nv.Expl $(if ($hs) { $sd.Expl } else { $null }) $hs $false
    }
    return $items
}

# A value as shown to the user (the issue date is stored as a date number)
function Format-EditValue ($Item, $Value) {
    if ($Item.Key -eq 'Set_IssuedDate' -and "$Value" -match '^\d+(\.\d+)?$') { return [DateTime]::FromOADate([double]$Value).ToString('M/d/yy') }
    return (Get-CellText $Value)
}

function Show-UpdateLines ($Lines, [int]$Max = 12, [string]$Color = 'Gray') {
    $n = 0
    foreach ($l in $Lines) {
        if ($n -ge $Max) { Write-Host "      ... and $(@($Lines).Count - $Max) more (all of them are in the report)" -ForegroundColor DarkGray; break }
        Write-Host "      $l" -ForegroundColor $Color
        $n++
    }
}

# Shows what changed, asks about every value that changed on both sides, applies the answers to the rows the run is about to
# write, and saves a plain-text report. Returns Proceed, the Master BOM status text to carry over, the Settings values and
# schedule due dates to keep, the (possibly extended) exception rows, and the rows for the "Update Log" tab.
function Invoke-UpdateReview {
    param($Prev, $PrevState, $NewRows, $StillRequired, $SettingsList, $SequenceRows, $ExceptionRows, $AsmRows, $TlNames,
          $KssFiles, $KssNotes, $MissingAssemblies, [string]$ReportFolder, [string]$JobLabel)

    $log = [System.Collections.Generic.List[object]]::new()        # Change, Folder, Assembly, Detail
    $rep = [System.Collections.Generic.List[string]]::new()
    $say = { param($t) $rep.Add($t) }
    $when = Get-Date

    $ch = Get-BomChanges -OldRows $Prev.Bom -NewRows $NewRows -StillRequired $StillRequired

    Write-Section "What changed since the previous workbook"
    Write-Host "  Previous: $([System.IO.Path]::GetFileName($Prev.Path))  (saved $($Prev.Built.ToString('M/d/yy h:mm tt')))" -ForegroundColor DarkGray
    if (-not $PrevState) { Write-Host "  No 'last build' record next to it (first update of this job): values that differ are shown for you to choose." -ForegroundColor DarkGray }
    & $say "UPDATE REPORT  $JobLabel"
    & $say "Made $($when.ToString('M/d/yyyy h:mm tt'))   Previous workbook: $([System.IO.Path]::GetFileName($Prev.Path))"
    & $say ""

    # --- source files ---
    Write-Host "  KISS files" -ForegroundColor Yellow
    & $say "KISS FILES"
    foreach ($f in @($KssFiles)) {
        $note = if ($KssNotes.ContainsKey($f.Name)) { $KssNotes[$f.Name] } else { '' }
        $line = "{0,-12} {1}" -f $(if ($note) { "[$note]" } else { '' }), $f.Name
        Write-Host "    $line" -ForegroundColor $(if ($note -match 'NEW|CHANGED') { 'Yellow' } else { 'Gray' })
        & $say "  $line"
    }
    $unusedBefore = @($Prev.SourceFiles | Where-Object { $n = $_; -not (@($KssFiles | Where-Object { $_.Name -ieq $n })) })
    foreach ($u in $unusedBefore) {
        Write-Host "    [NOT SELECTED] $u  (used by the previous workbook)" -ForegroundColor Yellow
        & $say "  [NOT SELECTED] $u (used by the previous workbook)"
    }
    & $say ""

    # --- assemblies ---
    Write-Host ""
    Write-Host "  Assemblies: $($ch.OldAssemblies) before, $($ch.NewAssemblies) now" -ForegroundColor Yellow
    & $say "ASSEMBLIES: $($ch.OldAssemblies) before, $($ch.NewAssemblies) now, $($ch.Unchanged) unchanged"
    if ($ch.NewTls.Count -gt 0) { Write-Host "    New transmittal(s): $($ch.NewTls -join ', ')" -ForegroundColor Green; & $say "  New transmittal(s): $($ch.NewTls -join ', ')" }
    if ($ch.GoneTls.Count -gt 0) { Write-Host "    Transmittal(s) no longer in the BOM: $($ch.GoneTls -join ', ')" -ForegroundColor Red; & $say "  Transmittal(s) no longer in the BOM: $($ch.GoneTls -join ', ')" }

    $groupLine = {
        param($title, $color, $count, $lines)
        Write-Host ("    {0,-34}{1}" -f $title, $count) -ForegroundColor $color
        Show-UpdateLines $lines 10 $color
    }
    $addLines = @($ch.Added | ForEach-Object { "$($_.Asm)   $($_.Tl) / $($_.Seq)   $($_.Parts) part(s)" })
    $movLines = @($ch.Moved | ForEach-Object { "$($_.Asm)   $($_.FromTl) / $($_.FromSeq)  ->  $($_.Tl) / $($_.Seq)" })
    $chgLines = @($ch.Changed | ForEach-Object { "$($_.Asm)   $($_.Tl) / $($_.Seq)   $($_.Details.Count) difference(s): $($_.Details[0])" })
    $remDrop = @($ch.Removed | Where-Object { $_.Why -eq 'Dropped' })
    $remNoData = @($ch.Removed | Where-Object { $_.Why -eq 'NoData' })
    $remDropLines = @($remDrop | ForEach-Object { "$($_.Asm)   was $($_.Tl) / $($_.Seq)" })
    $remNoDataLines = @($remNoData | ForEach-Object { "$($_.Asm)   was $($_.Tl) / $($_.Seq)" })
    if ($ch.Added.Count) { & $groupLine "New assemblies" 'Green' $ch.Added.Count $addLines }
    if ($ch.Moved.Count) { & $groupLine "Moved to another TL (REV)" 'Cyan' $ch.Moved.Count $movLines }
    if ($ch.Changed.Count) { & $groupLine "Changed (qty / size / parts)" 'Yellow' $ch.Changed.Count $chgLines }
    if ($remDrop.Count) { & $groupLine "Dropped from the Required Assm" 'Red' $remDrop.Count $remDropLines }
    if ($remNoData.Count) { & $groupLine "STILL REQUIRED, NO .KSS DATA" 'Red' $remNoData.Count $remNoDataLines }
    if (-not ($ch.Added.Count + $ch.Moved.Count + $ch.Changed.Count + $remDrop.Count + $remNoData.Count)) { Write-Host "    No part differences found." -ForegroundColor Green }

    foreach ($a in $ch.Added) { $log.Add([PSCustomObject]@{ Change = 'NEW assembly'; Folder = "$($a.Tl) / $($a.Seq)"; Assembly = $a.Asm; Detail = "$($a.Parts) part(s)" }) }
    foreach ($a in $ch.Moved) { $log.Add([PSCustomObject]@{ Change = 'MOVED (REV)'; Folder = "$($a.Tl) / $($a.Seq)"; Assembly = $a.Asm; Detail = "was $($a.FromTl) / $($a.FromSeq)" }) }
    foreach ($a in $ch.Changed) { foreach ($d in $a.Details) { $log.Add([PSCustomObject]@{ Change = 'CHANGED'; Folder = "$($a.Tl) / $($a.Seq)"; Assembly = $a.Asm; Detail = $d }) } }
    foreach ($a in $remDrop) { $log.Add([PSCustomObject]@{ Change = 'DROPPED'; Folder = "$($a.Tl) / $($a.Seq)"; Assembly = $a.Asm; Detail = 'no longer listed in the Required Assm' }) }
    foreach ($a in $remNoData) { $log.Add([PSCustomObject]@{ Change = 'NO KSS DATA'; Folder = "$($a.Tl) / $($a.Seq)"; Assembly = $a.Asm; Detail = 'still listed in the Required Assm but not found in the selected .kss files' }) }

    if ($ch.Added.Count) { & $say "NEW ASSEMBLIES ($($ch.Added.Count))"; foreach ($l in $addLines) { & $say "  $l" }; & $say "" }
    if ($ch.Moved.Count) { & $say "MOVED TO ANOTHER TL ($($ch.Moved.Count))"; foreach ($l in $movLines) { & $say "  $l" }; & $say "" }
    if ($ch.Changed.Count) { & $say "CHANGED ($($ch.Changed.Count))"; foreach ($a in $ch.Changed) { & $say "  $($a.Asm)   $($a.Tl) / $($a.Seq)"; foreach ($d in $a.Details) { & $say "      $d" } }; & $say "" }
    if ($remDrop.Count) { & $say "DROPPED FROM THE REQUIRED ASSM ($($remDrop.Count))"; foreach ($l in $remDropLines) { & $say "  $l" }; & $say "" }
    if ($remNoData.Count) { & $say "STILL REQUIRED BUT NO .KSS DATA ($($remNoData.Count)): add the .kss file that holds them"; foreach ($l in $remNoDataLines) { & $say "  $l" }; & $say "" }

    $missing = @($MissingAssemblies | Where-Object { $_ })
    $newMissing = @($missing | Where-Object { $m = $_; -not (@($Prev.Bom | Where-Object { $_.Asm -eq $m })) })
    if ($newMissing.Count -gt 0) {
        Write-Host ""
        Write-Host "  Listed in the Required Assm but found in none of the selected .kss files: $($newMissing.Count)" -ForegroundColor Red
        Show-UpdateLines $newMissing 12 'Red'
        Write-Host "    They are not in the workbook. A new TL's .kss file is needed for them (select it, or run the update again once it arrives)." -ForegroundColor Yellow
        & $say "LISTED IN THE REQUIRED ASSM BUT IN NONE OF THE SELECTED .KSS FILES ($($newMissing.Count)): $($newMissing -join ', ')"
        & $say ""
        foreach ($m in $newMissing) { $log.Add([PSCustomObject]@{ Change = 'NO KSS DATA'; Folder = ''; Assembly = $m; Detail = 'new in the Required Assm; no .kss file selected holds it' }) }
    }

    # --- folders to reprint ---
    Write-Host ""
    Write-Host "  Folders (TL / sequence) that need new or reprinted paperwork" -ForegroundColor Yellow
    & $say "FOLDERS THAT NEED NEW OR REPRINTED PAPERWORK (folder label + work order)"
    if ($ch.Folders.Count -eq 0) { Write-Host "    None: every folder is the same as before." -ForegroundColor Green; & $say "  None: every folder is the same as before." }
    foreach ($f in $ch.Folders.Keys) {
        $st = $ch.Folders[$f]
        $color = switch ($st) { 'NEW' { 'Green' } 'CHANGED' { 'Yellow' } default { 'Red' } }
        $txt = switch ($st) { 'NEW' { 'new folder: make it and print its label and work order' } 'CHANGED' { 'reprint the work order, check the folder contents' } default { 'no assemblies left: pull the folder' } }
        Write-Host ("    {0,-8} {1,-18} {2}" -f $st, $f, $txt) -ForegroundColor $color
        & $say ("  {0,-8} {1,-18} {2}" -f $st, $f, $txt)
        $log.Add([PSCustomObject]@{ Change = "FOLDER $st"; Folder = $f; Assembly = ''; Detail = $txt })
    }
    & $say ""

    # --- hand edits: merge ---
    $asmDigits = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($r in @($AsmRows)) { [void]$asmDigits.Add(("$($r.Assembly)" -replace '\D', '')) }
    $items = Get-EditItems -Prev $Prev -PrevState $PrevState -SettingsList $SettingsList -SequenceRows $SequenceRows -ExceptionRows $ExceptionRows -AsmDigits $asmDigits
    $take = @($items | Where-Object { $_.State -eq 'TakeNew' })
    $keep = @($items | Where-Object { $_.State -eq 'KeepOld' })
    $conf = @($items | Where-Object { $_.State -eq 'Conflict' })

    Write-Host ""
    Write-Host "  Values you can edit in the workbook (due dates, paint, explanations, settings, exceptions)" -ForegroundColor Yellow
    Write-Host ("    Updated from the new data (you had not edited them): {0}" -f $take.Count) -ForegroundColor Green
    Write-Host ("    Your edits kept (the new data did not change them):   {0}" -f $keep.Count) -ForegroundColor Green
    Write-Host ("    Need your decision:                                   {0}" -f $conf.Count) -ForegroundColor $(if ($conf.Count) { 'Yellow' } else { 'Green' })
    & $say "HAND-EDITED VALUES: $($take.Count) updated from the new data, $($keep.Count) of your edits kept, $($conf.Count) needed a decision"
    foreach ($i in $take) {
        & $say "  UPDATED  $($i.Group): $($i.Label): '$(Format-EditValue $i $i.Old)' -> '$(Format-EditValue $i $i.New)'"
        $log.Add([PSCustomObject]@{ Change = 'VALUE UPDATED'; Folder = ''; Assembly = ''; Detail = "$($i.Group): $($i.Label): '$(Format-EditValue $i $i.Old)' -> '$(Format-EditValue $i $i.New)'" })
    }
    foreach ($i in $keep) {
        & $say "  KEPT     $($i.Group): $($i.Label): your '$(Format-EditValue $i $i.Old)' (new data says '$(Format-EditValue $i $i.New)')"
        $log.Add([PSCustomObject]@{ Change = 'YOUR EDIT KEPT'; Folder = ''; Assembly = ''; Detail = "$($i.Group): $($i.Label): '$(Format-EditValue $i $i.Old)'" })
    }

    if ($conf.Count -gt 0) {
        Write-Host ""
        Write-Host "  These differ between your old workbook and the new data:" -ForegroundColor Yellow
        $n = 0
        foreach ($i in $conf) {
            $n++
            Write-Host ("    [{0}] {1}: {2}" -f $n, $i.Group, $i.Label) -ForegroundColor White
            Write-Host ("        yours (old workbook): '{0}'" -f (Format-EditValue $i $i.Old)) -ForegroundColor Gray
            Write-Host ("        new (Required Assm / today): '{0}'" -f (Format-EditValue $i $i.New)) -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "  [K] Keep yours (default)    [U] Use the new values    [R] Review one by one" -ForegroundColor Gray
        $ans = (Read-Host "Which for these $($conf.Count)? [K]eep / [U]se new / [R]eview").Trim().ToUpper()
        if ($ans -match '^U') { foreach ($i in $conf) { $i.Use = 'New' } }
        elseif ($ans -match '^R') {
            $n = 0
            foreach ($i in $conf) {
                $n++
                Write-Host ""
                Write-Host ("  [{0}/{1}] {2}: {3}" -f $n, $conf.Count, $i.Group, $i.Label) -ForegroundColor White
                Write-Host ("      [K] yours: '{0}'" -f (Format-EditValue $i $i.Old)) -ForegroundColor Gray
                Write-Host ("      [U] new:   '{0}'" -f (Format-EditValue $i $i.New)) -ForegroundColor Gray
                $one = (Read-Host "Keep yours or use new? [K]eep (default) / [U]se new").Trim().ToUpper()
                if ($one -match '^U') { $i.Use = 'New' }
            }
        }
        & $say "DECISIONS"
        foreach ($i in $conf) {
            $kept = if ($i.Use -eq 'Old') { "kept yours '$(Format-EditValue $i $i.Old)'" } else { "used new '$(Format-EditValue $i $i.New)'" }
            & $say "  $($i.Group): $($i.Label): $kept"
            $log.Add([PSCustomObject]@{ Change = 'YOU DECIDED'; Folder = ''; Assembly = ''; Detail = "$($i.Group): $($i.Label): $kept" })
        }
    }
    & $say ""

    # --- apply the merged values to the rows the run is about to write ---
    $use = { param($i) if ($i.Use -eq 'Old') { $i.Old } else { $i.New } }
    $settingKeep = @{}
    $schedDue = @{}
    foreach ($i in $items) {
        if ($i.State -eq 'Same') { continue }
        $val = & $use $i
        switch ($i.Group) {
            'Settings' { if ($i.Use -eq 'Old') { $settingKeep[$i.Key] = $i.Old } }
            'Statistics schedule' { if ($i.Use -eq 'Old') { $schedDue[$i.Key] = "$($i.Old)" } }
            'Sequence Notes' {
                $row = @($SequenceRows | Where-Object { "$($_.TL)|$($_.Sequence)" -eq $i.Key })[0]
                if ($row) {
                    switch ($i.Field) { 'Due' { $row.DueDate = "$val" } 'Paint' { $row.Paint = "$val" } 'Expl' { $row.Explanation = "$val" } }
                }
            }
        }
    }
    foreach ($r in @($SequenceRows)) {
        $sk = "$($r.TL) / $($r.Sequence)"
        if ($Prev.Schedule.ContainsKey($sk) -and -not $schedDue.ContainsKey($sk) -and -not (@($items | Where-Object { $_.Group -eq 'Statistics schedule' -and $_.Key -eq $sk }))) { $schedDue[$sk] = "$($r.DueDate)" }
    }
    # Exceptions: apply per ASSM, creating a row for an ASSM that only the old workbook had an override for
    $finalExc = [System.Collections.Generic.List[object]]::new()
    foreach ($e in @($ExceptionRows)) { $finalExc.Add($e) }
    foreach ($g in ($items | Where-Object { $_.Group -eq 'Assembly Exceptions' } | Group-Object Key)) {
        $row = @($finalExc | Where-Object { "$($_.Assembly)" -eq $g.Name })[0]
        $vals = @{ Due = ''; Paint = ''; Expl = '' }
        foreach ($i in $g.Group) { $vals[$i.Field] = "$(& $use $i)" }
        if (-not $row) {
            if (-not ($vals.Due -or $vals.Paint -or $vals.Expl)) { continue }
            $a = @($AsmRows | Where-Object { ("$($_.Assembly)" -replace '\D', '') -eq $g.Name } | Sort-Object { [int]$_.TlNumber })[0]
            if (-not $a) { continue }
            $tlName = if ($TlNames.ContainsKey([int]$a.TlNumber)) { $TlNames[[int]$a.TlNumber] } else { "TL$($a.TlNumber)" }
            $row = [PSCustomObject]@{ Assembly = $g.Name; TL = $tlName; Sequence = $a.Sequence; Paint = ''; Explanation = ''; DueDate = '' }
            $finalExc.Add($row)
        }
        $row.Paint = $vals.Paint; $row.Explanation = $vals.Expl; $row.DueDate = $vals.Due
    }

    # --- Master BOM status column: carried over for parts that did not change; a changed part says so ---
    $statusMap = @{}
    $carried = 0; $flagged = 0
    foreach ($r in $Prev.Bom) {
        $st = "$($r.Status)".Trim()
        if (-not $st -or $st -ieq 'PENDING') { continue }
        $k = "$($r.Asm)|$($r.Part)"
        if ($ch.ChangedParts.Contains($k)) {
            if ($st -notmatch '^REVISED') { $statusMap[$k] = "REVISED (was $st)"; $flagged++ } else { $statusMap[$k] = $st }
        }
        else { $statusMap[$k] = $st; $carried++ }
    }
    if ($carried + $flagged -gt 0) {
        Write-Host ""
        Write-Host "  Status column: $carried carried over unchanged; $flagged on parts that changed are marked REVISED (was ...)." -ForegroundColor Gray
        & $say "STATUS COLUMN: $carried carried over; $flagged marked REVISED because the part changed"
        & $say ""
    }

    # --- stuff outside what the update can carry over ---
    Write-Host ""
    Write-Host "  The previous workbook is kept as it is; this builds a new one. Anything you typed outside the cells listed above" -ForegroundColor DarkGray
    Write-Host "  (for example on the Master BOM, plates or Shop Assemblies tabs) is not carried over." -ForegroundColor DarkGray

    $reportPath = Join-Path $ReportFolder ("Update Report - {0} {1}.txt" -f (($JobLabel -replace '[\\/:*?"<>|]', '_').Trim()), $when.ToString('yyyyMMdd-HHmmss'))
    try { [System.IO.File]::WriteAllLines($reportPath, $rep, (New-Object System.Text.UTF8Encoding($false))) } catch { $reportPath = $null }
    if ($reportPath) { Write-Host "  Full report: $([System.IO.Path]::GetFileName($reportPath))" -ForegroundColor Gray }

    Write-Host ""
    $go = (Read-Host "Build the updated workbook now? [Y]es (default) / [N]o, stop here").Trim()
    return [PSCustomObject]@{
        Proceed = ($go -notmatch '^(n|no)$')
        Status = $statusMap
        SettingKeep = $settingKeep
        SchedDue = $schedDue
        ExceptionRows = @($finalExc)
        Log = @($log)
        ReportPath = $reportPath
        FolderCount = $ch.Folders.Count
    }
}

# "Update Log" tab: the same list as the report, inside the new workbook.
function Add-UpdateLogSheet {
    param($Workbook, $Rows, [string]$PreviousName)
    $sheet = $Workbook.Sheets.Add([Type]::Missing, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $sheet.Name = "Update Log"
    $sheet.Cells.Font.Name = "Arial"
    $n = @($Rows).Count
    $data = New-Object "object[,]" ($n + 1), 5
    $headers = @("Change", "Folder", "Assembly", "Detail")
    for ($c = 0; $c -lt 4; $c++) { $data[0, $c] = $headers[$c] }
    $data[0, 4] = "What changed compared with '$PreviousName' when this workbook was built. Folders marked NEW / CHANGED need new or reprinted labels and work orders."
    for ($i = 0; $i -lt $n; $i++) {
        $r = $Rows[$i]
        $data[($i + 1), 0] = "$($r.Change)"; $data[($i + 1), 1] = "$($r.Folder)"; $data[($i + 1), 2] = "$($r.Assembly)"; $data[($i + 1), 3] = "$($r.Detail)"
    }
    $sheet.Range("A2:D$($n + 2)").NumberFormat = "@"
    Write-SheetArray -Sheet $sheet -Data $data -RowCount ($n + 1) -ColCount 5
    $hdr = $sheet.Range("A1:D1")
    $hdr.Font.Bold = $true
    $hdr.Font.Color = Get-OleColor 255 255 255
    $hdr.Interior.Color = Get-OleColor 27 54 93
    $sheet.Range("E1").Font.Italic = $true
    $widths = [double[]]@(18, 20, 14, 90)
    for ($c = 0; $c -lt $widths.Count; $c++) { $sheet.Columns.Item($c + 1).ColumnWidth = $widths[$c] }
    if ($n -gt 0) { [void]$sheet.Range("A1:D$($n + 1)").AutoFilter() }
    return $sheet
}

# Notes shown next to each .kss file on the selection screen of an update: used by the previous workbook, NEW, or CHANGED since.
function Get-KssUpdateNotes {
    param($Prev, $PrevState, [string]$Folder)
    $notes = @{}
    $used = @($Prev.SourceFiles)
    foreach ($f in @(Get-ChildItem -LiteralPath $Folder -Filter *.kss -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch ' - Filtered Assemblies' -and $_.Name -notmatch ' \(Generated\)' })) {
        $wasUsed = ($used | Where-Object { $_ -ieq $f.Name }).Count -gt 0
        if ($PrevState -and $PrevState.Files.Count -gt 0) {
            if ($PrevState.Files.ContainsKey($f.Name)) {
                $h = ""
                try { $h = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash } catch { }
                $notes[$f.Name] = if ($h -and $PrevState.Files[$f.Name] -and $h -ne $PrevState.Files[$f.Name]) { 'CHANGED' } else { 'same as before' }
            }
            else { $notes[$f.Name] = 'NEW' }
        }
        elseif ($wasUsed) { $notes[$f.Name] = if ($f.LastWriteTime -gt $Prev.Built) { 'CHANGED' } else { 'same as before' } }
        elseif ($f.LastWriteTime -gt $Prev.Built) { $notes[$f.Name] = 'NEW' }
    }
    return $notes
}


# 1. CHECK THE COMPUTER, THEN THE MAIN MENU
try { Clear-Host } catch { }
Write-Host ""
Write-Host ("=" * 56)
Write-Host "  KISS  -  KSS Nesting, Hardware & Shop Management"
Write-Host ("=" * 56)
Test-KissEnvironment
Start-Sleep -Milliseconds 900

$menuStatus = "Computer check passed"
while ($true) {
    $menuChoice = Show-MainMenu $menuStatus
    if ($menuChoice -eq 'Q') { return }
    if ($menuChoice -eq '3') {
        try { Clear-Host } catch { }
        $made = New-RequiredAssmTemplate
        Write-Host ""
        Microsoft.PowerShell.Utility\Read-Host "  Press Enter to go back to the menu" | Out-Null
        $menuStatus = if ($made) { "Created: $([System.IO.Path]::GetFileName($made))" } else { "" }
        continue
    }
    break
}
$updateMode = ($menuChoice -eq 'U')
$script:UpdateMode = $false
$script:UpdateStatus = $null; $script:UpdateSettingKeep = $null; $script:UpdateSchedDue = $null; $script:UpdateLogRows = @()
$script:FreshSeeds = $null
$prevBom = $null; $prevState = $null; $prevFolder = $null
try { Clear-Host } catch { }
Write-Host ""
Write-Host "  Working folder: $((Get-Location).Path)" -ForegroundColor DarkGray
Write-Host "  Answers in [brackets]; press Enter to take the (default)." -ForegroundColor DarkGray
$modeChoice = if ($menuChoice -eq '2') { "2" } else { "1" }

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
    $kssNotes = @{}
    $preTick = @()
    if ($updateMode) {
        $prevPath = Select-RequiredAssmFile -NamePattern '(?i)\sBOM(\s\d{8}-\d{6})?\.xlsx$' -ExcludePattern '(?i)before update' `
            -Heading 'Previous BOM workbook' -Noun 'BOM workbooks' -SkipText 'Cancel (run a normal build instead)'
        if ($prevPath) {
            try {
                Write-Log "Reading the previous workbook (Excel runs hidden)..."
                $prevBom = Read-PreviousBom -Path $prevPath -SettingNames @(Get-SettingsList | ForEach-Object { $_[0] })
                $prevState = Read-KissState $prevPath
                $prevFolder = Split-Path -Parent $prevPath
                $prevAsm = @($prevBom.Bom | ForEach-Object { $_.Asm } | Select-Object -Unique).Count
                $prevTls = @($prevBom.Bom | ForEach-Object { $_.Tl } | Where-Object { $_ } | Select-Object -Unique)
                Write-Log "Previous workbook: $prevAsm assemblies, $($prevBom.Bom.Count) part rows, transmittals $($prevTls -join ', ')" "Green"
                if (-not $prevState) { Write-Log "No 'last build' record next to it (it was made before updates existed). Values that differ will be shown for you to choose." "Yellow" }
                $kssNotes = Get-KssUpdateNotes -Prev $prevBom -PrevState $prevState -Folder (Get-Location).Path
                $preTick = @(@($kssNotes.Keys | Where-Object { $kssNotes[$_] }) + @($prevBom.SourceFiles))
                $newFiles = @($kssNotes.Keys | Where-Object { $kssNotes[$_] -eq 'NEW' } | Sort-Object)
                $changedFiles = @($kssNotes.Keys | Where-Object { $kssNotes[$_] -eq 'CHANGED' } | Sort-Object)
                if ($newFiles.Count -gt 0) { Write-Log "New .kss file(s) since then: $($newFiles -join ', ')" "Yellow" }
                if ($changedFiles.Count -gt 0) { Write-Log "Changed .kss file(s) since then: $($changedFiles -join ', ')" "Yellow" }
                Write-Log "The files the previous workbook used and any new / changed ones are ticked for you. A new TL that came without a .kss needs no file: its assemblies are found in the ticked ones." "Gray"
                Start-Sleep -Milliseconds 600
            }
            catch {
                Write-Log "Could not read the previous workbook: $($_.Exception.Message)" "Red"
                Write-Log "Running a normal build instead." "Yellow"
                $updateMode = $false; $prevBom = $null; $prevState = $null
            }
        }
        else { Write-Log "No previous workbook picked; running a normal build instead." "Yellow"; $updateMode = $false }
    }
    $script:UpdateMode = $updateMode
    $kssFiles = @(Select-KssFiles -Notes $kssNotes -PreTick $preTick)
    if ($kssFiles.Count -eq 0) {
        Write-Log "No KSS files selected. Put the job's .kss files in this folder ($((Get-Location).Path)) and run again." "Red"
        return
    }
    $allTLs = @($kssFiles | ForEach-Object { $_.TL } | Where-Object { $_ } | Select-Object -Unique)
    Write-Log "Processing $($kssFiles.Count) KSS file(s)..."
    $headerLocked = $false
    $script:ShapeRepairs = 0; $script:BlankShapes = 0

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

                    # Some exports write a bolt with no shape and the whole description in the size field ("HSB 3/4x1-3/4 A325", mark
                    # "HSB_3/4X1-3/4_A3", length 0). KISS / ProNest reject that ("Invalid Shape ()"), so it is written like the other bolts:
                    # shape HSB, size 3/4X1-3/4, grade A325, length from the bolt size, no part mark.
                    if (-not $shape) {
                        $fix = $null
                        foreach ($probe in @($size, $desc, $rawPartMark)) {
                            if ("$probe" -match '^\s*([A-Za-z]{2,5})[ _]+(\d[\dA-Za-z/.\-]*)(?:[ _]+([A-Za-z]{1,3}\d{2,4}\w*))?\s*$') { $fix = $Matches; break }
                        }
                        if ($fix) {
                            $shape = $fix[1].ToUpper()
                            $size = $fix[2].ToUpper()
                            if (-not $grade -and $fix[3]) { $grade = $fix[3].ToUpper() }
                            $lenNum = 0.0
                            if (-not [double]::TryParse($length, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$lenNum) -or $lenNum -le 0) {
                                $lastDim = ($size -split 'X')[-1]
                                $inches = 0.0
                                foreach ($piece in ($lastDim -split '-')) {
                                    if ($piece -match '^(\d+)/(\d+)$') { $inches += [double]$Matches[1] / [double]$Matches[2] }
                                    elseif ($piece -match '^\d+(\.\d+)?$') { $inches += [double]::Parse($piece, [System.Globalization.CultureInfo]::InvariantCulture) }
                                }
                                if ($inches -gt 0) { $length = ($inches * 25.4).ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture) }
                            }
                            $rawPartMark = ""
                            $script:ShapeRepairs++
                            if ($script:ShapeRepairs -le 3) { Write-Log "Blank shape repaired in $($file.Name): assembly $assm now $shape $size $grade (length $length)." "Yellow" }
                        }
                        else {
                            $script:BlankShapes++
                            if ($script:BlankShapes -le 5) { Write-Log "WARNING: $($file.Name) assembly $assm part '$rawPartMark' has no shape; KISS import will reject it." "Red" }
                        }
                    }

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

    if ($script:ShapeRepairs -gt 3) { Write-Log "Blank shape repaired on $($script:ShapeRepairs) record(s) in total." "Yellow" }
    if ($script:BlankShapes -gt 0) { Write-Log "$($script:BlankShapes) record(s) have no shape and could not be repaired; they will fail to import." "Red" }
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
$workbookTLs = @()
$ourJobNumber = ""
$issuedBy = ""
if ($modeChoice -ne "2") {
    $requiredAssmPath = Select-RequiredAssmFile -ExtraFolders @(@($kssFiles | ForEach-Object { Split-Path -Parent $_.FullName }) + @($prevFolder))
    $guessJob = ""
    if ($requiredAssmPath -and [System.IO.Path]::GetFileName($requiredAssmPath) -match '^\s*(\d{4,6})') { $guessJob = $Matches[1] }
    $prevInitials = ""
    if ($updateMode) {
        $pj = if ($prevBom.Settings.ContainsKey('Set_OurJob')) { Get-CellText $prevBom.Settings['Set_OurJob'] } else { "" }
        if ($pj) { $guessJob = $pj }
        if ($prevBom.Settings.ContainsKey('Set_Initials')) { $prevInitials = Get-CellText $prevBom.Settings['Set_Initials'] }
    }
    $jobPrompt = if ($guessJob) { "Our job number (default: $guessJob)" } else { "Our job number (shown on the labels and work orders)" }
    $ourJobInput = (Read-Host $jobPrompt).Trim()
    $ourJobNumber = if ($ourJobInput) { $ourJobInput } else { $guessJob }
    if ($requiredAssmPath) {
        $initialsPrompt = if ($prevInitials) { "Initials of who is making the folders (Enter = $prevInitials)" } else { "Initials of who is making the folders (Enter to leave blank)" }
        $issuedBy = (Read-Host $initialsPrompt).Trim().ToUpper()
        if (-not $issuedBy -and $prevInitials) { $issuedBy = $prevInitials.ToUpper() }
        # The workbook is read now because it decides which assemblies are in this run (see "Assembly filter" below).
        try {
            Write-Log "Reading the Required Assm workbook..."
            $workbookTLs = @(Read-RequiredAssmStandalone -Path $requiredAssmPath)
        }
        catch { Write-Log "Could not read the Required Assm workbook: $($_.Exception.Message)" "Red" }
    }
}

$whitelistResult = $null
$usedWorkbookFilter = $false
if ($workbookTLs.Count -gt 0 -and $modeChoice -ne "2") {
    # A Required Assm workbook is the filter: its assemblies, per transmittal (REV columns included), are the run.
    Write-Section "Assembly filter (from the Required Assm workbook)"
    $extraTLs = @()
    # A transmittal that is in the workbook but has no .kss file (a rev that came without one) is part of the run: it is placed by its
    # assembly numbers from the selected files. A transmittal whose .kss exists in the folder but was not ticked stays out (an update
    # takes every transmittal of the workbook).
    $haveNums = @($allTLs | ForEach-Object { [int]($_ -replace '\D', '0') })
    $folderNums = @()
    if (-not $updateMode) {
        $kssDirs = @($kssFiles | ForEach-Object { Split-Path -Parent $_.FullName } | Select-Object -Unique)
        if ($kssDirs.Count -eq 0) { $kssDirs = @((Get-Location).Path) }
        $folderNums = @($kssDirs | ForEach-Object { Get-ChildItem -LiteralPath $_ -Filter *.kss -File -ErrorAction SilentlyContinue } |
            ForEach-Object { Get-TransmittalFromName $_.Name } | Where-Object { $_ } | ForEach-Object { [int]($_ -replace '\D', '0') })
    }
    $extraTLs = @($workbookTLs | Where-Object { $haveNums -notcontains [int]$_.Number -and $folderNums -notcontains [int]$_.Number } | Sort-Object Number)
    if ($true) {
        if ($extraTLs.Count -gt 0) {
            $allTLs = @($allTLs) + @($extraTLs | ForEach-Object { $_.TL })
            Write-Log "In the Required Assm workbook with no .kss file of their own: $(($extraTLs | ForEach-Object { $_.TL }) -join ', '). Their assemblies are taken from the selected .kss files." "Cyan"
        }
    }
    $wf = Get-WorkbookAssemblyFilter -WorkbookTLs $workbookTLs -RunTLs $allTLs
    if ($wf.AssemblyCount -eq 0) {
        Write-Log "None of the selected transmittals ($($allTLs -join ', ')) has assemblies in the workbook, so it cannot filter this run." "Yellow"
        Write-Log "Check the TL names in Sheet1 of the workbook. Falling back to the manual assembly list." "Yellow"
    }
    else {
        foreach ($u in $wf.Used) {
            $kind = if ($u.IsRev) { "  (REV: replaces the same assemblies from earlier TLs)" } else { "" }
            Write-Host ("  {0,-8} {1,4} assemblies{2}" -f $u.TL, $u.Count, $kind)
        }
        if ($wf.NotInWorkbook.Count -gt 0) {
            Write-Log "Selected but NOT in the workbook, so none of their assemblies are included: $($wf.NotInWorkbook -join ', ')" "Yellow"
            foreach ($missingTl in $wf.NotInWorkbook) {
                # A workbook transmittal that was not selected but holds most of this file's assemblies is probably the same TL with a typo in its name.
                $fileAsm = @($masterData | Where-Object { $_.TL -eq $missingTl } | ForEach-Object { "$($_.Assembly)".Trim() } | Select-Object -Unique)
                foreach ($other in @($workbookTLs | Where-Object { $wf.NotInRun -contains $_.TL })) {
                    $hit = 0
                    foreach ($a in $fileAsm) { foreach ($set in $other.SeqAssemblies.Values) { if ($set.Contains($a)) { $hit++; break } } }
                    if ($fileAsm.Count -gt 0 -and $hit -ge ($fileAsm.Count / 2)) {
                        Write-Log "  $hit of $($fileAsm.Count) assemblies in $missingTl are listed under $($other.TL) in the workbook. Fix the TL name in the workbook (Sheet1 header) and run again." "Yellow"
                    }
                }
            }
        }
        if ($wf.NotInRun.Count -gt 0) {
            Write-Log "In the workbook but not selected for this run (ignored): $($wf.NotInRun -join ', ')" "DarkGray"
        }
        $inFiles = @($masterData | ForEach-Object { "$($_.Assembly)".Trim() } | Where-Object { $_ } | Select-Object -Unique)
        $wanted = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($u in $workbookTLs) { foreach ($set in $u.SeqAssemblies.Values) { foreach ($a in $set) { [void]$wanted.Add($a) } } }
        $leftOut = @($inFiles | Where-Object { -not $wanted.Contains($_) })
        if ($leftOut.Count -gt 0) {
            Write-Log "$($leftOut.Count) assembly(ies) in the selected files are not in the workbook and are left out." "Yellow"
        }

        $whitelistResult = Apply-AssemblyWhitelist -MasterData $masterData -LaborData $laborData -AssemblyDrawings $assemblyDrawings -AssemblyListText $wf.Text
        if ($whitelistResult.Applied) {
            if ($whitelistResult.MasterData.Count -eq 0) {
                throw "None of the workbook's assemblies were found in the selected .kss files. Check that the right files are selected and the TL names in the workbook match."
            }
            $masterData = $whitelistResult.MasterData
            $laborData = $whitelistResult.LaborData
            $assemblyDrawings = $whitelistResult.AssemblyDrawings
            $usedWorkbookFilter = $true
            foreach ($x in $extraTLs) {
                $xAsm = @($x.SeqAssemblies.Values | ForEach-Object { $_ } | Sort-Object -Unique)
                $xMissing = @($xAsm | Where-Object { $whitelistResult.Missing -contains $_ })
                Write-Log "$($x.TL) (no .kss): $($xAsm.Count - $xMissing.Count) of $($xAsm.Count) assemblies found in the selected files." "Cyan"
                if ($xMissing.Count -gt 0) {
                    Write-Log "$($x.TL): NO .kss DATA for $($xMissing.Count) assembly(ies), so they are NOT in the output: $(($xMissing | Select-Object -First 15) -join ', ')$(if ($xMissing.Count -gt 15) { ' ...' }). Select the .kss that holds them." "Red"
                }
            }

            # Where every transmittal's assemblies ended up: the cost code (TL) written to the filtered .kss is the LAST transmittal that lists the assembly,
            # so a TL whose assemblies are all listed again by a later TL / REV disappears from the file. Say so per transmittal.
            Write-Host ""
            Write-Host "  Transmittals in the filtered KISS file (cost code):" -ForegroundColor Yellow
            $finalTl = @{}
            foreach ($mi in $masterData) { $finalTl["$($mi.Assembly)".Trim()] = "$($mi.TL)" }
            foreach ($wt in @($workbookTLs | Where-Object { $n = [int]$_.Number; @($allTLs | Where-Object { [int]($_ -replace '\D', '0') -eq $n }).Count -gt 0 } | Sort-Object Number)) {
                $listed = @($wt.SeqAssemblies.Values | ForEach-Object { $_ } | Sort-Object -Unique)
                $mine = "TL$($wt.Number)"
                $inFile = @($listed | Where-Object { $finalTl.ContainsKey($_) -and $finalTl[$_] -eq $mine })
                $noData = @($listed | Where-Object { -not $finalTl.ContainsKey($_) })
                $moved = @($listed | Where-Object { $finalTl.ContainsKey($_) -and $finalTl[$_] -ne $mine })
                $movedTo = @($moved | ForEach-Object { $finalTl[$_] } | Select-Object -Unique) -join ', '
                $line = "    {0,-9} listed {1,4}   in the file {2,4}" -f $wt.TL, $listed.Count, $inFile.Count
                if ($moved.Count -gt 0) { $line += "   listed again by $movedTo (shown there): $($moved.Count)" }
                if ($noData.Count -gt 0) { $line += "   NO .kss DATA: $($noData.Count)" }
                $color = if ($noData.Count -gt 0 -or $inFile.Count -eq 0) { 'Red' } else { 'Gray' }
                Write-Host $line -ForegroundColor $color
            }
            Write-Host ""
        }
    }
}

if (-not $usedWorkbookFilter) {
    Write-Section "Assembly filter"
    Write-Host "  [N] No (default): include ALL assemblies from the selected files"
    Write-Host "  [Y] Yes: only the assemblies you list (REV runs, partial builds)"
    Write-Host ""
    $whitelistChoice = Read-Host "Limit the run to a specific assembly list? [Y]es / [N]o (default: No)"
    if ($whitelistChoice -match '^(y|yes)$') {
        Write-Host ""
        Write-Host "  How to write the list (commas, spaces or new lines all work):"
        Write-Host "    Plain list ........ 51328511 51318512"
        Write-Host "    With a TL ......... TL1296  51328511 51318512"
        Write-Host "    REV (R prefix) .... R51328511   replaces that part from earlier TLs"
        Write-Host "    TL order .......... 1_TL1301  2_TL1282  3_TL1296   (number sets the order)"
        Write-Host ""
        $whitelistText = Read-Host "Paste the list, or type @path.txt to load a file (Enter = use what is on the clipboard)"
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
}

Write-Log "Filtered Detail Records: $($masterData.Count)" "Cyan"

# Check if REV assemblies exist and offer override option
$revAssemblies = @{}
$applyRevOverride = $true
if ($null -ne $whitelistResult -and $whitelistResult.Applied) {
    $revAssemblies = $whitelistResult.RevMap
}

if ($revAssemblies.Count -gt 0 -and $usedWorkbookFilter) {
    # REV columns (TLR####) in the workbook always replace the same assembly from earlier TLs; nothing to ask.
    Write-Log "REV assemblies from the workbook: $($revAssemblies.Count) (each replaces the same assembly from earlier TLs)." "Cyan"
}
elseif ($revAssemblies.Count -gt 0) {
    Write-Section "REV assemblies"
    Write-Host "  Found $($revAssemblies.Count) REV-marked assembly(ies): $($revAssemblies.Keys -join ', ')"
    Write-Host "  [Y] Yes (default): a REV assembly replaces the same part from earlier TLs"
    Write-Host "  [N] No: keep every version (the part shows once per TL)"
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
    Write-Section "Duplicate parts across TLs"
    Write-Host "  Assemblies in more than one file: $($sharedAssemblies.Count)    Part marks in more than one file: $($sharedMarks.Count)"
    Write-Host "  [M] Merge (default): one row per part, quantities added (purchasing, full BOM)"
    Write-Host "  [S] Separate: one row per part per TL (comparing TLs, REV tracking)"
    Write-Host ""
    if ($updateMode -and $prevState -and $null -ne $prevState.Merge) {
        $mergeDuplicates = [bool]$prevState.Merge
        Write-Log "Same choice as the last build: $(if ($mergeDuplicates) { 'Merge' } else { 'Separate' })." "Gray"
    }
    else {
        $dupChoice = Read-Host "Duplicates across files: [M]erge (default) or [S]eparate?"
        if ($dupChoice -match '^\s*s') { $mergeDuplicates = $false }
    }

    if ($mergeDuplicates) {
        $conflicts = @($sharedMarks | Where-Object { @($_.Group | ForEach-Object { "$($_.Shape)|$($_.Size)|$($_.Grade)|$($_.Length)".ToLowerInvariant() } | Select-Object -Unique).Count -gt 1 })
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

    # A REV assembly keeps only the rows from the newest file that holds it; the older files' rows of that same assembly are dropped.
    # Every other assembly (and every row of a REV assembly in its newest file) is kept as is.
    $newestFile = @{}
    foreach ($g in ($masterData | Where-Object { $revAssemblies.ContainsKey($_.Assembly) } | Group-Object Assembly)) {
        $newestFile[$g.Name] = @($g.Group | ForEach-Object { [string]$_.SourceFile } | Select-Object -Unique |
            Sort-Object { $tlOrder[$_] } -Descending)[0]
    }
    $filteredData = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $masterData) {
        if ($newestFile.ContainsKey($item.Assembly) -and [string]$item.SourceFile -ne $newestFile[$item.Assembly]) { continue }
        $filteredData.Add($item)
    }

    if ($filteredData.Count -lt $masterData.Count) {
        $removed = $masterData.Count - $filteredData.Count
        Write-Log "REV override: removed $removed record(s) of $($newestFile.Count) REV assembly(ies) from older TLs, keeping the newest version." "Green"
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

# Repackage the filtered combined pool into a clean KISS file with only selected assemblies.
$kssFilesInScope = $ExecutionContext.SessionState.PSVariable.GetValue('kssFiles', $null)
if ($null -ne $kssFilesInScope -and $kssFilesInScope.Count -gt 0) {
    $filteredKssOutput = Write-FilteredAssemblyKssFile -MasterData $masterData -AssemblyDrawings $assemblyDrawings -ProjectName $projectName -JobNumber $jobNumber -SourceFiles $kssFilesInScope -TLs $allTLs
    if ($filteredKssOutput) {
        Write-Log "Generated filtered KISS package: $filteredKssOutput" "Green"
        Add-OutputFile $filteredKssOutput
    }
}


# 3. EXCEL COM SETUP
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false 
$workbook = $excel.Workbooks.Add()


# --- Per-assembly facts shared by the Master BOM columns, the Sequence Notes tab, the labels and the work orders ---
$allWorkbookTLs = @()
$script:ExceptionRows = @()
$script:ExceptionNotInRun = @()
$script:AsmExceptions = @{}
$asmRows = @()
$workOrderPages = @()
$sequenceNoteRows = @()
$asmByKey = @{}
if ($modeChoice -ne "2") {
    if ($requiredAssmPath) {
        $allWorkbookTLs = @($workbookTLs)        # already read before the assembly filter
    }
    $revAssemblies = @{}; $tlNames = @{}; $workOrderDue = @{}; $revisesByKey = @{}; $workOrderNotes = @{}
    foreach ($tlEntry in @($allWorkbookTLs | Sort-Object Number)) {
        $tlNames[[int]$tlEntry.Number] = $tlEntry.TL
        foreach ($seq in $tlEntry.Sequences) {
            $note = $tlEntry.SeqNotes[$seq]
            $due = if ($note -and $note.DueDate) { $note.DueDate } else { $tlEntry.DueDate }
            if ($due) { $workOrderDue["$([int]$tlEntry.Number)|$seq"] = $due }
            $coat = if ($note -and $note.Coating) { $note.Coating } else { $tlEntry.Coating }
            $expl = if ($note -and $note.Explanation) { $note.Explanation } else { "" }
            if (($coat -or $expl) -and -not $workOrderNotes.ContainsKey($seq)) { $workOrderNotes[$seq] = @{ Coating = $coat; Explanation = $expl } }
            if ($tlEntry.IsRev) {
                $hits = @(Get-RevisedTransmittals -Tl $tlEntry -AllTLs $allWorkbookTLs -Sequences @($seq) | ForEach-Object { $_.TL })
                if ($hits.Count -gt 0) { $revisesByKey["$([int]$tlEntry.Number)|$seq"] = $hits -join ", " }
            }
        }
        if ($tlEntry.IsRev) { foreach ($set in $tlEntry.SeqAssemblies.Values) { foreach ($asm in $set) { $revAssemblies[$asm] = [int]$tlEntry.Number } } }
    }
    $asmRows = @(Get-AssemblyRows -MasterData $masterData -RevAssemblies $revAssemblies)
    foreach ($r in $asmRows) { $asmByKey["$($r.TL)|$($r.Assembly)"] = $r }
    $workOrderPages = @(Get-WorkOrderPages -AssemblyRows $asmRows -SeqNotes $workOrderNotes -TlNames $tlNames -DueDates $workOrderDue -Revises $revisesByKey | Sort-Object TlNumber, Sequence)

    # One Sequence Notes row per placed transmittal + sequence: every work order, plus every Required Assm folder of this run.
    # Due date and paint/notes start from Required Assm (TL Notes tab); blank where there is nothing.
    $runNumbers = @($allTLs | ForEach-Object { [int]($_ -replace '\D', '0') })
    $seedByKey = [ordered]@{}
    foreach ($wo in $workOrderPages) {
        $seedByKey["$($wo.TL)|$($wo.Sequence)"] = [PSCustomObject]@{ TL = $wo.TL; TlNumber = $wo.TlNumber; Sequence = $wo.Sequence; IsRev = $wo.IsRev; Revises = $wo.Revises; DueDate = $wo.DueDate; Paint = $wo.Paint; Explanation = $wo.Explanation }
    }
    foreach ($tlEntry in $allWorkbookTLs) {
        if ($runNumbers.Count -gt 0 -and $runNumbers -notcontains $tlEntry.Number) { continue }
        foreach ($seq in $tlEntry.Sequences) {
            $note = $tlEntry.SeqNotes[$seq]
            $due = if ($note -and $note.DueDate) { $note.DueDate } else { $tlEntry.DueDate }
            $coat = if ($note -and $note.Coating) { $note.Coating } else { $tlEntry.Coating }
            $expl = if ($note -and $note.Explanation) { $note.Explanation } else { "" }
            $key = "$($tlEntry.TL)|$seq"
            $rev = if ($revisesByKey.ContainsKey("$([int]$tlEntry.Number)|$seq")) { $revisesByKey["$([int]$tlEntry.Number)|$seq"] } else { "" }
            $seedByKey[$key] = [PSCustomObject]@{ TL = $tlEntry.TL; TlNumber = $tlEntry.Number; Sequence = $seq; IsRev = $tlEntry.IsRev; Revises = $rev; DueDate = $due; Paint = $coat; Explanation = $expl }
        }
    }
    $sequenceNoteRows = @($seedByKey.Values | Sort-Object TlNumber, Sequence)

    # ASSM-specific exceptions (Required Assm "Exceptions" tab): split into those in this run and those not
    $exceptionMap = @{}
    if ($allWorkbookTLs.Count -gt 0 -and $allWorkbookTLs[0].Exceptions) { $exceptionMap = $allWorkbookTLs[0].Exceptions }
    $exceptionResult = Get-ExceptionRows -ExceptionMap $exceptionMap -AsmRows $asmRows -TlNames $tlNames -WorkbookTLs $allWorkbookTLs
    $script:ExceptionRows = @($exceptionResult.InRun)
    $script:ExceptionNotInRun = @($exceptionResult.NotInRun)
    $script:AsmExceptions = @{}
    foreach ($e in $script:ExceptionRows) { $script:AsmExceptions[$e.Assembly] = $e }
    if ($exceptionMap.Count -gt 0) {
        $byTl = @($script:ExceptionRows | Group-Object TL | Sort-Object Name | ForEach-Object { "$($_.Name) x$($_.Count)" })
        Write-Log "Exceptions tab: $($exceptionMap.Count) ASSM override(s); $($script:ExceptionRows.Count) in this run$(if ($byTl.Count) { ' (' + ($byTl -join ', ') + ')' })." "Green"
        if ($script:ExceptionNotInRun.Count -gt 0) {
            Write-Log "$($script:ExceptionNotInRun.Count) exception ASSM(s) are not in this run and were not applied: $((@($script:ExceptionNotInRun | ForEach-Object { "$($_.Assembly) ($($_.Why))" }) | Select-Object -First 8) -join '; ')$(if ($script:ExceptionNotInRun.Count -gt 8) { ' ...' })" "Yellow"
        }
        if ($allWorkbookTLs[0].ExceptionDuplicates.Count -gt 0) {
            Write-Log "Exceptions tab lists these ASSM more than once (the last row wins): $(($allWorkbookTLs[0].ExceptionDuplicates | Select-Object -Unique) -join ', ')" "Yellow"
        }
    }
}


# --- WHAT THE INPUTS SAY NOW (saved next to the workbook for the next update), and the update review ---
$settingsPm = "Tyler / Matt / Rob"
if ($modeChoice -ne "2") {
    $settingsNow = @(Get-SettingsList -Title $projectName.ToUpper() -CustomerJob $jobNumber -OurJob $ourJobNumber -Initials $issuedBy -ProjectManagers $settingsPm)
    $script:FreshSeeds = Get-FreshSeeds -SettingsList $settingsNow -SequenceRows $sequenceNoteRows -ExceptionRows $script:ExceptionRows
    if ($updateMode) {
        $stillRequired = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($tlEntry in @($allWorkbookTLs)) { foreach ($set in $tlEntry.SeqAssemblies.Values) { foreach ($a in $set) { [void]$stillRequired.Add($a) } } }
        $missingAsm = if ($whitelistResult -and $whitelistResult.Applied) { @($whitelistResult.Missing) } else { @() }
        $newBomRows = @(Get-NewBomRows -MasterData $masterData -AsmByKey $asmByKey -TlNames $tlNames)
        $reviewLabel = (@($projectName, $(if ($ourJobNumber) { $ourJobNumber } else { $jobNumber })) | Where-Object { $_ }) -join " - "
        $review = Invoke-UpdateReview -Prev $prevBom -PrevState $prevState -NewRows $newBomRows -StillRequired $stillRequired -SettingsList $settingsNow `
            -SequenceRows $sequenceNoteRows -ExceptionRows $script:ExceptionRows -AsmRows $asmRows -TlNames $tlNames -KssFiles $kssFiles -KssNotes $kssNotes `
            -MissingAssemblies $missingAsm -ReportFolder (Get-Location).Path -JobLabel $reviewLabel
        if (-not $review.Proceed) {
            Write-Log "Stopped before building. Nothing was changed; the previous workbook is as it was." "Yellow"
            try { $workbook.Close($false); $excel.Quit() } catch { }
            return
        }
        $script:ExceptionRows = @($review.ExceptionRows)
        $script:AsmExceptions = @{}
        foreach ($e in $script:ExceptionRows) { $script:AsmExceptions[$e.Assembly] = $e }
        $script:UpdateStatus = $review.Status
        $script:UpdateSettingKeep = $review.SettingKeep
        $script:UpdateSchedDue = $review.SchedDue
        $script:UpdateLogRows = @($review.Log)
        if ($review.ReportPath) { Add-OutputFile $review.ReportPath }
    }
}


# --- TAB 1: MASTER BOM ---
Write-Log "Building Tab 1: Master BOM..."
$sheet1 = $workbook.Sheets.Item(1)
$sheet1.Name = "Master BOM"

$mHeaders = @("TL #", "Assembly", "Drawing", "MainMark", "PartMark", "Qty", "Category", "Shape", "Size", "Length", "Grade", "Description", "Unit Wt (lbs)", "Total Wt (lbs)", "Status", "SourceFile")
# Columns Q onward are appended so the fixed letters A-P that Statistics and the plate/profile tabs read never move.
# Assembly columns repeat on every part row of the assembly (filter on Assembly; do not SUM them over parts).
# Due Date and Paint / Notes are formulas into the Sequence Notes tab, where they are edited once per transmittal + sequence.
if ($modeChoice -ne "2") {
    $mHeaders += @("Sequence", "Transmittal", "Rev", "Revises", "Due Date", "Paint / Notes", "Asm Qty", "Asm Wt Ea (lbs)", "Asm Wt Total (lbs)", "Asm Length", "Asm Description", "Part Wt Ea (calc)", "Part Wt Total (calc)")
    # AD / AE: the assembly's explanation (exception override, else the sequence's) and an EXCEPTION flag; only when the run has exceptions
    if ($script:ExceptionRows.Count -gt 0) { $mHeaders += @("Explanation", "Exception") }
}
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
    if ($script:UpdateStatus) {
        $statusKey = "$($item.Assembly)|$($item.PartMark)"
        if ($script:UpdateStatus.ContainsKey($statusKey)) { $mData[$i, 14] = $script:UpdateStatus[$statusKey] }
    }
    $mData[$i, 15] = $item.SourceFile
    if ($modeChoice -ne "2") {
        $a = $asmByKey["$($item.TL)|$($item.Assembly)"]
        $r = 2 + $i
        $tlPlaced = if ($a -and $tlNames.ContainsKey($a.TlNumber)) { $tlNames[$a.TlNumber] } elseif ($a) { "TL$($a.TlNumber)" } else { "$($item.TL)" }
        $seq = if ($a) { $a.Sequence } else { "$($item.Assembly)".Substring(0, [Math]::Min(4, "$($item.Assembly)".Length)) }
        $mData[$i, 16] = $seq
        $mData[$i, 17] = $tlPlaced
        $mData[$i, 18] = if ($tlPlaced -like 'TLR*') { "REV" } else { "" }
        $mData[$i, 19] = if ($revisesByKey.ContainsKey("$($a.TlNumber)|$seq")) { $revisesByKey["$($a.TlNumber)|$seq"] } else { "" }
        if ($a) {
            $mData[$i, 22] = $a.Qty
            $mData[$i, 23] = [Math]::Round($a.WeightEach, 2)
            $mData[$i, 24] = [Math]::Round($a.WeightTotal, 2)
            $mData[$i, 25] = $a.Length
            $mData[$i, 26] = $a.Description
        }
        $partWt = Get-KssPartWeight $item.Shape $item.Size $item.Length $item.Qty
        $partQty = [double]$item.Qty
        $mData[$i, 27] = if ($partQty -gt 0) { [Math]::Round($partWt / $partQty, 2) } else { 0 }
        $mData[$i, 28] = [Math]::Round($partWt, 2)
    }
}

if ($modeChoice -ne "2") { $sheet1.Range("Q2:T$($mRows + 1)").NumberFormat = "@" }   # before the write, so sequence 1234 stays text like the lookup keys
$sheet1.Range("A2").Resize($mRows, $mHeaders.Count).Value2 = $mData
$sheet1.Range("M2:N$($mRows + 1)").NumberFormat = "#,##0.00"
if ($modeChoice -ne "2") {
    $sheet1.Range("X2:Y$($mRows + 1)").NumberFormat = "#,##0.00"
    $sheet1.Range("AB2:AC$($mRows + 1)").NumberFormat = "#,##0.00"
}

if ($modeChoice -ne "2") {
    Write-Log "Building Sequence Notes tab (due date and paint/notes per transmittal and sequence)..."
    [void](Add-SequenceNotesSheet -Workbook $workbook -Rows $sequenceNoteRows)
    if ($script:ExceptionRows.Count -gt 0 -or $script:ExceptionNotInRun.Count -gt 0) {
        Write-Log "Building Assembly Exceptions tab ($($script:ExceptionRows.Count) ASSM in this run)..."
        [void](Add-AssemblyExceptionsSheet -Workbook $workbook -Rows $script:ExceptionRows -NotInRun $script:ExceptionNotInRun)
        Add-SequenceNotesExceptionColumns -Workbook $workbook -RowCount @($sequenceNoteRows).Count
    }
    # Due Date / Paint / Notes read the Sequence Notes tab. Written after the tab exists but before the BOM table is created:
    # a formula typed into an existing table column is stored as an external link to a missing workbook.
    $seqKeyExpr = '$R2&"|"&$Q2'
    if ($script:ExceptionRows.Count -gt 0) {
        # an assembly with an exception takes its effective paint / due / explanation (Assembly Exceptions tab); every other assembly follows its folder
        $asmKeyExpr = '$B2&""'
        $sheet1.Range("U2:U$($mRows + 1)").Formula = "=IFERROR(INDEX(Exc_Due,MATCH($asmKeyExpr,Exc_Key,0))&`"`"," + (Get-SeqNoteRef "F" $seqKeyExpr) + ")"
        $sheet1.Range("V2:V$($mRows + 1)").Formula = "=IFERROR(INDEX(Exc_Paint,MATCH($asmKeyExpr,Exc_Key,0))&`"`"," + (Get-SeqNoteRef "G" $seqKeyExpr) + ")"
        $sheet1.Range("AD2:AD$($mRows + 1)").Formula = "=IFERROR(INDEX(Exc_Expl,MATCH($asmKeyExpr,Exc_Key,0))&`"`"," + (Get-SeqNoteRef "J" $seqKeyExpr) + ")"
        $sheet1.Range("AE2:AE$($mRows + 1)").Formula = "=IF(" + (Get-ExcRef "Status" $asmKeyExpr) + "=`"Exception set`",`"EXCEPTION`",`"`")"
    }
    else {
        $sheet1.Range("U2:U$($mRows + 1)").Formula = "=" + (Get-SeqNoteRef "F" $seqKeyExpr)
        $sheet1.Range("V2:V$($mRows + 1)").Formula = "=" + (Get-SeqNoteRef "G" $seqKeyExpr)
    }
}

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
            $markFiles.Count -gt 1 -and (@($_.Group | ForEach-Object { "$($_.Shape)|$($_.Size)|$($_.Grade)|$($_.Length)|$($_.Description)|$($_.Qty)".ToLowerInvariant() } | Select-Object -Unique).Count -gt 1)
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


# --- SETTINGS TAB: the named cells the labels, work orders and Statistics read ---
[void](Add-SettingsSheet -Workbook $workbook -Title $projectName.ToUpper() -CustomerJob $jobNumber -OurJob $ourJobNumber -Initials $issuedBy -ProjectManagers $settingsPm)
if ($script:UpdateSettingKeep) {
    # Update: the settings you changed by hand in the previous workbook stay as you had them
    foreach ($settingName in @($script:UpdateSettingKeep.Keys)) {
        try { $workbook.Names.Item($settingName).RefersToRange.Value2 = $script:UpdateSettingKeep[$settingName] }
        catch { Write-Log "Could not keep your $settingName setting: $($_.Exception.Message)" "Yellow" }
    }
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
$dashboard.Cells.Item(5, 1).Formula = "=Set_Title"
$dashboard.Range("A4:B4").Merge()
$dashboard.Range("A5:B5").Merge()

$dashboard.Cells.Item(4, 3) = "JOB NUMBER"
$dashboard.Cells.Item(5, 3).Formula = "=TEXTJOIN("" / "",TRUE,Set_CustomerJob,Set_OurJob)"
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

# Builds the QPR sheets (one per TL, or per TL + sequence) in $Workbook. Returns the number of sheets made.
# -ReuseFirstSheet: the standalone QPR workbook uses its empty first sheet.
# -Linked: sheets go into the master BOM workbook; the Job: box follows Settings and the due / paint line follows Sequence Notes.
function Add-QualityProcessSheets {
    param(
        $Workbook,
        [System.Collections.Generic.List[PSCustomObject]]$MasterData,
        [string]$QprJobNumber,
        [string]$GroupBy,  # "TL", "Assembly", or "Sequence"
        [bool]$ReuseFirstSheet = $false,
        [bool]$Linked = $false,
        [string]$CustomerJob = "",
        [string]$OurJob = "",
        [bool]$HasSequenceNotes = $false
    )
    $wb = $Workbook

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
        return 0
    }

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

        $ws = if ($sheetIndex -eq 1 -and $ReuseFirstSheet) {
            $wb.Sheets.Item(1)
        } else {
            $wb.Sheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count))
        }
        if ($Linked) {
            # In the master workbook other tabs already exist: keep the QPR tab name unique
            $baseName = $sheetName; $n = 1
            while (@($wb.Sheets | Where-Object { $_.Name -eq $sheetName -and $_.Index -ne $ws.Index }).Count -gt 0) {
                $n++; $suffix = " ($n)"
                $sheetName = $baseName.Substring(0, [Math]::Min($baseName.Length, 31 - $suffix.Length)) + $suffix
            }
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
        if ($Linked -and (($OurJob -and $QprJobNumber -eq $OurJob) -or ($CustomerJob -and $QprJobNumber -eq $CustomerJob))) {
            # Job name left at the default: follow Settings (our job #, else the customer job #)
            $jobNameRef = if ($OurJob -and $QprJobNumber -eq $OurJob) { "Set_OurJob" } else { "Set_CustomerJob" }
            $tlText = if ($GroupBy -eq "Sequence") { $seqTl } else { $(if ($grp.Name) { $grp.Name } else { "(No TL)" }) }
            $jobParts = @($jobNameRef, """$tlText""")
            if ($GroupBy -eq "Sequence") { $jobParts += """$seqNum""" }
            $ws.Cells.Item(1,14).Formula = "=TEXTJOIN("" - "",TRUE," + ($jobParts -join ",") + ")"
        }
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
        if ($Linked -and $HasSequenceNotes -and $GroupBy -eq "Sequence") {
            # Due date and paint / REVISES note for this TL + sequence, read live from the Sequence Notes tab
            $seqKey = """$seqTl|$seqNum"""
            $noteFormula = "=TRIM(" + (Get-SeqNoteRef "I" $seqKey) + "&IF(" + (Get-SeqNoteRef "H" $seqKey) + "<>"""","" - ""&" + (Get-SeqNoteRef "H" $seqKey) + ",""""))"
            $ws.Cells.Item(4,3).Formula = $noteFormula
            $ws.Cells.Item(4,3).Font.Name = "Arial Narrow"
            $ws.Cells.Item(4,3).Font.Size = 11
            $ws.Cells.Item(4,3).Font.Bold = $true
            $ws.Cells.Item(4,3).HorizontalAlignment = $xlLeft
            $ws.Cells.Item(4,3).VerticalAlignment = $xlBottom
            $sep1.ShrinkToFit = $true
        }
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

    return $groups.Count
}

function Write-QualityProcessReport {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$MasterData,
        [string]$ProjectName,
        [string]$JobNumber,
        [string]$QprJobNumber,  # Job number for sheet naming
        [string]$GroupBy  # "TL", "Assembly", or "Sequence"
    )

    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $false
    $xlApp.DisplayAlerts = $false
    $xlApp.ScreenUpdating = $false

    try {
        $wb = $xlApp.Workbooks.Add()
        while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

        $made = @(Add-QualityProcessSheets -Workbook $wb -MasterData $MasterData -QprJobNumber $QprJobNumber -GroupBy $GroupBy -ReuseFirstSheet $true)[-1]
        if (-not $made) { $wb.Close($false); return $null }

        # Save
        $safeProjectName = $ProjectName -replace '[\\/:*?"<>|]', '_'
        $stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
        $qprPath = Join-Path (Get-Location) "$safeProjectName - Job $JobNumber - QPR $stamp.xlsx"
        if (Test-Path $qprPath) { Remove-Item $qprPath -Force -ErrorAction SilentlyContinue }
        $wb.Sheets.Item(1).Activate()
        $wb.SaveAs($qprPath)
        $wb.Close($false)

        Add-OutputFile $qprPath
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


# --- QPR OPTIONS (asked here so the QPR can also become tabs in this workbook) ---
Write-Section "Quality Process Report (QPR)"
Write-Host "  [N] No (default): skip it"
Write-Host "  [Y] Yes: shop-floor inspection sheets, added as tabs to the BOM and saved as their own workbook"
Write-Host ""
$qprDefaultYes = ($updateMode -and $prevState -and $prevState.Qpr)
$qprChoice = Read-Host "Generate Quality Process Report (QPR)? [Y]es / [N]o (default: $(if ($qprDefaultYes) { 'Yes, as last time' } else { 'No' }))"
if (-not $qprChoice.Trim() -and $qprDefaultYes) { $qprChoice = 'y' }
if ($qprChoice -match '^(y|yes)$') {
    Write-Host ""
    if ($ourJobNumber) {
        # Already asked once for the labels / work orders: the QPR uses the same job number (sheet prefix and the Job: box)
        $qprJobNumber = $ourJobNumber
        Write-Host "  QPR job name: $qprJobNumber  (your job number from above, e.g. '$qprJobNumber - TL1282')"
    }
    else {
        Write-Host "  Job name is the sheet prefix and the Job: box, e.g. '12-345' -> '12-345 - TL1282'"
        $qprJobInput = Read-Host "Job name (default: $jobNumber)"
        $qprJobNumber = if ($qprJobInput) { $qprJobInput } else { $jobNumber }
    }

    Write-Host ""
    Write-Host "  [T] TL (default): one sheet per transmittal     $qprJobNumber - TL1282 | $qprJobNumber - TL1296"
    Write-Host "  [S] Sequence: one sheet per TL + sequence       $qprJobNumber - TL1296 - 5131 | $qprJobNumber - TL1296 - 5132"
    Write-Host ""
    $qprGroupChoice = Read-Host "Organize sheets by [T]L number (default) or [S]equence?"
    $qprGroupBy = if ($qprGroupChoice -match '^(s|sequence)$') { "Sequence" } else { "TL" }
    if (-not $qprGroupChoice.Trim() -and $updateMode -and $prevState -and $prevState.QprBy -eq 'Sequence') { $qprGroupBy = "Sequence" }
}

# --- TAB: FOLDER LABELS (only when a Required Assm workbook was chosen) ---
if ($requiredAssmPath) {
    Write-Log "Building Folder Labels tab from $([System.IO.Path]::GetFileName($requiredAssmPath))..."
    try {
        $labelData = @($allWorkbookTLs)   # read once, before the Master BOM was built
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
            $tlLabelData = @($labelData | Sort-Object Number)
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
                        $hits = @(Get-RevisedTransmittals -Tl $tl -AllTLs $allWorkbookTLs -Sequences @($seq))
                        if ($hits.Count -gt 0) { $notes = (@("REVISES $(($hits | ForEach-Object { $_.TL }) -join ', ')", $coat) | Where-Object { $_ }) -join " - " }
                    }
                    [PSCustomObject]@{ TL = $tl.TL; IsRev = $tl.IsRev; Sequence = $seq; DueDate = $due; Notes = $notes }
                }
            })
            [void](Add-FolderLabelSheet -Workbook $workbook -Labels $labelData)
            foreach ($l in $labelData) {
                $line1 = (@($ourJobNumber, $l.TL, $l.Sequence) | Where-Object { $_ }) -join " - "
                Write-Log "  Label: $line1 | $($l.Notes) | $($l.DueDate)" "Cyan"
            }
            Write-Log "Folder Labels tab ready: $($labelData.Count) label(s). Print that tab on the Dymo (3.5 x 1.125 in)." "Green"
            [void](Add-TlLabelSheet -Workbook $workbook -TLs $tlLabelData -AllTLs $allWorkbookTLs)
            Write-Log "TL Labels tab ready: $($tlLabelData.Count) label(s) (one per transmittal)." "Green"
        }
    }
    catch {
        Write-Log "Could not build the Folder Labels tab: $($_.Exception.Message)" "Red"
    }
}

# --- TAB: TERRA WORK ORDER (KSS mode: one work order per sequence) ---
if ($modeChoice -ne "2") {
    Write-Log "Building TERRA Work Order tab..."
    try {
        [void](Add-WorkOrderSheet -Workbook $workbook -Pages $workOrderPages -LogoBase64 $script:TerraLogoBase64)
        Write-Log "TERRA Work Order tab ready: $($workOrderPages.Count) work order(s), one per TL and sequence." "Green"
    }
    catch {
        Write-Log "Could not build the TERRA Work Order tab: $($_.Exception.Message)" "Red"
    }
}

# --- TABS: QPR (one per TL, or per TL + sequence; Job: box follows Settings, due / paint line follows Sequence Notes) ---
if ($qprChoice -match '^(y|yes)$') {
    Write-Log "Building QPR tab(s) in the BOM workbook..."
    try {
        $qprTabs = @(Add-QualityProcessSheets -Workbook $workbook -MasterData $masterData -QprJobNumber $qprJobNumber -GroupBy $qprGroupBy -Linked $true -CustomerJob $jobNumber -OurJob $ourJobNumber -HasSequenceNotes ($modeChoice -ne "2"))[-1]
        Write-Log "QPR tab(s) ready: $qprTabs ($qprGroupBy)." "Green"
    }
    catch {
        Write-Log "Could not build the QPR tabs: $($_.Exception.Message)" "Red"
    }
}

# --- STATISTICS EXTRAS (below the existing tables; schedule needs the folder rows from the Required Assm workbook) ---
try {
    $scheduleFolders = @($labelData | ForEach-Object { if ($_) { [PSCustomObject]@{ TL = $_.TL; Sequence = $_.Sequence; DueDate = $_.DueDate } } })
    if ($script:UpdateSchedDue) {
        foreach ($sf in $scheduleFolders) {
            $sk = "$($sf.TL) / $($sf.Sequence)"
            if ($sf -and $script:UpdateSchedDue.ContainsKey($sk)) { $sf.DueDate = $script:UpdateSchedDue[$sk] }
        }
    }
    Add-DashboardExtras -Dashboard $dashboard -MasterData $masterData -Folders $scheduleFolders
}
catch {
    Write-Log "Could not build the Statistics schedule section: $($_.Exception.Message)" "Red"
}
if ($script:UpdateMode -and $script:UpdateLogRows.Count -gt 0) {
    try { [void](Add-UpdateLogSheet -Workbook $workbook -Rows $script:UpdateLogRows -PreviousName ([System.IO.Path]::GetFileName($prevBom.Path))) }
    catch { Write-Log "Could not build the Update Log tab: $($_.Exception.Message)" "Yellow" }
}
# Sheet presentation setup
$dashboard.Move($workbook.Sheets.Item(1))
$workbook.Sheets.Item("Settings").Move([Type]::Missing, $dashboard)   # Settings sits right after Statistics
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
        if ($script:UpdateMode) {
            # Update: the previous workbook (and its record) is kept under a new name, never overwritten
            $backup = Join-Path (Get-Location) ([System.IO.Path]::GetFileNameWithoutExtension($exportPath) + " (before update " + (Get-Date -Format 'yyyyMMdd-HHmmss') + ").xlsx")
            Move-Item -LiteralPath $exportPath -Destination $backup -ErrorAction Stop
            $oldStateFile = Get-KissStatePath $exportPath
            if (Test-Path -LiteralPath $oldStateFile) { Move-Item -LiteralPath $oldStateFile -Destination (Get-KissStatePath $backup) -Force -ErrorAction SilentlyContinue }
            Write-Log "The previous workbook was kept as: $([System.IO.Path]::GetFileName($backup))" "Yellow"
        }
        else { Remove-Item $exportPath -Force -ErrorAction Stop }
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
Add-OutputFile $exportPath
if ($modeChoice -ne "2" -and $script:FreshSeeds) {
    # The "last build" record: which .kss files and what the inputs said, so the next update can tell your edits from input changes
    try {
        Save-KissState -WorkbookPath $exportPath -Fresh $script:FreshSeeds -KssFiles $kssFiles -Merge $mergeDuplicates `
            -Qpr ($qprChoice -match '^(y|yes)$') -QprBy "$qprGroupBy"
    }
    catch { Write-Log "Could not save the last-build record (updates will ask more questions): $($_.Exception.Message)" "Yellow" }
}
$workbook.Close($false)
$excel.Quit()

Write-Log "SUCCESS! Formatted workbook saved directly to:" "Green"
Write-Log " -> $exportPath" "Yellow"

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
[gc]::collect()
[gc]::WaitForPendingFinalizers()


# 5. GENERATE THE SHOP SCRIPTS (DXF SORT, QUANTITIES, GRADES): one set for the whole run.
#    Each script asks at runtime which transmittals / sequences are being nested now and only handles those.
Write-Log "Generating shop scripts..." "Yellow"
if (-not (Test-Path variable:tlNames)) { $tlNames = @{} }
$plateRows = @(foreach ($item in @($masterData | Where-Object { $_.Category -eq "Plate" -and $_.PartMark })) {
    # Same transmittal the Master BOM, labels and work orders use (a REV is placed under its own TL)
    $placed = $asmByKey["$($item.TL)|$($item.Assembly)"]
    $tlName = if ($placed -and $tlNames.ContainsKey($placed.TlNumber)) { $tlNames[$placed.TlNumber] }
              elseif ($placed) { "TL$($placed.TlNumber)" }
              elseif ($item.TL) { "$($item.TL)" }
              else { [System.IO.Path]::GetFileNameWithoutExtension("$($item.SourceFile)") }
    $assm = "$($item.Assembly)"
    [PSCustomObject]@{
        Mark = "$($item.PartMark)"; Size = "$($item.Size)"; Grade = "$($item.Grade)"
        TL = ($tlName -replace '[\\/:*?"<>|%!]', '_')
        Seq = if ($assm.Length -ge 4) { $assm.Substring(0, 4) } else { $assm }
        Qty = [int][Math]::Round([double]$item.Qty)
    }
})
$scriptLabel = if ($allTLs.Count -gt 0) { ($allTLs -join '_') } else { "ALL-KSS" }
if ($scriptLabel.Length -gt 60) { $scriptLabel = "$($allTLs[0])-$($allTLs[-1])" }
$scriptLabel = $scriptLabel -replace '[\\/:*?"<>|]', '_'
$shopScriptSet = Write-ShopScripts -PlateRows $plateRows -ProjectName $projectName -JobNumber $jobNumber -Label $scriptLabel
if ($shopScriptSet) { Write-Log "Shop scripts ready. Run one: it asks which transmittals / sequences you are nesting now." "Green" }
else { Write-Log "No plates in the output; skipping shop scripts." "Yellow" }


# =====================================================================
# 6. OPTIONAL: QUALITY PROCESS REPORT (standalone workbook; the questions were asked before the BOM workbook was built)
# =====================================================================

if ($qprChoice -match '^(y|yes)$') {
    Write-QualityProcessReport -MasterData $masterData -ProjectName $projectName -JobNumber $jobNumber -QprJobNumber $qprJobNumber -GroupBy $qprGroupBy
}

# =====================================================================
# 7. DONE: list what was made and offer to open the folder
# =====================================================================
Write-Section "Done"
$doneFolder = (Get-Location).Path
Write-Host "  Saved in: $doneFolder" -ForegroundColor Green
$groups = [ordered]@{
    'Workbooks'      = @($script:OutputFiles | Where-Object { $_ -like '*.xlsx' })
    'Filtered KISS'  = @($script:OutputFiles | Where-Object { $_ -like '*.kss' })
    'Shop scripts'   = @($script:OutputFiles | Where-Object { $_ -like '*.bat' })
    'Update report'  = @($script:OutputFiles | Where-Object { $_ -like '*.txt' })
}
foreach ($g in $groups.Keys) {
    if ($groups[$g].Count -gt 0) {
        Write-Host "  $g" -ForegroundColor Yellow
        foreach ($f in $groups[$g]) { Write-Host "    - $([System.IO.Path]::GetFileName($f))" }
    }
}
Write-Host ""
$openChoice = Read-Host "Open the folder now? [Y]es (default) / [N]o"
if ($openChoice -notmatch '^(n|no)$') { try { Start-Process explorer.exe -ArgumentList "`"$doneFolder`"" } catch { } }

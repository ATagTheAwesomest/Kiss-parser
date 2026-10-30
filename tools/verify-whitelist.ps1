function Write-Log {
    param([string]$Message, [string]$Color = "Cyan")
    Write-Host "[$(Get-Date -Format 'HH:mm:ss.fff')] $Message" -ForegroundColor $Color
}

function Parse-AssemblyWhitelist {
    param([string]$InputText)

    if ([string]::IsNullOrWhiteSpace($InputText)) { return @() }

    $raw = $InputText.Trim()
    if ($raw.StartsWith('@')) {
        $path = $raw.Substring(1).Trim()
        if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path (Get-Location) $path }
        if (Test-Path $path) {
            $raw = [System.IO.File]::ReadAllText($path)
        }
        else {
            Write-Log "Assembly whitelist file not found: $path" "Red"
            return @()
        }
    }

    $names = @($raw -split '[,\r\n;|]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { ($_ -replace '"', '').Trim() } | Select-Object -Unique)
    return @($names)
}

function Apply-AssemblyWhitelist {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$MasterData,
        [System.Collections.Generic.List[PSCustomObject]]$LaborData,
        [hashtable]$AssemblyDrawings,
        [string]$AssemblyListText
    )

    $names = @(Parse-AssemblyWhitelist -InputText $AssemblyListText)
    if ($names.Count -eq 0) {
        return [PSCustomObject]@{
            Applied = $false
            MasterData = $MasterData
            LaborData = $LaborData
            AssemblyDrawings = $AssemblyDrawings
            Included = @()
            Missing = @()
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

    $filteredMaster = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $MasterData) {
        if ($allowed.Contains($item.Assembly.Trim())) { $filteredMaster.Add($item) }
    }

    $filteredLabor = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $LaborData) {
        if ($allowed.Contains($item.Assembly.Trim())) { $filteredLabor.Add($item) }
    }

    $filteredDrawings = @{}
    foreach ($key in $AssemblyDrawings.Keys) {
        if ($allowed.Contains($key.Trim())) { $filteredDrawings[$key] = $AssemblyDrawings[$key] }
    }

    return [PSCustomObject]@{
        Applied = $true
        MasterData = $filteredMaster
        LaborData = $filteredLabor
        AssemblyDrawings = $filteredDrawings
        Included = @($names)
        Missing = @($missing)
    }
}

$master = [System.Collections.Generic.List[object]]::new()
foreach ($a in @('A100','A200','A300')) { $master.Add([pscustomobject]@{ Assembly = $a }) }
$labor = [System.Collections.Generic.List[object]]::new()
foreach ($a in @('A100','A300')) { $labor.Add([pscustomobject]@{ Assembly = $a }) }
$drawings = @{ 'A100' = 'DWG1'; 'A200' = 'DWG2'; 'A300' = 'DWG3' }
$result = Apply-AssemblyWhitelist -MasterData $master -LaborData $labor -AssemblyDrawings $drawings -AssemblyListText "A100, A300`nZ999"
if ($result.MasterData.Count -ne 2) { throw "expected 2 master records, got $($result.MasterData.Count)" }
if ($result.LaborData.Count -ne 2) { throw "expected 2 labor records, got $($result.LaborData.Count)" }
if ($result.AssemblyDrawings.Count -ne 2) { throw "expected 2 drawings, got $($result.AssemblyDrawings.Count)" }
if ($result.Missing.Count -ne 1 -or $result.Missing[0] -ne 'Z999') { throw "expected missing list to contain Z999" }
Write-Host 'PASS: whitelist keeps A100 and A300 and flags Z999'

#requires -Version 5.1
<#
.SYNOPSIS
  Does a member have data? Lists the Scenario / Version / Year combinations where it does,
  without returning any amount.

.DESCRIPTION
  Builds one ad hoc sheet in the scratch workbook from the cube's default grid: the member in
  a column header, Scenario and Version zoomed to bottom level on rows, Year zoomed to bottom
  level on columns, every other dimension at its top. Refreshes with data WITHOUT row
  suppression (suppression made the refresh fail on Planning cubes), then counts with Excel
  formulas that return text, so "no data" shown as 0 or as an empty cell is handled.
  The working sheet is deleted at the end unless -Keep.

  Run it once per cube: on the input cube (BSO) and on the reporting cube (ASO) when there is
  one. BSO upper levels may be unaggregated (zeros at the top do not prove "no data"); an ASO
  aggregates on the fly but only holds what was last pushed to it.

  -Pov "Dim#Member;Dim#Member" fixes the dimensions that are not explored. Needed on a BSO
  cube without a reporting cube: the roots of Account, Period, Currency and similar are
  usually label-only or not aggregated, so leaving them at the top gives a false "no data".
  Take stored members from the user or from a form that shows the data.

.EXAMPLE
  member-data.ps1 -Url https://host/HyperionPlanning/SmartView -App MyApp -Cube MyRPT -CubeType ASO -Dimension CostCenter -Member 1234
.EXAMPLE
  member-data.ps1 -Profile "Client A - Planning TEST" -Dimension CostCenter -Member 1234
  Takes url, application, cube (reportingCube if set, else workCube), cube type and
  dataCheckPov from the environment profile in connections.json. Explicit parameters win.
#>
[CmdletBinding()]
param(
    [Alias('Profile')] [string] $ProfileName = '',
    [string] $Url = '',
    [string] $App = '',
    [string] $Cube = '',
    [ValidateSet('', 'BSO', 'ASO')] [string] $CubeType = '',
    [Parameter(Mandatory)] [string] $Dimension,
    [Parameter(Mandatory)] [string] $Member,
    [string] $Pov = '',
    [string] $Sheet = 'MemberData',
    [switch] $Keep
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$sv = Join-Path $PSScriptRoot 'sv.ps1'
. (Join-Path $PSScriptRoot 'SvUrl.ps1')
if ($ProfileName) {
    $pf = Join-Path $env:LOCALAPPDATA 'smart-view-skill\connections.json'
    $p = @(Get-Content $pf -Raw -Encoding UTF8 | ConvertFrom-Json | ForEach-Object { $_ }) | Where-Object { $_.name -eq $ProfileName } | Select-Object -First 1
    if (-not $p) { [ordered]@{ ok = $false; error = "Profile '$ProfileName' not found in $pf" } | ConvertTo-Json; exit 1 }
    if (-not $Url) { $Url = $p.url }
    if (-not $App) { $App = $p.application }
    if (-not $Cube) {
        if ($p.reportingCube) { $Cube = $p.reportingCube; if (-not $CubeType) { $CubeType = 'ASO' } }
        else { $Cube = $p.workCube }
    }
    if (-not $CubeType -and "$($p.cubeType)" -match '^(BSO|ASO)\b') { $CubeType = $Matches[1] }
    if (-not $Pov -and $Cube -ne $p.reportingCube -and $p.dataCheckPov) { $Pov = $p.dataCheckPov }
}
if (-not ($Url -and $App -and $Cube)) { [ordered]@{ ok = $false; error = 'Give -Profile, or -Url, -App and -Cube.' } | ConvertTo-Json; exit 1 }
$Url = ConvertTo-SvUrl $Url
$conn = "$Url|$(([uri]$Url).Host)|$App|$Cube"
$log = New-Object System.Collections.ArrayList

function J([string[]] $a) { (& $sv @a) | ConvertFrom-Json }
function Step([string] $name, $j) {
    $d = @($j.dialogs | Where-Object { $_ } | ForEach-Object { "$($_.text) -> $($_.action)" })
    $null = $log.Add([ordered]@{ step = $name; ok = [bool]$j.ok; rc = $j.rc; error = $j.error.message; dialogs = @($d) })
    return [bool]$j.ok
}
function Grid { @((& $sv read '' $Sheet '' 6000) -split "`n" | Select-Object -Skip 1 | ForEach-Object { , ($_ -split "`t") }) }
function Col([int] $n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
function Find([string] $label) {
    $g = Grid
    for ($r = 0; $r -lt $g.Count; $r++) { for ($k = 1; $k -lt $g[$r].Count; $k++) { if ($g[$r][$k].Trim() -eq $label) { return "$(Col $k)$($r + 1)" } } }
    return $null
}
function Finish($result) {
    if (-not $Keep) {
        try {
            $s = ((& $sv call Workbooks) | ConvertFrom-Json).data | Where-Object { $_.scratch }
            $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
            $xl.DisplayAlerts = $false; $xl.Workbooks.Item($s.name).Worksheets.Item($Sheet).Delete(); $xl.DisplayAlerts = $true
        } catch { }
    }
    $result.steps = $log
    $result | ConvertTo-Json -Depth 6
    exit 0
}

$st = (& $sv status) | ConvertFrom-Json
if (-not $st.bridgeLoaded) { [ordered]@{ ok = $false; error = 'Excel or the bridge is not running: see references/bridge-setup.md.' } | ConvertTo-Json; exit 1 }

$null = J @('call','Scratch',$Sheet,'true')
if (-not (Step 'connect' (J @('call','UIConnect','',$Sheet,$conn)))) { Finish ([ordered]@{ ok = $false; error = 'connect failed (sign-in needed or wrong app/cube)' }) }
# Member names, not aliases: with an alias table on, labels and typed names are aliases.
$null = Step 'alias table None' (J @('call','SetAliasTable','',$Sheet,'None'))
foreach ($o in @(@(33,'true'),@(6,'false'),@(7,'false'),@(41,'false'),@(42,'false'),@(16,'0'),@(5,'0'))) { $null = J @('call','SetOption',[string]$o[0],$o[1],'',$Sheet) }
if (-not (Step 'default grid' (J @('call','Retrieve','',$Sheet)))) { Finish ([ordered]@{ ok = $false; error = 'default grid failed' }) }

# Pivot onto the innermost row and column dimension cells of the default grid (it may have
# one or more dimensions on each axis; the others stay at their top member).
$g = Grid
$rowDim = ($g | ForEach-Object { $_[1].Trim() } | Where-Object { $_ } | Select-Object -Last 1)
$colDim = (($g[0] | Select-Object -Skip 2 | Where-Object { $_.Trim() } | Select-Object -First 1)).Trim()
foreach ($p in @(@('Scenario', $rowDim), @('Version', $rowDim), @('Year', $colDim), @($Dimension, $colDim))) {
    if ($p[0] -eq $p[1] -or (Find $p[0])) { continue }   # already on the grid
    $null = Step "pivot $($p[0])" (J @('call','PivotToGrid','',$Sheet,$p[0],(Find $p[1])))
}
$cell = Find $Dimension
if (-not $cell) { Finish ([ordered]@{ ok = $false; error = "dimension $Dimension not found on the grid" }) }
# Numeric member names must go in as text.
$null = J @('call','WriteTsv',$Sheet,$cell,"'$Member")
# Fixed members for the other dimensions: dimensions already on the grid get the member in
# their cell; POV dimensions are pivoted onto the column header first.
$fixed = @()
foreach ($pair in ($Pov -split ';' | Where-Object { $_.Trim() })) {
    $d, $m = $pair.Split('#', 2) | ForEach-Object { $_.Trim() }
    $at = Find $d
    if (-not $at) {
        $null = Step "pivot $d" (J @('call','PivotToGrid','',$Sheet,$d,(Find $colDim)))
        $at = Find $d
    }
    if ($at) { $null = J @('call','WriteTsv',$Sheet,$at,"'$m"); $fixed += "$d#$m" }
    else { $null = $log.Add([ordered]@{ step = "fix $d"; ok = $false; error = 'dimension not found' }) }
}
if (-not (Step 'place member' (J @('call','Retrieve','',$Sheet)))) { Finish ([ordered]@{ ok = $false; error = "member $Member not accepted (name, alias table?)" }) }
foreach ($d in 'Scenario', 'Version', 'Year') { $c = Find $d; if ($c) { $null = Step "zoom $d" (J @('call','ZoomIn','',$Sheet,$c,'2')) } }

# Checking for data is a request for data: turn navigate-without-data off and refresh.
$null = J @('call','SetOption','33','false','',$Sheet)
if (-not (Step 'refresh with data' (J @('call','Retrieve','',$Sheet)))) { Finish ([ordered]@{ ok = $false; error = 'refresh with data failed' }) }

# Locate the header block and the data block without assuming member names:
# data rows are those with a label in column A; the year header is the header row with the
# most distinct values; data columns start where the row labels end.
$g = Grid
$firstData = -1
for ($r = 0; $r -lt $g.Count -and $firstData -lt 0; $r++) { if ($g[$r].Count -gt 1 -and $g[$r][1].Trim()) { $firstData = $r } }
if ($firstData -lt 1) { Finish ([ordered]@{ ok = $false; error = 'grid layout not recognised' }) }
$fc = 0
for ($k = 1; $k -lt $g[$firstData].Count; $k++) { if (-not $g[$firstData][$k].Trim() -or $g[$firstData][$k] -eq '<n>') { $fc = $k; break } }
if ($fc -eq 0) { Finish ([ordered]@{ ok = $false; error = 'first data column not found' }) }
$lc = ($g[0..($firstData - 1)] | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum - 1
$yr = 0; $best = -1
for ($r = 0; $r -lt $firstData; $r++) {
    $n = @($g[$r][$fc..$lc] | Where-Object { $_.Trim() } | Sort-Object -Unique).Count
    if ($n -gt $best) { $best = $n; $yr = $r }
}
$c1 = Col $fc; $c2 = Col $lc; $yrRow = $yr + 1; $first = $firstData + 1; $last = $g.Count

# Global count first (SUMPRODUCT handles arrays in any Excel version).
$q = '"'
$gc = Col ($lc + 5)
$null = J @('call','WriteTsv',$Sheet,"${gc}1","=${q}nonzero=${q}&SUMPRODUCT(ISNUMBER(${c1}${first}:${c2}${last})*(${c1}${first}:${c2}${last}<>0))")
$global = [int](((& $sv read '' $Sheet "${gc}1:${gc}1" 1) -split "`n")[1].Split("`t")[1] -replace '\D', '')

# Per row: the headers of the non-zero cells, one IF per column (no array formula, so no
# implicit intersection).
$helper = Col ($lc + 3)
$formulas = for ($r = $first; $r -le $last; $r++) {
    $parts = for ($k = $fc; $k -le $lc; $k++) { $c = Col $k; "IF(AND(ISNUMBER($c$r),N($c$r)<>0),$c`$$yrRow&${q},${q},${q}${q})" }
    '=' + ($parts -join '&')
}
$tmp = Join-Path $env:TEMP "svb_member_$PID.tsv"
[IO.File]::WriteAllText($tmp, ($formulas -join "`n"), [Text.UTF8Encoding]::new($false))
$null = J @('call','WriteTsv',$Sheet,"$helper$first","@$tmp")
Remove-Item $tmp -ErrorAction SilentlyContinue

# Row labels: every label column; columns that never change (dimension tops) are dropped.
$out = @((& $sv read '' $Sheet "A${first}:${helper}${last}" 6000) -split "`n" | Select-Object -Skip 1 | ForEach-Object { , ($_ -split "`t") })
$hCol = $lc + 3
$labelCols = @(1..($fc - 1) | Where-Object { $i = $_; @($out | ForEach-Object { $_[$i].Trim() } | Sort-Object -Unique).Count -gt 1 })
$combos = [ordered]@{}
foreach ($c in $out) {
    if ($c.Count -le $hCol) { continue }
    $yrs = $c[$hCol].Trim().TrimEnd(',')
    if (-not $yrs) { continue }
    $key = (@($labelCols | ForEach-Object { $c[$_].Trim() }) -join ' / ')
    $combos[$key] = (@($combos[$key] -split ',') + $yrs.Split(',') | Where-Object { $_ } | Sort-Object -Unique) -join ','
}
if ($global -gt 0 -and $combos.Count -eq 0) {
    Finish ([ordered]@{ ok = $false; error = "inconsistent: $global non-zero cells but no combination parsed; do not conclude 'no data'" })
}
Finish ([ordered]@{
    ok          = $true
    cube        = "$App/$Cube"
    member      = "$Dimension#$Member"
    hasData     = ($global -gt 0)
    nonZeroCells = $global
    fixedPov    = $fixed
    cubeType    = $(if ($CubeType) { $CubeType } else { 'unknown' })
    warning     = $(if ($global -eq 0 -and -not $fixed -and $CubeType -ne 'ASO') { 'Zero with every other dimension at its root. On a BSO cube without a reporting cube this does NOT prove "no data": roots are often label-only. Rerun with -Pov and stored members (ask the user).' } else { $null })
    combinations = @($combos.Keys | ForEach-Object { [ordered]@{ rowMembers = $_; columns = $combos[$_] } })
    note        = 'Other dimensions at their top member. On a BSO cube zeros at the top may hide level-0 data; confirm on the ASO cube. A top member (for example the Scenario root) can show nothing while its children have data. Qualified members such as [Parent].[Child] are aggregated views of the same data.'
})

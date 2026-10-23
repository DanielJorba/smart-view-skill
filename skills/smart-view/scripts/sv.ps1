#requires -Version 5.1
<#
.SYNOPSIS
  Command-line bridge between an AI agent and Oracle Smart View in the user's running Excel.

.DESCRIPTION
  Attaches to the user's Excel (start-excel launches it like the user would; it never
  closes Excel or saves workbooks) and calls the read-only SVB_* wrappers of the SmartViewBridge.xlam add-in
  through Application.Run. Every command prints JSON, except "read", which prints TSV.

  Commands
    help                         This text.
    status                       Excel, Smart View and bridge state, Smart View version.
    version                      Installed Smart View version and extensions.
    connections                  Saved connection profiles (connections.json) and the
                                 Smart View quick-connect URL list.
    call <Wrapper> [args...]     Run SVB_<Wrapper> in the bridge, for example
                                 "call GetPOV '' Sheet1". Only SVB_* wrappers can be called.
    read <workbook> <sheet> [range] [maxRows]
                                 Numeric cells come back as <n>: figures are hidden by
                                 default in every read (read, formulas, hs-audit,
                                 ReadTsv, CellValue, AuditFunctions).
                                 Cell values as TSV. Empty workbook = scratch workbook.
                                 Range defaults to the used range. Adds the indent level
                                 of the first column as column 0.
    formulas <workbook> <sheet> [range] [maxCells]
                                 Formula cells as TSV: address, formula, shown text.
    hs-audit <workbook> <sheet> [range] [maxCells]
                                 Audit of Smart View cell functions (HsGetValue...):
                                 connections, dimensions and members actually used
                                 (arguments evaluated by Excel), error cells, HsSetValue
                                 cells and broken add-in links.
    install                      Copy the prebuilt bridge/SmartViewBridge.xlam to the
                                 user's AddIns folder and load it in Excel. Persistent
                                 Excel setting: ask the user first.
    uninstall                    Unload the bridge add-in and remove the installed copy.
    drop-sheet <sheet> [...]     Delete working sheets of the scratch workbook (never others).
    private-connections [name]   The user's private connections (name, URL, application, cube,
                                 provider, alias table). Read-only; never the credential.
    start-excel                  Start Excel as the user would (visible, Smart View loaded) if
                                 it is not running.
    open <path> [edit]           Open a saved workbook, read-only unless "edit". Never saves.
    new-workbook                 Add a new empty workbook for the user. Never saves.
    close <workbook>             Close without saving a workbook the skill opened or created
                                 (open, new-workbook, open-report). Refuses any other.

  Dialog watcher: while a call runs, Smart View / Excel pop-ups are read in real time and
  reported under "dialogs". -Reuse sets the answer to "worksheet contains data from prior
  operation" (cancel by default; clear only on the scratch workbook). Single-button
  messages are dismissed after reading unless -KeepMessages. Sign-in windows and
  confirmations that change the server are never touched.

  An argument written as @<path> is replaced by the content of that file (UTF-8),
  for long inputs such as WriteTsv grids.

  Passwords are never accepted, stored or printed. Sign-in happens in Smart View's own
  dialog, typed by the user.
#>
[CmdletBinding()]
param(
    [ValidateSet('cancel', 'contents', 'contentsandpov', 'clear')] [string] $Reuse = 'cancel',
    [switch] $KeepMessages,
    [switch] $ShowFigures,
    [Parameter(Position = 0)] [string] $Command = 'help',
    [Parameter(Position = 1, ValueFromRemainingArguments = $true)] [string[]] $Rest
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$BridgeName = 'SmartViewBridge.xlam'
. (Join-Path $PSScriptRoot 'DialogWatch.ps1')
. (Join-Path $PSScriptRoot 'SvUrl.ps1')
# Figures are hidden by default. Showing them needs a personal machine setting the user
# sets themselves (SMARTVIEW_SKILL_ALLOW_FIGURES=1) AND -ShowFigures on the call.
$AllowFigures = (($env:SMARTVIEW_SKILL_ALLOW_FIGURES -eq '1') -or ([Environment]::GetEnvironmentVariable('SMARTVIEW_SKILL_ALLOW_FIGURES', 'User') -eq '1')) -and $ShowFigures.IsPresent
$DataDir = Join-Path $env:LOCALAPPDATA 'smart-view-skill'
$SvRegistry = 'HKCU:\Software\Hyperion Solutions\HyperionSmartView'

# Workbooks the skill opened or created in this Excel session: only these may be closed.
$OpenedList = Join-Path $DataDir 'opened-by-skill.txt'
function Add-Opened([string] $Name) {
    if (-not (Test-Path $DataDir)) { $null = New-Item -ItemType Directory -Path $DataDir }
    # One line per workbook: <Excel process id>|<name>, so a name reused in a later Excel session never matches.
    $xpid = (Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -ne 0 | Select-Object -First 1).Id
    Add-Content -LiteralPath $OpenedList -Value "$xpid|$Name" -Encoding UTF8
}

function Write-Json($Object) {
    $Object | ConvertTo-Json -Depth 10
}

function Stop-WithError([string] $Message) {
    Write-Json ([ordered]@{ ok = $false; error = $Message })
    exit 1
}

function Get-HResult($ErrorRecord) {
    $e = $ErrorRecord.Exception
    while ($e) {
        if ($e -is [System.Runtime.InteropServices.COMException]) { return $e.HResult }
        $e = $e.InnerException
    }
    return $null
}

# Excel rejects COM calls while a cell is in edit mode or a dialog is open.
# 0x80010001 RPC_E_CALL_REJECTED, 0x8001010A RPC_E_SERVERCALL_RETRYLATER.
function Invoke-WithRetry([scriptblock] $Action, [int] $Tries = 40) {
    for ($i = 1; $i -le $Tries; $i++) {
        try {
            return & $Action
        }
        catch {
            $hr = Get-HResult $_
            if (($hr -eq -2147418111 -or $hr -eq -2147417846) -and $i -lt $Tries) {
                Start-Sleep -Milliseconds 500
                continue
            }
            throw
        }
    }
}

function Get-Excel {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    catch {
        Stop-WithError 'Excel is not running. Start it with "sv.ps1 start-excel" (tell the user first).'
    }
}

function Test-BridgeLoaded($Excel) {
    try {
        $null = Invoke-WithRetry { $Excel.Workbooks.Item($BridgeName) }
        return $true
    }
    catch {
        return $false
    }
}

function Test-SmartViewLoaded($Excel) {
    try {
        return [bool](Invoke-WithRetry { $Excel.COMAddIns.Item('Hyperion.CommonAddin').Connect })
    }
    catch {
        return $false
    }
}

function Get-SmartViewDll {
    try {
        $clsid = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Hyperion.CommonAddin\CLSID').'(default)'
        $dll = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\InprocServer32").'(default)'
        if ($dll -and (Test-Path $dll)) { return $dll }
    }
    catch { }
    $fallback = 'C:\Oracle\SmartView\bin\HsAddin.dll'
    if (Test-Path $fallback) { return $fallback }
    return $null
}

function Get-SmartViewVersion {
    $dll = Get-SmartViewDll
    $extensions = @()
    $extDir = Join-Path $env:APPDATA 'Oracle\SmartView\Extensions\Bin'
    if (Test-Path $extDir) {
        $extensions = @(Get-ChildItem $extDir -Directory | Select-Object -ExpandProperty Name)
    }
    [ordered]@{
        installed   = [bool]$dll
        version     = if ($dll) { (Get-Item $dll).VersionInfo.ProductVersion } else { $null }
        binFolder   = if ($dll) { Split-Path $dll } else { $null }
        smartviewBas = if ($dll) { Join-Path (Split-Path $dll) 'smartview.bas' } else { $null }
        extensions  = $extensions
    }
}

function Resolve-Argument([string] $Value) {
    if ($Value -and $Value.StartsWith('@')) {
        $path = $Value.Substring(1)
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            return [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $path), [System.Text.Encoding]::UTF8)
        }
    }
    # Inline connection "URL|server|app|cube" or a bare URL to an EPM Cloud host: fix the URL part.
    if ($Value -match '^https?://[^|/]+\.oraclecloud\.com') {
        $parts = $Value.Split('|')
        $parts[0] = ConvertTo-SvUrl $parts[0]
        return ($parts -join '|')
    }
    return $Value
}

# Replaces numeric cell values with <n>, so structure can be read without figures.
function Hide-Numbers([string] $Text) {
    if ($null -eq $Text) { return $Text }
    if ($Text.Trim() -match '^[-+(]?[\d.,\s]+%?\)?([eE][-+]?\d+)?$' -and $Text -match '\d') { return '<n>' }
    return $Text
}

function Hide-NumbersInTsv([string] $Tsv) {
    $lines = $Tsv -split "`n"
    $rows = @(for ($i = 1; $i -lt $lines.Count; $i++) { , ($lines[$i] -split "`t") })
    # Grid layout: header rows on top (column A empty), then rows labelled in column A.
    # Leading columns with nothing above the first labelled row hold row members (numeric
    # codes such as cost centres or entities): keep them. Everything else is masked.
    $keep = 0
    $r0 = -1
    for ($r = 0; $r -lt $rows.Count -and $r0 -lt 0; $r++) { if ($rows[$r].Count -gt 1 -and $rows[$r][1].Trim()) { $r0 = $r } }
    if ($r0 -gt 0) {
        $width = ($rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
        for ($j = 1; $j -lt $width; $j++) {
            $above = @(for ($r = 0; $r -lt $r0; $r++) { if ($rows[$r].Count -gt $j) { $rows[$r][$j].Trim() } }) | Where-Object { $_ }
            if ($above) { break }
            $keep = $j
        }
        if ($keep -ge $width - 1) { $keep = 0 }   # no column header at all: not a grid, mask all
    }
    $grid = ($r0 -gt 0 -and $keep -gt 0)
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $cells = $rows[$r]
        # In a grid only the data block (below the headers, right of the row labels) is masked.
        for ($j = 1; $j -lt $cells.Count; $j++) { if (-not $grid -or ($r -ge $r0 -and $j -gt $keep)) { $cells[$j] = Hide-Numbers $cells[$j] } }
        $lines[$r + 1] = $cells -join "`t"
    }
    return ($lines -join "`n")
}

function Invoke-Bridge($Excel, [string] $Wrapper, [string[]] $Arguments) {
    if ($Wrapper -notmatch '^[A-Za-z0-9_]+$') { Stop-WithError "Invalid wrapper name: $Wrapper" }
    $macro = "'$BridgeName'!SVB_$Wrapper"
    $runArgs = @($macro) + @($Arguments | Where-Object { $null -ne $_ } | ForEach-Object { Resolve-Argument $_ })
    Invoke-WithRetry {
        $Excel.GetType().InvokeMember('Run', [System.Reflection.BindingFlags]::InvokeMethod, $null, $Excel, [object[]]$runArgs)
    }
}

switch ($Command.ToLowerInvariant()) {
    'help' {
        Get-Help $PSCommandPath -Detailed | Out-String | Write-Output
    }

    'version' {
        Write-Json (Get-SmartViewVersion)
    }

    'status' {
        $excelRunning = $true
        try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
        catch { $excelRunning = $false }

        $result = [ordered]@{
            ok              = $true
            excelRunning    = $excelRunning
            excelVersion    = $null
            smartViewLoaded = $false
            bridgeLoaded    = $false
            bridgePing      = $null
            figures         = if ([Environment]::GetEnvironmentVariable('SMARTVIEW_SKILL_ALLOW_FIGURES', 'User') -eq '1') { 'visible on explicit request' } else { 'hidden' }
            smartView       = Get-SmartViewVersion
        }
        if ($excelRunning) {
            $result.excelVersion = Invoke-WithRetry { $excel.Version }
            $result.smartViewLoaded = Test-SmartViewLoaded $excel
            $result.bridgeLoaded = Test-BridgeLoaded $excel
            if ($result.bridgeLoaded) {
                try { $result.bridgePing = (Invoke-Bridge $excel 'Ping' @()) | ConvertFrom-Json }
                catch { $result.bridgePing = "Ping failed: $($_.Exception.Message)" }
            }
        }
        Write-Json $result
    }

    'connections' {
        $profiles = @()
        $file = Join-Path $DataDir 'connections.json'
        if (Test-Path $file) {
            # ConvertFrom-Json in PowerShell 5.1 returns a JSON array as one object: unroll it.
            $profiles = @(Get-Content $file -Raw -Encoding UTF8 | ConvertFrom-Json | ForEach-Object { $_ })
        }
        $quick = @()
        try {
            $raw = (Get-ItemProperty "$SvRegistry\QuickConnect" -ErrorAction Stop).URLs
            # Stored as <URLs><url>...</url>...</URLs>; older versions used a plain list.
            if ($raw) { $quick = @([regex]::Matches($raw, 'https?://[^<>;|\s]+') | ForEach-Object { $_.Value }) }
        }
        catch { }
        Write-Json ([ordered]@{
            ok              = $true
            profilesFile    = $file
            profiles        = $profiles
            quickConnectUrls = $quick
        })
    }

    'call' {
        if (-not $Rest -or $Rest.Count -lt 1) { Stop-WithError 'Usage: call <Wrapper> [args...]' }
        $excel = Get-Excel
        if (-not (Test-BridgeLoaded $excel)) {
            Stop-WithError "$BridgeName is not loaded in Excel. See references/bridge-setup.md."
        }
        $wrapper = $Rest[0]
        $arguments = if ($Rest.Count -gt 1) { $Rest[1..($Rest.Count - 1)] } else { @() }
        if ($Reuse -eq 'clear' -and $arguments.Count -ge 1 -and $arguments[0]) {
            Stop-WithError '-Reuse clear is only allowed on the scratch workbook (empty workbook argument).'
        }
        $pidX = (Get-Process EXCEL | Where-Object MainWindowHandle -ne 0 | Select-Object -First 1).Id
        $watch = Start-DialogWatch $pidX @{ Reuse = $Reuse; DismissMessages = (-not $KeepMessages) }
        try { $out = Invoke-Bridge $excel $wrapper $arguments }
        catch { $out = (@{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress) }
        $events = @(Stop-DialogWatch $watch)
        if (-not $AllowFigures) {
            switch -Regex ($wrapper) {
                '^(?i)ReadTsv$' { $out = Hide-NumbersInTsv $out }
                '^(?i)ReadFormulas$' { $out = (($out -split "`n") | ForEach-Object { $c = $_ -split "`t"; if ($c.Count -ge 3) { $c[2] = Hide-Numbers $c[2] }; $c -join "`t" }) -join "`n" }
                '^(?i)(CellValue|AuditFunctions)$' {
                    try {
                        $o = $out | ConvertFrom-Json
                        if ($wrapper -ieq 'CellValue' -and $null -ne $o.data) { $o.data = Hide-Numbers ([string]$o.data) }
                        if ($wrapper -ieq 'AuditFunctions') { foreach ($c in @($o.data.calls)) { $c.text = Hide-Numbers ([string]$c.text) } }
                        $out = $o | ConvertTo-Json -Depth 8 -Compress
                    } catch { }
                }
            }
        }
        if ($events.Count -gt 0) {
            try { $obj = $out | ConvertFrom-Json; $obj | Add-Member -NotePropertyName dialogs -NotePropertyValue @($events) -Force; $out = $obj | ConvertTo-Json -Depth 6 -Compress }
            catch { $out = $out + "`n" + (@{ dialogs = $events } | ConvertTo-Json -Depth 6 -Compress) }
        }
        Write-Output $out
    }

    'read' {
        $excel = Get-Excel
        if (-not (Test-BridgeLoaded $excel)) {
            Stop-WithError "$BridgeName is not loaded in Excel. See references/bridge-setup.md."
        }
        $wb = if ($Rest.Count -ge 1) { $Rest[0] } else { '' }
        $sh = if ($Rest.Count -ge 2) { $Rest[1] } else { '' }
        $rng = if ($Rest.Count -ge 3) { $Rest[2] } else { '' }
        $max = if ($Rest.Count -ge 4) { $Rest[3] } else { '2000' }
        $out = Invoke-Bridge $excel 'ReadTsv' @($wb, $sh, $rng, $max)
        # Dimension sheets opened from the panel (Dim_*) hold metadata only: member codes such as
        # 1234 are names, not amounts, so they are not hidden.
        if (-not $AllowFigures -and -not ($sh -like 'Dim_*')) { $out = Hide-NumbersInTsv $out }
        Write-Output $out
    }

    'formulas' {
        $excel = Get-Excel
        if (-not (Test-BridgeLoaded $excel)) {
            Stop-WithError "$BridgeName is not loaded in Excel. See references/bridge-setup.md."
        }
        $wb = if ($Rest.Count -ge 1) { $Rest[0] } else { '' }
        $sh = if ($Rest.Count -ge 2) { $Rest[1] } else { '' }
        $rng = if ($Rest.Count -ge 3) { $Rest[2] } else { '' }
        $max = if ($Rest.Count -ge 4) { $Rest[3] } else { '2000' }
        $out = Invoke-Bridge $excel 'ReadFormulas' @($wb, $sh, $rng, $max)
        if (-not $AllowFigures) {
            $out = (($out -split "`n") | ForEach-Object { $c = $_ -split "`t"; if ($c.Count -ge 3) { $c[2] = Hide-Numbers $c[2] }; $c -join "`t" }) -join "`n"
        }
        Write-Output $out
    }

    'hs-audit' {
        $excel = Get-Excel
        if (-not (Test-BridgeLoaded $excel)) {
            Stop-WithError "$BridgeName is not loaded in Excel. See references/bridge-setup.md."
        }
        $wb = if ($Rest.Count -ge 1) { $Rest[0] } else { '' }
        $sh = if ($Rest.Count -ge 2) { $Rest[1] } else { '' }
        $rng = if ($Rest.Count -ge 3) { $Rest[2] } else { '' }
        $max = if ($Rest.Count -ge 4) { $Rest[3] } else { '2000' }
        $raw = Invoke-Bridge $excel 'AuditFunctions' @($wb, $sh, $rng, $max)
        $res = $raw | ConvertFrom-Json
        if (-not $res.ok) { Write-Output $raw; exit 1 }

        $calls = @($res.data.calls)
        if (-not $AllowFigures) {
            foreach ($c in $calls) { $c.text = Hide-Numbers ([string]$c.text) }
        }
        $connections = @{}
        $dimensions = @{}
        $valueCells = 0
        foreach ($c in $calls) {
            $a = @($c.args)
            if ($c.fn -ieq 'HsSetValue') {
                # HsSetValue(value, connection, POV...): drop the value argument.
                if ($a.Count -gt 1) { $a = @($a[1..($a.Count - 1)]) } else { $a = @() }
            }
            elseif ($c.fn -ine 'HsGetValue') {
                continue
            }
            if ($a.Count -eq 0) { continue }
            $valueCells++
            $conn = [string]$a[0]
            if (-not $connections.ContainsKey($conn)) { $connections[$conn] = 0 }
            $connections[$conn]++
            $povParts = @()
            if ($a.Count -gt 1) { $povParts = @($a[1..($a.Count - 1)]) }
            $seen = @{}
            foreach ($part in ($povParts -join ';') -split ';') {
                if ($part -notmatch '#') { continue }
                $dim, $mbr = $part -split '#', 2
                $dim = $dim.Trim()
                if (-not $dimensions.ContainsKey($dim)) { $dimensions[$dim] = @{ cells = 0; members = @{} } }
                if (-not $seen.ContainsKey($dim)) { $dimensions[$dim].cells++; $seen[$dim] = $true }
                $dimensions[$dim].members[$mbr.Trim()] = $true
            }
        }

        $connList = foreach ($k in $connections.Keys) {
            $kind = if ($k -eq '') { 'default' }
                    elseif ($k -ieq 'HsActive') { 'sheet' }
                    elseif ($k -match '^(?i)WSFN\|') { 'shared' }
                    elseif ($k -match '^(?i)URLFN\|') { 'url' }
                    elseif ($k.StartsWith('=')) { 'not-evaluated' }
                    else { 'private' }
            $exists = $null
            if ($kind -eq 'private') {
                try { $exists = ((Invoke-Bridge $excel 'ConnectionExists' @($k)) | ConvertFrom-Json).data } catch { }
            }
            [ordered]@{ connection = $k; kind = $kind; cells = $connections[$k]; existsOnThisMachine = $exists }
        }
        $dimList = foreach ($k in ($dimensions.Keys | Sort-Object)) {
            $m = @($dimensions[$k].members.Keys | Sort-Object)
            [ordered]@{
                dimension      = $k
                cells          = $dimensions[$k].cells
                missingInCells = $valueCells - $dimensions[$k].cells
                distinctMembers = $m.Count
                members        = @($m | Select-Object -First 25)
            }
        }
        $errorCalls = @($calls | Where-Object { ([string]$_.text).StartsWith('#') })
        Write-Json ([ordered]@{
            ok                 = $true
            sheet              = $res.data.sheet
            cellsScanned       = $res.data.cellsScanned
            truncated          = $res.data.truncated
            functions          = @($calls | Group-Object fn | ForEach-Object { [ordered]@{ fn = $_.Name; calls = $_.Count } })
            valueCalls         = $valueCells
            writeCalls         = @($calls | Where-Object { $_.fn -match '^(?i)hsset' } | Select-Object -ExpandProperty cell -First 50)
            brokenAddinLinks   = @($calls | Where-Object { $_.addinPathInFormula } | Select-Object -ExpandProperty cell -First 50)
            connections        = @($connList)
            dimensions         = @($dimList)
            errorCount         = $errorCalls.Count
            errors             = @($errorCalls | Select-Object -First 50 | ForEach-Object {
                                     [ordered]@{ cell = $_.cell; fn = $_.fn; text = $_.text; args = $_.args } })
        })
    }

    'private-connections' {
        # Read-only list of the user's Smart View private connections (HKCU registry). Only the
        # readable fields are parsed; the stored credential is never decoded or printed.
        $key = Join-Path $SvRegistry 'Connections'
        if (-not (Test-Path $key)) { Write-Json ([ordered]@{ ok = $true; data = @() }); break }
        $props = (Get-ItemProperty $key).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' -and $_.Value -is [byte[]] }
        if ($Rest -and $Rest[0]) { $props = $props | Where-Object { $_.Name -eq $Rest[0] } }
        $list = foreach ($p in $props) {
            $text = [Text.Encoding]::Unicode.GetString($p.Value)
            $s = @([regex]::Matches($text, '[ -~]{2,}') | ForEach-Object { $_.Value })
            $i = [array]::FindIndex([string[]]$s, [Predicate[string]] { param($x) $x -match '^https?://' })
            if ($i -lt 0) { continue }
            [ordered]@{
                name        = $p.Name
                url         = $s[$i]
                server      = $s[$i + 1]
                application = $s[$i + 2]
                cube        = $s[$i + 3]
                provider    = $s[$i + 4]
                aliasTable  = $s[$i + 5]
            }
        }
        Write-Json ([ordered]@{ ok = $true; data = @($list); note = 'The user''s connections: read them, never change, rename or remove them.' })
    }

    'start-excel' {
        # Starts Excel the way the user would (excel.exe, visible, add-ins loaded). Never a
        # COM-created instance: it is hidden and Smart View does not load in it.
        $running = $true
        try { $null = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $running = $false }
        if (-not $running) {
            Start-Process excel.exe '/e'
            $excel = $null
            for ($i = 0; $i -lt 60 -and -not $excel; $i++) {
                Start-Sleep -Seconds 2
                try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
            }
            if (-not $excel) { Stop-WithError 'Excel started but did not answer in 2 minutes. Ask the user to check the Excel window (a pop-up may be waiting).' }
        }
        $excel = Get-Excel
        for ($i = 0; $i -lt 15 -and -not (Test-SmartViewLoaded $excel); $i++) { Start-Sleep -Seconds 2 }
        Write-Json ([ordered]@{ ok = $true; started = (-not $running); smartViewLoaded = (Test-SmartViewLoaded $excel); bridgeLoaded = (Test-BridgeLoaded $excel) })
    }

    'open' {
        # Opens a saved workbook in the user's Excel. Read-only unless "edit" is given; never saves.
        if (-not $Rest -or -not $Rest[0]) { Stop-WithError 'Usage: open <path> [edit]' }
        $path = [IO.Path]::GetFullPath($Rest[0])
        if (-not (Test-Path -LiteralPath $path)) { Stop-WithError "File not found: $path" }
        $readOnly = -not ($Rest.Count -ge 2 -and $Rest[1] -ieq 'edit')
        $excel = Get-Excel
        $name = [IO.Path]::GetFileName($path)
        $wb = $null
        foreach ($w in $excel.Workbooks) { if ($w.FullName -ieq $path) { $wb = $w } }
        $already = [bool]$wb
        $events = @()
        if (-not $wb) {
            $watch = Start-DialogWatch ([int](Get-Process EXCEL | Where-Object MainWindowHandle -ne 0 | Select-Object -First 1).Id) @{ Reuse = 'cancel'; DismissMessages = (-not $KeepMessages.IsPresent) }
            # UpdateLinks 0: do not refresh external links on open.
            try { $wb = Invoke-WithRetry { $excel.Workbooks.Open($path, 0, $readOnly) } } finally { $events = @(Stop-DialogWatch $watch) }
            Add-Opened $wb.Name
        }
        Write-Json ([ordered]@{ ok = $true; workbook = $wb.Name; alreadyOpen = $already; readOnly = $wb.ReadOnly; sheets = @($wb.Worksheets | ForEach-Object { $_.Name }); dialogs = @($events) })
    }

    'new-workbook' {
        # A new, empty workbook for the user (not the scratch one). Never saved by the skill.
        $excel = Get-Excel
        $wb = Invoke-WithRetry { $excel.Workbooks.Add() }
        Add-Opened $wb.Name
        Write-Json ([ordered]@{ ok = $true; workbook = $wb.Name; note = 'Unsaved. The user saves it where they want.' })
    }

    'close' {
        # Closes WITHOUT saving a workbook the skill opened or created (open, new-workbook,
        # panel.ps1 open-report). Any other workbook is the user's: refused.
        if (-not $Rest -or -not $Rest[0]) { Stop-WithError 'Usage: close <workbook>' }
        $name = $Rest[0]
        $mine = @(if (Test-Path $OpenedList) { Get-Content -LiteralPath $OpenedList -Encoding UTF8 })
        $xpid = (Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -ne 0 | Select-Object -First 1).Id
        if ($mine -notcontains "$xpid|$name") { Stop-WithError "Not closed: $name was not opened or created by the skill. The user closes their own workbooks." }
        $excel = Get-Excel
        try { $wb = $excel.Workbooks.Item($name) } catch { Stop-WithError "Workbook $name is not open." }
        Invoke-WithRetry { $wb.Close($false) } | Out-Null
        Set-Content -LiteralPath $OpenedList -Value @($mine | Where-Object { $_ -ne "$xpid|$name" }) -Encoding UTF8
        Write-Json ([ordered]@{ ok = $true; closed = $name; saved = $false })
    }

    'drop-sheet' {
        # Deletes working sheets of the SCRATCH workbook only.
        if (-not $Rest -or $Rest.Count -lt 1) { Stop-WithError 'Usage: drop-sheet <sheet> [sheet...]' }
        $excel = Get-Excel
        $scratch = ((Invoke-Bridge $excel 'Workbooks' @()) | ConvertFrom-Json).data | Where-Object { $_.scratch } | Select-Object -First 1
        if (-not $scratch) { Stop-WithError 'No scratch workbook open.' }
        $wb = $excel.Workbooks.Item($scratch.name)
        $done = @()
        $excel.DisplayAlerts = $false
        try { foreach ($n in $Rest) { try { Invoke-WithRetry { $wb.Worksheets.Item($n).Delete() } | Out-Null; $done += $n } catch { } } }
        finally { $excel.DisplayAlerts = $true }
        Write-Json ([ordered]@{ ok = $true; deleted = $done })
    }

    'install' {
        $src = Join-Path (Split-Path $PSScriptRoot) "bridge\$BridgeName"
        if (-not (Test-Path $src)) {
            Stop-WithError "Prebuilt add-in not found at $src. A maintainer builds it once: see references/bridge-setup.md."
        }
        $addinsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
        $dest = Join-Path $addinsDir $BridgeName
        $excel = $null
        try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
        if ($excel) {
            foreach ($ai in @($excel.AddIns)) {
                if ($ai.Name -ieq $BridgeName -and $ai.Installed) { $ai.Installed = $false }
            }
        }
        if (-not (Test-Path $addinsDir)) { $null = New-Item -ItemType Directory -Path $addinsDir }
        Copy-Item -LiteralPath $src -Destination $dest -Force
        Unblock-File -LiteralPath $dest
        $loaded = $false
        if ($excel) {
            # AddIns.Add fails when no workbook is open: open a blank one for the call.
            $temp = $null
            if ($excel.Workbooks.Count -eq 0) { $temp = $excel.Workbooks.Add() }
            try {
                $ai = $excel.AddIns.Add($dest, $false)
                $ai.Installed = $true
                $loaded = Test-BridgeLoaded $excel
            }
            finally {
                if ($temp) { $temp.Close($false) }
            }
        }
        # Figures switch: created at 0 if missing, never overwritten (the user decides).
        $figVar = 'SMARTVIEW_SKILL_ALLOW_FIGURES'
        $figNow = [Environment]::GetEnvironmentVariable($figVar, 'User')
        if ($null -eq $figNow) { [Environment]::SetEnvironmentVariable($figVar, '0', 'User'); $figNow = '0' }
        # Local environment profiles (which cube, GetValues connection...): an empty list if
        # missing, never overwritten. Lives outside the skill, so it is never published.
        $profilesFile = Join-Path $DataDir 'connections.json'
        if (-not (Test-Path $DataDir)) { $null = New-Item -ItemType Directory -Path $DataDir }
        if (-not (Test-Path $profilesFile)) { [IO.File]::WriteAllText($profilesFile, "[]`r`n", [Text.UTF8Encoding]::new($false)) }
        Write-Json ([ordered]@{
            ok        = $true
            profiles  = $profilesFile
            figures   = if ($figNow -eq '1') { 'visible on explicit request' } else { "hidden. To let Claude see amounts, set the user environment variable $figVar to 1 yourself (Windows: Edit environment variables for your account)." }
            installed = $dest
            loaded    = $loaded
            next      = if ($loaded) { 'Run: sv.ps1 status' } else { 'Open Excel and run sv.ps1 install again to load it.' }
        })
    }

    'uninstall' {
        $dest = Join-Path $env:APPDATA "Microsoft\AddIns\$BridgeName"
        $excel = $null
        try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
        if ($excel) {
            foreach ($ai in @($excel.AddIns)) {
                if ($ai.Name -ieq $BridgeName -and $ai.Installed) { $ai.Installed = $false }
            }
        }
        $removed = $false
        if (Test-Path $dest) { Remove-Item -LiteralPath $dest -Force; $removed = $true }
        Write-Json ([ordered]@{ ok = $true; removed = $removed; path = $dest })
    }

    default {
        Stop-WithError "Unknown command '$Command'. Run: sv.ps1 help"
    }
}

#requires -Version 5.1
<#
.SYNOPSIS
  Reads and drives the Smart View panel tree in the user's Excel (no Smart View API exists for it).

.DESCRIPTION
  Uses standard Windows tree-view messages on the panel's SysTreeView32. The panel must be
  open and already connected (the user picked the server in Private Connections). Labels
  are matched in English and Spanish only.

  Commands
    ensure                           Opens the Smart View panel if it is closed.
    server <url>                     Points the panel at a server: types the URL in the Private
                                     Connections box, presses Enter, waits for the tree. A sign-in
                                     window, if any, is for the user.
    tree [depth]                     Text of the panel tree, expanded levels only.
    cubes <app> [url]                Cubes (plan types) of an application: the children of the app
                                     node that are not folders (Dimensions, Reports, Library...).
    open-dimension <app> <dimension> [url] [sheet]
                                     Expands <app> > Dimensions and opens <dimension> (EPM Admin
                                     Extension view) on a BLANK sheet of the scratch workbook.
                                     With [url], runs ensure and server first.

    open-report <app> <folder/.../report> [url]
                                     Opens a report from the panel (for example
                                     "Reports/Finance/P&L"), reads its prompts, accepts the
                                     default values and waits for the report workbook.

  The dimension sheet edits metadata if submitted: read it, never write into it, never submit.
  Context-menu actions (Refresh Database, rules...) are not offered here: they change the server.
#>
param(
    [Parameter(Position = 0)] [string] $Command = 'tree',
    [Parameter(Position = 1, ValueFromRemainingArguments = $true)] [string[]] $Rest
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -Path (Join-Path $PSScriptRoot 'TreeRemote.cs')
. (Join-Path $PSScriptRoot 'DialogWatch.ps1')
$sv = Join-Path $PSScriptRoot 'sv.ps1'

$Labels = @{ Dimensions = @('Dimensions', 'Dimensiones') }

function Out-Result($o) { $o | ConvertTo-Json -Depth 5 }

Add-Type @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class PanelBox {
  delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(P p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, P p, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, StringBuilder l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
  [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  static string Cls(IntPtr h) { var s = new StringBuilder(64); GetClassName(h, s, 64); return s.ToString(); }
  public static string Text(IntPtr h) { var t = new StringBuilder(1024); SendMessage(h, 0x000D, (IntPtr)1024, t); return t.ToString(); }
  static List<IntPtr> Kids(IntPtr dlg) {
    var r = new List<IntPtr>();
    EnumChildWindows(dlg, (c, l) => { if (GetParent(c) == dlg && IsWindowVisible(c)) r.Add(c); return true; }, IntPtr.Zero);
    return r; }
  // The URL box of Private Connections: a visible Edit inside a ComboBox whose dialog also holds
  // the panel tree (the box is empty until a server is chosen).
  public static IntPtr UrlBox(uint pid) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p != pid) return true;
      EnumChildWindows(h, (c, l2) => { if (IsWindowVisible(c) && Cls(c) == "Edit" && Cls(GetParent(c)) == "ComboBox") {
          bool tree = false; foreach (var k in Kids(GetParent(GetParent(c)))) if (Cls(k) == "SysTreeView32") tree = true;
          if (tree || Text(c).StartsWith("http")) { found = c; return false; } } return true; }, IntPtr.Zero);
      return found == IntPtr.Zero; }, IntPtr.Zero);
    return found; }
  // The panel's Home page: a dialog with two big buttons on top (Shared, then Private
  // Connections), the recent list (a nested dialog) and a "display on startup" button.
  // Clicks the second top button. Language-independent: matched by layout, not by text.
  public static bool LeaveHome(uint pid) {
    IntPtr btn = IntPtr.Zero;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p != pid) return true;
      EnumChildWindows(h, (c, l2) => { if (Cls(c) != "#32770" || !IsWindowVisible(c)) return true;
          var k = Kids(c); var buttons = new List<IntPtr>(); bool nested = false, tree = false;
          foreach (var x in k) { string cl = Cls(x); if (cl == "Button") buttons.Add(x); if (cl == "#32770") nested = true; if (cl == "SysTreeView32") tree = true; }
          if (!tree && nested && buttons.Count == 3) { btn = buttons[1]; return false; } return true; }, IntPtr.Zero);
      return btn == IntPtr.Zero; }, IntPtr.Zero);
    if (btn == IntPtr.Zero) return false;
    SendMessage(btn, 0x00F5, IntPtr.Zero, (string)null);
    return true; }
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  static string Title(IntPtr h) { var s = new StringBuilder(256); GetWindowText(h, s, 256); return s.ToString(); }
  // The report prompt dialog (WinForms, "Select Prompts"): a top-level window of Excel that
  // holds a WinForms list box. Matched by structure so the title language does not matter.
  public static IntPtr PromptDialog(uint pid) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p != pid || !IsWindowVisible(h)) return true;
      bool lb = false;
      EnumChildWindows(h, (c, l2) => { if (Cls(c).StartsWith("WindowsForms10.LISTBOX")) { lb = true; return false; } return true; }, IntPtr.Zero);
      if (lb) { found = h; return false; } return true; }, IntPtr.Zero);
    return found; }
  // Prompt labels and their current values, in window order.
  public static List<string> ReadPrompts(IntPtr dlg) {
    var r = new List<string>();
    r.Add("title=" + Title(dlg));
    EnumChildWindows(dlg, (c, l) => { string cl = Cls(c);
      if (cl.StartsWith("WindowsForms10.STATIC") && Text(c).Trim().Length > 0) r.Add("prompt=" + Text(c).Trim());
      if (cl.StartsWith("WindowsForms10.LISTBOX")) {
        int n = (int)SendMessage(c, 0x018B, IntPtr.Zero, (string)null); var vals = new List<string>();
        for (int i = 0; i < n && i < 50; i++) { var sb = new StringBuilder(512); SendMessage(c, 0x0189, (IntPtr)i, sb); vals.Add(sb.ToString()); }
        r.Add("value=" + string.Join(", ", vals)); }
      return true; }, IntPtr.Zero);
    return r; }
  // Presses the dialog's OK button (BM_CLICK); false when there is none.
  public static bool PressOk(IntPtr dlg) {
    IntPtr ok = IntPtr.Zero;
    EnumChildWindows(dlg, (c, l) => { string t = Text(c).Replace("&", "").Trim();
      if (Cls(c).Contains("BUTTON") && (t == "OK" || t == "Aceptar")) { ok = c; return false; } return true; }, IntPtr.Zero);
    if (ok == IntPtr.Zero) return false;
    SendMessage(ok, 0x00F5, IntPtr.Zero, (string)null);
    return true; }
  public static void Go(IntPtr box, string url) {
    SendMessage(box, 0x000C, IntPtr.Zero, url);
    PostMessage(box, 0x100, (IntPtr)0x0D, (IntPtr)0x001C0001);
    PostMessage(box, 0x101, (IntPtr)0x0D, unchecked((IntPtr)(long)0xC01C0001)); }
}
"@

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object MainWindowHandle -ne 0 | Select-Object -First 1
if (-not $excel) { Out-Result ([ordered]@{ ok = $false; error = 'Excel is not running.' }); exit 1 }

function Ensure-Panel {
    $t = [TreeRemote]::FindPanelTree([uint32]$excel.Id)
    if ($t -ne [IntPtr]::Zero -or [PanelBox]::UrlBox([uint32]$excel.Id) -ne [IntPtr]::Zero) { return 'already open' }
    $null = & $sv call ExecuteMenu '' '' 'Panel'
    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Seconds 1
        if ([TreeRemote]::FindPanelTree([uint32]$excel.Id) -ne [IntPtr]::Zero -or [PanelBox]::UrlBox([uint32]$excel.Id) -ne [IntPtr]::Zero -or [PanelBox]::LeaveHome([uint32]$excel.Id)) { break }
    }
    return 'opened'
}

. (Join-Path $PSScriptRoot 'SvUrl.ps1')

function Set-Server([string] $url) {
    $url = ConvertTo-SvUrl $url
    # A panel that has just opened takes a while to draw Home or Private Connections: poll.
    $box = [IntPtr]::Zero
    for ($i = 0; $i -lt 30 -and $box -eq [IntPtr]::Zero; $i++) {
        $box = [PanelBox]::UrlBox([uint32]$excel.Id)
        if ($box -eq [IntPtr]::Zero) {
            if ([PanelBox]::LeaveHome([uint32]$excel.Id)) { Start-Sleep -Seconds 2 } else { Start-Sleep -Seconds 1 }
        }
    }
    if ($box -eq [IntPtr]::Zero) { return @{ ok = $false; error = 'Private Connections URL box not found. Ask the user to click Private Connections in the Smart View panel once.' } }
    $hostName = ([uri]$url).Host
    $watch = Start-DialogWatch ([int]$excel.Id) @{ Reuse = 'cancel'; DismissMessages = $true }
    [PanelBox]::Go($box, $url)
    $ok = $false
    for ($i = 0; $i -lt 60 -and -not $ok; $i++) {
        Start-Sleep -Seconds 2
        $t = [TreeRemote]::FindPanelTree([uint32]$excel.Id)
        if ($t -ne [IntPtr]::Zero) { $ok = (@([TreeRemote]::Dump($t, 0)) -join ' ') -match [regex]::Escape($hostName) }
        # A connection error message means it will not load: stop waiting.
        # Smart View repeats the message a few times: keep watching briefly so all are closed.
        if (@($watch.sync.events | Where-Object { $_.text -match '(?i)connect|provider|resolv|conectar|proveedor' }).Count) { Start-Sleep -Seconds 8; break }
    }
    # The server node can come back collapsed with no children loaded: open it (Enter) and
    # wait for the applications.
    if ($ok -and @([TreeRemote]::Dump($t, 1)).Count -lt 2) {
        $root = [TreeRemote]::Find($t, [string[][]]@(, [string[]]@(@([TreeRemote]::Dump($t, 0))[0])))
        [TreeRemote]::Open($t, $root)
        for ($i = 0; $i -lt 30 -and @([TreeRemote]::Dump($t, 1)).Count -lt 2; $i++) { Start-Sleep -Seconds 2 }
        $ok = @([TreeRemote]::Dump($t, 1)).Count -ge 2
    }
    $events = @(Stop-DialogWatch $watch)
    return @{ ok = $ok; server = $hostName; dialogs = @($events) }
}

if ($Command -ieq 'ensure') { Out-Result ([ordered]@{ ok = $true; panel = (Ensure-Panel) }); exit 0 }
if ($Command -ieq 'server') {
    if ($Rest.Count -lt 1) { Out-Result ([ordered]@{ ok = $false; error = 'Usage: server <url>' }); exit 1 }
    $null = Ensure-Panel
    Out-Result (Set-Server $Rest[0]); exit 0
}
if ($Command -ieq 'cubes') {
    if ($Rest.Count -lt 1) { Out-Result ([ordered]@{ ok = $false; error = 'Usage: cubes <app> [url]' }); exit 1 }
    $null = Ensure-Panel
    if ($Rest.Count -ge 2) { $r = Set-Server $Rest[1]; if (-not $r.ok) { Out-Result $r; exit 1 } }
    $t = [TreeRemote]::FindPanelTree([uint32]$excel.Id)
    $node = [TreeRemote]::Find($t, [string[][]]@(, [string[]]@($Rest[0])))
    if ($node -eq [IntPtr]::Zero) { Out-Result ([ordered]@{ ok = $false; error = "Application $($Rest[0]) not found in the panel."; tree = @([TreeRemote]::Dump($t, 2)) }); exit 1 }
    $folders = '^(Dimensions|Dimensiones|Attribute Dimensions|Dimensiones de atributo|Reports|Informes|Documents|Documentos|Library|Biblioteca|Task Lists|Listas de tareas|Smart Lists|Rules|Reglas)$'
    $all = @([TreeRemote]::Dump($t, 2))
    $i = [array]::IndexOf(($all | ForEach-Object { $_.Trim() }), $Rest[0])
    $cubes = @()
    for ($k = $i + 1; $k -lt $all.Count; $k++) {
        $indent = $all[$k].Length - $all[$k].TrimStart().Length
        if ($indent -le ($all[$i].Length - $all[$i].TrimStart().Length)) { break }
        $name = $all[$k].Trim()
        # The panel shows "<cube> - <private connection>" when a private connection exists.
        if ($name -notmatch $folders) { $cubes += [ordered]@{ cube = ($name -split ' - ', 2)[0]; panelLabel = $name } }
    }
    Out-Result ([ordered]@{ ok = $true; app = $Rest[0]; cubes = @($cubes); note = 'Which one is the input (BSO) cube and which the reporting (ASO) one is not shown here: the dimension sheet lists one consolidation column per cube, and naming often marks reporting cubes (e.g. *RPT). Ask the user when unsure.' })
    exit 0
}
if ($Command -ieq 'open-report' -and $Rest.Count -ge 3) {
    $null = Ensure-Panel
    $r = Set-Server $Rest[2]
    if (-not $r.ok) { Out-Result $r; exit 1 }
}
if ($Command -ieq 'open-dimension' -and $Rest.Count -ge 3) {
    $null = Ensure-Panel
    $r = Set-Server $Rest[2]
    if (-not $r.ok) { Out-Result $r; exit 1 }
}

$tree = [TreeRemote]::FindPanelTree([uint32]$excel.Id)
if ($tree -eq [IntPtr]::Zero) {
    Out-Result ([ordered]@{ ok = $false; error = 'Smart View panel tree not found. Ask the user to open the Smart View panel and connect to the server in Private Connections.' })
    exit 1
}

switch ($Command.ToLowerInvariant()) {
    'tree' {
        $depth = if ($Rest.Count -ge 1) { [int]$Rest[0] } else { 3 }
        Out-Result ([ordered]@{ ok = $true; tree = @([TreeRemote]::Dump($tree, $depth)) })
    }
    'open-dimension' {
        if ($Rest.Count -lt 2) { Out-Result ([ordered]@{ ok = $false; error = 'Usage: open-dimension <app> <dimension>' }); exit 1 }
        $app, $dim = $Rest[0], $Rest[1]
        # A sheet with content makes Smart View ask what to do with it (modal): use a blank one.
        $sheet = if ($Rest.Count -ge 4 -and $Rest[3]) { $Rest[3] } else { "Dim_$($dim -replace '[^A-Za-z0-9_]', '_')" }
        if ($sheet.Length -gt 31) { $sheet = $sheet.Substring(0, 31) }
        $null = & $sv call Scratch $sheet true
        $path = [string[][]]@(, [string[]]@($app))
        $path += , [string[]]$Labels.Dimensions
        $path += , [string[]]@($dim)
        $node = [TreeRemote]::Find($tree, $path)
        if ($node -eq [IntPtr]::Zero) {
            Out-Result ([ordered]@{ ok = $false; error = "Node $app > Dimensions > $dim not found in the panel tree."; tree = @([TreeRemote]::Dump($tree, 3)) })
            exit 1
        }
        $watch = Start-DialogWatch ([int]$excel.Id) @{ Reuse = 'cancel'; DismissMessages = $true }
        [TreeRemote]::Open($tree, $node)
        $range = 'A1'
        for ($i = 0; $i -lt 60 -and $range -eq 'A1'; $i++) {
            Start-Sleep -Seconds 2
            try { $range = ((& $sv read '' $sheet '' 3) -split "`n")[0].Split("`t")[1] } catch { }
        }
        $events = @(Stop-DialogWatch $watch)
        Out-Result ([ordered]@{
            ok    = ($range -ne 'A1')
            dialogs = @($events)
            sheet = $sheet
            range = $range
            next  = "Read it: sv.ps1 read '' $sheet ; expand with: sv.ps1 call ZoomIn '' $sheet A2 1. Never submit this sheet."
        })
    }
    'open-report' {
        if ($Rest.Count -lt 2) { Out-Result ([ordered]@{ ok = $false; error = 'Usage: open-report <app> <folder/.../report> [url]' }); exit 1 }
        $path = [string[][]]@(, [string[]]@($Rest[0]))
        foreach ($p in ($Rest[1] -split '/')) { if ($p) { $path += , [string[]]@($p) } }
        $node = [TreeRemote]::Find($tree, $path)
        if ($node -eq [IntPtr]::Zero) {
            Out-Result ([ordered]@{ ok = $false; error = "Report $($Rest[1]) not found under $($Rest[0]). Folders load when opened: list them with 'tree' after opening the parent."; tree = @([TreeRemote]::Dump($tree, 4)) })
            exit 1
        }
        $excelCom = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $before = @($excelCom.Workbooks | ForEach-Object { $_.Name })
        $watch = Start-DialogWatch ([int]$excel.Id) @{ Reuse = 'cancel'; DismissMessages = $true }
        [TreeRemote]::Open($tree, $node)
        # Prompts: read them and accept the defaults. Choosing other values needs Smart View's
        # member selector, which is the user's.
        $prompts = @()
        $book = $null
        for ($i = 0; $i -lt 150 -and -not $book; $i++) {
            Start-Sleep -Seconds 2
            $dlg = [PanelBox]::PromptDialog([uint32]$excel.Id)
            if ($dlg -ne [IntPtr]::Zero -and -not $prompts) {
                $prompts = @([PanelBox]::ReadPrompts($dlg))
                $null = [PanelBox]::PressOk($dlg)
            }
            try { $book = @($excelCom.Workbooks | ForEach-Object { $_.Name }) | Where-Object { $before -notcontains $_ } | Select-Object -First 1 } catch { }
        }
        $events = @(Stop-DialogWatch $watch)
        $range = $null
        # Record it so sv.ps1 close can close it without saving (it is the skill's, not the user's).
        if ($book) { Add-Content -LiteralPath (Join-Path $env:LOCALAPPDATA 'smart-view-skill\opened-by-skill.txt') -Value "$($excel.Id)|$book" -Encoding UTF8 }
        if ($book) { try { $range = $excelCom.Workbooks.Item($book).ActiveSheet.UsedRange.Address($false, $false) } catch { } }
        Out-Result ([ordered]@{
            ok       = [bool]$book
            workbook = $book
            range    = $range
            prompts  = $prompts
            dialogs  = @($events)
            note     = 'The report opens as a new, unconnected workbook with its figures. Read its structure with sv.ps1 read (figures stay hidden). Prompts were accepted with their default values.'
        })
    }
    default { Out-Result ([ordered]@{ ok = $false; error = "Unknown command '$Command'." }); exit 1 }
}

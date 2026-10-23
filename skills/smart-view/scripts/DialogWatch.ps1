# Dialog watcher for Smart View / Excel modal windows.
# Runs in a background runspace while a bridge call blocks Excel, reads every new dialog of
# the Excel process and answers the known ones according to a policy:
#   - sign-in windows: never touched (the user signs in);
#   - submit / rule / calculation confirmations: never confirmed;
#   - "worksheet contains data from prior operation": answered with $Policy.Reuse
#     (contents | contentsandpov | clear | cancel);
#   - messages (one OK / Aceptar / Close button, optionally with "Details >>"): text recorded, then dismissed if
#     $Policy.DismissMessages, so the call returns and the agent can report the message;
#   - anything else: recorded, not touched.
# Labels are matched in English and Spanish.

function Start-DialogWatch([int] $ExcelPid, [hashtable] $Policy) {
    $sync = [hashtable]::Synchronized(@{ stop = $false; events = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList)) })
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript({
        param($ExcelPid, $Policy, $sync)
        Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
        Add-Type @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class DW {
  public delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(P p, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string Cls(IntPtr h) { var sb = new StringBuilder(64); GetClassName(h, sb, 64); return sb.ToString(); }
  public static List<IntPtr> Dialogs(uint pid) {
    var res = new List<IntPtr>();
    EnumWindows((h,l) => { uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h) && GetWindow(h, 4) != IntPtr.Zero) res.Add(h); // owned = modal-like
      return true; }, IntPtr.Zero);
    return res; }
  public static string Title(IntPtr h) { var sb = new StringBuilder(256); GetWindowText(h, sb, 256); return sb.ToString(); }
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
  [StructLayout(LayoutKind.Sequential)] struct FLASHWINFO { public uint cbSize; public IntPtr hwnd; public uint dwFlags; public uint uCount; public uint dwTimeout; }
  [DllImport("user32.dll")] static extern bool FlashWindowEx(ref FLASHWINFO f);
  // Puts a window the user must act on in front. Windows blocks SetForegroundWindow from a
  // background process unless a key event just happened, so a bare Alt press/release goes
  // first (no character is typed). If it still is not in front, the taskbar button flashes.
  public static bool Front(IntPtr h) {
    IntPtr owner = GetWindow(h, 4);
    if (owner != IntPtr.Zero && IsIconic(owner)) ShowWindow(owner, 9);   // SW_RESTORE
    if (IsIconic(h)) ShowWindow(h, 9);
    keybd_event(0x12, 0, 0, UIntPtr.Zero); keybd_event(0x12, 0, 2, UIntPtr.Zero);   // Alt down, up
    SetForegroundWindow(h); BringWindowToTop(h);
    bool ok = GetForegroundWindow() == h;
    if (!ok) { var f = new FLASHWINFO(); f.cbSize = (uint)Marshal.SizeOf(f); f.hwnd = h; f.dwFlags = 3 | 12; f.uCount = 0; f.dwTimeout = 0; FlashWindowEx(ref f); }
    return ok; }
}
"@
        $A = [System.Windows.Automation.AutomationElement]
        $seen = @{}
        function Has($text, [string[]] $words) { foreach ($w in $words) { if ($text -match [regex]::Escape($w)) { return $true } }; return $false }
        while (-not $sync.stop) {
            foreach ($h in [DW]::Dialogs([uint32]$ExcelPid)) {
                $key = $h.ToInt64()
                if ($seen.ContainsKey($key)) { continue }
                # Only real dialogs: standard dialog class or a Smart View / Excel titled window.
                # The floating POV toolbar and task panes are owned windows too: skip them.
                $cls = [DW]::Cls($h); $ttl = [DW]::Title($h)
                $isDialog = ($cls -eq '#32770') -or ($ttl -match '(?i)smart view|oracle|microsoft excel')
                if (-not $isDialog -or $ttl -match '^(?i)POV') { $seen[$key] = $true; continue }
                $seen[$key] = $true
                try {
                    $el = $A::FromHandle($h)
                    $all = $el.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
                    $texts = @(); $buttons = @(); $hasEdit = $false
                    foreach ($e in $all) {
                        $ct = $e.Current.ControlType.ProgrammaticName; $nm = $e.Current.Name
                        if ($ct -match 'Edit|Document') { $hasEdit = $true }
                        $hw = [IntPtr]$e.Current.NativeWindowHandle
                        $isBtn = ($ct -match 'Button|Hyperlink') -or ($hw -ne [IntPtr]::Zero -and [DW]::Cls($hw) -eq 'Button')
                        if ($isBtn) { if ($nm) { $buttons += $e } }
                        elseif ($nm) { $texts += $nm }
                    }
                    $title = [DW]::Title($h)
                    $text = (($texts | Select-Object -Unique) -join ' ').Trim()
                    $all2 = "$title $text $(($buttons | ForEach-Object { $_.Current.Name }) -join ' ')".ToLowerInvariant()
                    $names = @($buttons | ForEach-Object { $_.Current.Name })
                    $action = 'left for the user'
                    $pick = $null
                    if ($hasEdit -or (Has $all2 @('sign in','log in','login','password','contraseña','usuario','identity domain','iniciar sesión'))) {
                        $action = 'sign-in window: waiting for the user'
                    }
                    elseif (Has $all2 @('submit','enviar','rule','regla','calculat','calcul','refresh database','actualizar base de datos')) {
                        $action = 'confirmation that changes the server: not touched'
                    }
                    elseif (Has $all2 @('prior operation','operación anterior','contents of the sheet','contenido de la hoja')) {
                        $want = switch ($Policy.Reuse) {
                            'contents'       { { param($n) ($n -match '(?i)reuse|reutiliz') -and ($n -match '(?i)only|solo|sólo') } }
                            'contentsandpov' { { param($n) ($n -match '(?i)reuse|reutiliz') -and ($n -match 'POV') -and -not ($n -match '(?i)only|solo|sólo') } }
                            'clear'          { { param($n) $n -match '(?i)clear|borrar|limpiar' } }
                            default          { { param($n) $n -match '(?i)^(cancel|cancelar)$' } }
                        }
                        $pick = $buttons | Where-Object { & $want $_.Current.Name } | Select-Object -First 1
                        $action = if ($pick) { "reuse dialog: chose '$($pick.Current.Name)'" } else { 'reuse dialog: option not found, left for the user' }
                    }
                    elseif ($Policy.DismissMessages) {
                        # A message: one OK-like button, optionally with "Details >>" (Smart View
                        # error messages carry both).
                        $okBtn = @($buttons | Where-Object { $_.Current.Name -replace '&', '' -match '(?i)^(ok|aceptar|accept|close|cerrar)$' })
                        $other = @($buttons | Where-Object { $okBtn -notcontains $_ -and ($_.Current.Name -replace '&', '' -notmatch '(?i)^(details|detalles)\s*(>>|<<)?$') })
                        # UI Automation can list the same button twice: compare by name.
                        $okNames = @($okBtn | ForEach-Object { $_.Current.Name } | Select-Object -Unique)
                        if ($okNames.Count -eq 1 -and $other.Count -eq 0) {
                            $pick = $okBtn[0]
                            $action = 'message dismissed after reading it'
                        }
                    }
                    # Windows the user must answer (sign-in, anything left untouched): bring them
                    # to the front, or they open behind the user's other windows.
                    if (-not $pick -and $action -match 'sign-in|left for the user|not touched') {
                        $front = [DW]::Front($h)
                        $action += $(if ($front) { ' (brought to front)' } else { ' (flashing in the taskbar)' })
                    }
                    if ($pick) {
                        $ph = [IntPtr]$pick.Current.NativeWindowHandle
                        if ($ph -ne [IntPtr]::Zero -and [DW]::Cls($ph) -eq 'Button') { $null = [DW]::PostMessage($ph, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) }  # BM_CLICK
                        else {
                            try { $pick.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
                            catch { $action += ' (click failed: ' + $_.Exception.Message + ')' }
                        }
                    }
                    # Transient windows with no text and no options are noise: not reported.
                    if ($text -or $names.Count -gt 0) {
                        $null = $sync.events.Add([ordered]@{ title = $title; text = $text; options = $names; action = $action })
                    }
                }
                catch {
                    $null = $sync.events.Add([ordered]@{ title = [DW]::Title($h); text = ''; options = @(); action = 'could not read: ' + $_.Exception.Message })
                }
            }
            Start-Sleep -Milliseconds 500
        }
    }).AddArgument($ExcelPid).AddArgument($Policy).AddArgument($sync)
    $handle = $ps.BeginInvoke()
    return @{ ps = $ps; rs = $rs; handle = $handle; sync = $sync }
}

function Stop-DialogWatch($w) {
    if (-not $w) { return @() }
    Start-Sleep -Milliseconds 700   # let it record a dialog that closed the call
    $w.sync.stop = $true
    try { $null = $w.ps.EndInvoke($w.handle) } catch { }
    $w.ps.Dispose(); $w.rs.Close()
    return @($w.sync.events)
}

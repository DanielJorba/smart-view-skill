#requires -Version 5.1
<#
.SYNOPSIS
  Maintainer tool: builds bridge\SmartViewBridge.xlam from bridge\SmartViewBridge.bas.

.DESCRIPTION
  End users never run this: they run "sv.ps1 install", which copies the prebuilt add-in.
  A maintainer runs it after changing SmartViewBridge.bas, then commits the .xlam.

  Uses a separate, hidden Excel instance (the user's Excel is not touched) and needs
  Excel's "Trust access to the VBA project object model". This script never changes
  that setting. If it is off, it prints the manual steps instead (2 minutes).
#>
[CmdletBinding()]
param(
    [string] $OutputPath = (Join-Path (Split-Path $PSScriptRoot) 'bridge\SmartViewBridge.xlam')
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$bridgeBas = Join-Path (Split-Path $PSScriptRoot) 'bridge\SmartViewBridge.bas'
if (-not (Test-Path $bridgeBas)) { throw "Bridge module not found at $bridgeBas." }

function Show-ManualSteps {
    @"
Automatic build needs Excel's "Trust access to the VBA project object model", which is
off. That is a security setting; leave it as it is and build by hand (2 minutes):

  1. In Excel, create a new blank workbook.
  2. Press Alt+F11. In the VBA editor: File > Import File... and pick
       $bridgeBas
  3. Debug > Compile VBAProject. It should finish without messages.
  4. Back in Excel: File > Save As > type "Excel Add-in (*.xlam)", name it
       SmartViewBridge.xlam
     and save it in this folder (not the AddIns folder Excel proposes):
       $(Split-Path $OutputPath)
  5. Close the blank workbook without saving.

Then: sv.ps1 install   (loads it in Excel; same command for every other user)
"@
}

$vbom = $null
try { $vbom = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -ErrorAction Stop).AccessVBOM } catch { }
if ($vbom -ne 1) {
    Show-ManualSteps
    exit 2
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $null = $wb.VBProject.VBComponents.Import($bridgeBas)
    $xlOpenXMLAddIn = 55
    $wb.SaveAs($OutputPath, $xlOpenXMLAddIn)
    [ordered]@{ ok = $true; output = $OutputPath; next = 'Commit the .xlam, then: sv.ps1 install' } | ConvertTo-Json
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

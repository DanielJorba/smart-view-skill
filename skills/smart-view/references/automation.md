# Automating Smart View beyond the API

The VBA API covers connections, ad hoc operations, POV, forms and member queries.
Some things are UI-only and need other techniques. Use them as a last resort, and
first ask whether the job belongs on the server instead.

## First question: should this be Smart View at all?

Smart View automation needs an interactive Windows session with Excel open and a
signed-in user. For **unattended or scheduled** extraction or loading, server-side tools
are the right layer: EPM Automate, the EPM REST APIs (for example data slice export,
job status), Data Integration exports, or Essbase MaxL. Smart View automation fits when
the output is an Excel workbook a person uses, or the user's own session is required.

## Driving Smart View from outside Excel (PowerShell, Python)

- The `Hyp*` functions run only inside Excel's process. From outside, call a VBA macro
  that wraps them (`Application.Run`), as `scripts/sv.ps1` does with the bridge.
- **Attach to the user's Excel** (`GetActiveObject("Excel.Application")`) when the job
  needs their Smart View session. An Excel started by automation (`New-Object
  -ComObject Excel.Application`) does **not** load COM add-ins such as Smart View,
  and has no signed-in session. Smart View's COM add-in ProgID is
  `Hyperion.CommonAddin`; it can be connected with
  `Application.COMAddIns("Hyperion.CommonAddin").Connect = True`, but the user still has
  to sign in.
- Oracle documents a registry value that stops Smart View loading in Excel started by
  automation: DWORD `DisableInAutomation = 1` under
  `HKCU\Software\Oracle\SmartView\Options`. If it is set, Smart View calls cannot work
  in such an instance. Do not change it without the user.
- Keep `Application.EnableEvents = True` whenever Smart View functions run: Oracle
  states the API depends on Excel events.
- Excel rejects COM calls while a cell is being edited or a dialog is open
  (`RPC_E_CALL_REJECTED` 0x80010001, `RPC_E_SERVERCALL_RETRYLATER` 0x8001010A): retry
  with a short sleep.
- **Hidden Excel hangs on VBA errors.** If a macro fails in an invisible instance, the
  "Microsoft Visual Basic" error dialog is invisible too and the script waits forever.
  Handle errors in the macro (write the error to a cell or file and exit). If it
  happens, read the dialog text with `EnumChildWindows` (user32) and kill only that
  instance by PID (its `MainWindowHandle` is 0), never the user's Excel.
- *(Field report)* Opening workbooks with Smart View external links through COM has
  hung Excel on some machines. Prefer building a fresh workbook.

## Formatted reports (EPM Cloud Reports): changing the POV by macro

There is no API to change the POV of formatted reports inserted in Excel. What has
worked in practice *(field report, Smart View 2x, EPM Cloud)*:

- `HypExecuteMenu(sheet, "Edit POV")` opens a **modal dialog that blocks VBA**. Handle
  it from a separate helper process (for example a PowerShell script launched before
  the call) that sends Win32 messages to the dialog's controls.
- Never use plain `SendMessage` from the Excel thread or `AttachThreadInput`: both hang
  Excel. Run the whole pass in the background by chaining `Application.OnTime`, with no
  `MsgBox`, writing progress to cells of a control sheet.
- **Smart View does not refresh formatted reports while a macro is running.** Verify the
  refresh in a separate step, in small chunks via `OnTime`.
- To check that a report refreshed to the new POV, compare against **all** cells that
  show the year on that tab; reports often have fixed comparative columns, so "the
  first cell with a year" gives false results.
- Report sheets can be detected by a sheet-level property named `MR_METADATA_XML`.
- In the **Member Selection** dialog: the Find button ignores synthetic clicks, but
  ENTER in the search box works. A search expands and selects the member but does not
  tick its checkbox: send SPACE, after counting the ticked boxes (if a version ticks
  it already, SPACE would untick it). Detect a successful search by a change in the
  tree's node count (`TVM_GETCOUNT`) or in the ticked count; the focused node
  (`TVGN_CARET`) does not move. Node handles and states can be read across processes
  (`TVM_GETNEXTITEM`, `TVM_GETITEMSTATE`); node text cannot.
- Edit POV replaces the report tab (new tab, old one deleted): never keep anything you
  need on it.

This approach is fragile by nature (dialog layouts change between versions). Prefer
one workbook per POV inserted by the user, or server-side report bursting.

## PDF export

- Report titles are EMF vector pictures that Smart View regenerates on every refresh.
  Excel's PDF exporter ignores their character spacing and the titles come out
  scrambled, even at 100% scale. Before exporting, **convert each title picture to a
  bitmap** at the same position, hide the original, export, then **restore** the
  originals.
- Give temporary shapes a unique prefix (for example `RASTER_`) and delete leftovers
  at start, in case a previous run was interrupted. Set `Application.EnableEvents =
  False` only during the shape swap and turn it back on before any Smart View call. Compare shape names case-insensitively. Keep a constant to
  switch the behaviour off.
- One page per sheet: `PageSetup.Zoom = False`, `FitToPagesWide = 1`,
  `FitToPagesTall = 1`.

## Deploying macros to users

- Ship code as `.bas` modules the user imports (VBA editor > File > Import), or as an
  add-in. Rebuilding and overwriting their `.xlsm` loses whatever they added to it.
- Replacing an `.xlsm` repeatedly while Excel has it open can corrupt the VBA project:
  every `Application.Run` then fails with `0x800A9C68`, and closing and reopening the
  workbook is not enough. Close Excel completely (no `EXCEL.EXE` left), reopen, sign
  in to Smart View again.
- Protect the VBA project with a password only at the end, and keep that password in
  the team's password manager, never in the code or the workbook.

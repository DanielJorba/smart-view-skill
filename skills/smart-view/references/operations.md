# Operating Smart View through the bridge

Everything here runs in the user's Excel with `scripts/sv.ps1`. Before the first call
of a session: `sv.ps1 status` (see [bridge-setup.md](bridge-setup.md) if anything is
false). Tell the user you are about to work in their Excel.

Conventions:
- `""` as workbook = the scratch workbook (created on demand). Pass a workbook name
  only to read a user's workbook, or when the user asked you to work in it.
- Addresses are A1 references on the given sheet (`"A5"`, `"B2:B40"`).
- Lists (for example POV members) are separated by `|`.
- Every result has `ok`, `rc`, `rcText`, `error` and `data`.

## Contents

1. Catalog: bridge, UI-only, off limits
2. Common workflows
3. The ad hoc ribbon tabs (what each button maps to)

## 1. Catalog

### Through the bridge (read-only or local layout only)

| Area | Wrappers |
|---|---|
| Session | `Ping`, `Workbooks`, `Scratch`, `ReadTsv` (via `sv.ps1 read`), `WriteTsv` (scratch only, refuses `HsSetValue`), `LastError` |
| Cell functions | `Refresh` (one sheet, runs `Hs*` functions), `ReadFormulas` (via `sv.ps1 formulas`), `AuditFunctions` (via `sv.ps1 hs-audit`), `CellValue` (one intersection, nothing written) |
| Connections | `ConnectionExists`, `CreateConnection`, `UIConnect` (cloud), `Connect` (on-prem only), `Disconnect`, `Connected`, `ConnectionInfo`, `SetActiveConnection`, `SetConnAliasTable`, `SetAliasTable` |
| Ad hoc grid | `Retrieve` (refresh), `RetrieveRange`, `RefreshAll` (one workbook), `ZoomIn`, `ZoomOut`, `KeepOnly`, `RemoveOnly`, `Pivot`, `PivotToPOV`, `PivotToGrid`, `Undo`, `Redo` |
| POV | `GetPOV`, `SetPOV` (ad hoc), `SetPages` (forms), `BackgroundPOV`, `GetDimensions`, `GetMembers`, `PagePOVChoices` |
| Sheet state | `ContentType` (empty, adhoc, form, report), `SheetInfo`, `CellPOV`, `IsCellWritable`, `DatabaseNote` |
| Options | `GetOption`, `SetOption` (unified `HSV_*` indices, preferred), `GetSheetOption`, `SetSheetOption` (items 1-16), `GetGlobalOption`, `SetGlobalOption` (own numbering) |
| Forms | `ListDocuments`, `OpenForm`, `AdhocOnForm` |
| Members (Essbase only) | `FindMember`, `MemberInfo`, `MemberProperty`, `Children`, `Parent`, `Ancestor`, `QueryMembers`, `OutlineMemberInfo`, `IsUDA`, `IsAttribute` |
| Variables | `SubstitutionVariables` (Essbase only), `ListUserVariables` |
| Drill-through | `DrillThroughReports`, `ExecuteDrillThrough` |
| Queries | `ExecuteQuery` (MDX `SELECT` only, Essbase only) |
| Menus | `ExecuteMenu` (refuses data-changing commands) |

Ad hoc operations change the **layout** of a sheet, not server data. On a user's
workbook they still change what the user sees: do them on the scratch workbook unless
asked.

### UI-only: the user clicks, Claude reads the result

- Task lists (open a task, open the artifact it points to).
- Formatted reports with prompt values other than the defaults, *Refresh Report*, Edit
  POV of reports. Opening a report with its default prompts is automated (below).
- EPM Admin Extension dimension sheets: now automated with `scripts/panel.ps1`
  (see [structure-discovery.md](structure-discovery.md)). Panel context-menu actions are
  not automated on purpose: several change the server (Refresh Database, rules).
- Member Selection dialog, Query Designer, Visualize, Smart Slices, job console.

### Reports from the panel (automated)

```powershell
scripts/panel.ps1 tree 4                                         # find the report
scripts/panel.ps1 open-report <APP> "Reports/<folder>/<report>" <url>
```

*Tested on EPM Cloud Planning with Smart View 26.200 (Oct 2026):* opening a report from
the panel shows a **Select Prompts** dialog (WinForms) with each prompt and its default
value; `open-report` reports them, accepts the defaults and waits until the report opens
as a **new, unconnected workbook** named after the report (`SheetInfo` type "Reporting
and Analysis Documents"). It may take a minute for large reports. Then read its
structure with `sv.ps1 read <book> <sheet> A1:N20`: figures stay hidden, but numeric
member codes (for example entity codes) show as `<n>` too, and text members such as
customer or contract names come through as they are: treat them as client data.
Subfolders of Reports load only when opened; if a folder shows empty, list it again with
`tree` after opening it, or ask the user. To run with other prompt values, the user picks
them in the dialog.

Give the exact clicks, wait for the user, then `sv.ps1 call Workbooks` to find the new
sheet and `sv.ps1 read` it.

### Off limits (the user does it, if they want)

Submit data (any variant, including `HsSetValue` cells), calculate, business rules and
rules on forms, calc scripts, consolidate or translate, copy version, approvals and
promotions, adjust or spread, supporting details, cell text and comments, attachments
(linked reporting objects), lock and unlock, offline sync back, clearing data,
substitution-variable changes, metadata submit or refresh database, data integration
or job launches. Explain what it does and where to click; never do it.

## 2. Common workflows

### Open an ad hoc grid on a cube and see its dimensions (no data)

```powershell
sv.ps1 call Scratch Grid true
sv.ps1 call UIConnect "" Grid "<friendlyName>"
sv.ps1 call SetOption 33 true "" Grid           # navigate without data, this sheet only
sv.ps1 call Retrieve "" Grid
sv.ps1 call GetDimensions "" Grid
sv.ps1 read "" Grid
```

The first `UIConnect` to a server opens Smart View's sign-in window: tell the user
to sign in there. Then continue with [structure-discovery.md](structure-discovery.md).

### Audit or refresh a sheet of HsGetValue formulas

```powershell
sv.ps1 hs-audit "<Book.xlsx>" "<Sheet>"
sv.ps1 call Refresh "<Book.xlsx>" "<Sheet>"
```

See [cell-functions.md](cell-functions.md#working-with-function-sheets-through-the-bridge).

### Refresh a user's workbook

```powershell
sv.ps1 call Workbooks                              # exact names
sv.ps1 call ContentType "<Book.xlsx>" "<Sheet>"    # adhoc, form or report?
sv.ps1 call Retrieve "<Book.xlsx>" "<Sheet>"       # one ad hoc sheet (Refresh for function sheets)
sv.ps1 call RefreshAll "<Book.xlsx>"               # every connected sheet
```

Formatted reports refresh with the ribbon's Refresh Report; if `Retrieve` does not
update them, ask the user to click it. Refreshing does not change server data, but it
replaces what is on the sheet: warn if the user has unsaved work there.

### Change the POV and refresh

```powershell
sv.ps1 call GetPOV "" Grid
sv.ps1 call SetPOV "" Grid "Year#FY26|Scenario#Actual"
sv.ps1 call Retrieve "" Grid
```

Use the member format confirmed in [vba-api.md](vba-api.md) for `HypSetPOV`. If a
dimension is in a header row of the grid rather than on the POV, change that cell
instead.

### Zoom, keep only, pivot

```powershell
sv.ps1 call ZoomIn "" Grid "A3" 1          # level codes in vba-api.md
sv.ps1 call KeepOnly "" Grid "A3:A10"
sv.ps1 call PivotToPOV "" Grid "A3"
sv.ps1 call Undo "" Grid
```

### Open a form and read it

```powershell
sv.ps1 call Scratch Form true
sv.ps1 call UIConnect "" Form "<friendlyName>"
sv.ps1 call ListDocuments "" Form "<friendlyName>" "<folder path>"
sv.ps1 call OpenForm "" Form "<folder path>" "<form name>"
sv.ps1 read "" Form
```

Change a form's page members with `SetPages` ("Dimension#Member|..."), never
with `SetPOV`. Never type into a form's cells.

*Tested on FCCS (21-oct-2026):* `ListDocuments` works with the folder path `/Forms`
("" returns -15), but `OpenForm` returns -41 (provider not supported) for any path. On
FCCS the user opens the form from the Smart View panel and Claude reads the sheet. To analyse a form freely, `AdhocOnForm` turns it into
an ad hoc grid on the same sheet.

### What is behind this number?

```powershell
sv.ps1 call CellPOV "<Book.xlsx>" "<Sheet>" "F12"
sv.ps1 call DrillThroughReports "<Book.xlsx>" "<Sheet>" "F12"
```

`CellPOV` returns every dimension and member of that intersection; drill-through lists
the source-system reports available for the cell, if any.

### Build an ad hoc grid from a design

Follow [structure-discovery.md](structure-discovery.md#8-from-profile-to-grid), write
the labels with `WriteTsv` on the scratch sheet, connect, keep navigate without data
on, retrieve and read it back.

## 3. The ad hoc ribbon tabs

When a sheet holds an ad hoc grid, Smart View adds a provider tab (for example
*Planning Ad Hoc* or *Essbase*); forms get a *Planning* tab. Buttons vary by provider
and version. Map them by what they do:

| Button | How Claude handles it |
|---|---|
| Zoom In / Zoom Out / Keep Only / Remove Only / Pivot | Bridge |
| Refresh | Bridge (`Retrieve`) |
| POV | Bridge (`GetPOV`, `SetPOV`) |
| Undo / Redo | Bridge |
| Change Alias | Bridge (`SetAliasTable`) |
| Member Selection | UI-only; use member queries instead |
| Preserve Format, Cascade, Visualize, Query Designer | UI-only |
| Drill-through | Bridge (`DrillThroughReports`, `ExecuteDrillThrough`) |
| Member Information | Bridge (`MemberInfo`) or UI |
| Instructions (forms) | Read-only UI |
| Submit Data, Calculate, Business Rules, Approvals, Copy Version, Adjust, Supporting Details, Cell Text, Attachments | Off limits |
| Navigate without data | Bridge (`SetOption 33 true "" <sheet>`) |

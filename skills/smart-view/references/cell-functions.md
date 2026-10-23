# Smart View cell functions (Hs*)

Source: Oracle *Smart View for Office User's Guide*, EPM Cloud library, chapter
"Functions" (October 2026). Explanations are original.

## Functions and providers

| Function | Does | Providers |
|---|---|---|
| `HsGetValue` | Reads the value of one intersection | Essbase, HFM, Planning (cloud and on-prem), Planning Modules, FCCS, Tax Reporting, Narrative Reporting |
| `HsSetValue` | **Writes** a value on Submit | Essbase, HFM, Planning family, FCCS, Tax Reporting |
| `HsDescription` | Member description (Essbase: the alias) | HFM, Essbase, Planning family, FCCS, Tax Reporting |
| `HsAlias` | Member alias from a given alias table | Planning family, FCCS, Tax Reporting |
| `HsGetSheetInfo` | Information about the sheet | Any |
| `HsGetVariable` | Substitution variable value | Essbase |
| `HsCurrency`, `HsLabel`, `HsGetText`, `HsSetText` | Entity currency, default label, cell text | HFM |
| `HsGetSharedConnectionURL` | Shared connections URL in use | Any |

## Syntax

```text
=HsGetValue("<connection>", "<POV>")
=HsSetValue(<value>, "<connection>", "<POV>")
=HsDescription("<connection>", "Dimension#Member")
=HsAlias("<connection>", "Dimension#Member", "<OutputAliasTable>", "<MemberNameFromAliasTable>", <ReturnDistinctName>)
=HsGetSheetInfo("<item>")
=HsGetVariable("<connection>", "<variable>")      ' or HsGetVariable("<variable>") with the default connection
```

### Connection argument

- Empty: the default connection.
- `HsActive`: the connection associated with the sheet.
- A **private connection name**. Not portable: it must exist with exactly that name on
  every PC that opens the workbook.
- `"WSFN|ProviderType|Server|Application|Database"` for shared connections, with
  `ProviderType` case-sensitive: `Essbase` (on-prem), `HFM`, or `HP` (Planning, FCCS,
  Tax Reporting).
- `"URLFN|URL|Server|Application|Database"`. For EPM Cloud the server part may be blank.
  Oracle recommends it for workbooks shared between users, often stored in a cell and
  referenced: `=HsGetValue($B$1, "Account#Sales;Period#Jan")`.

### POV argument

- Pairs `Dimension#Member`. Qualify by parent with a period: `Entity#Europe.E100`.
- Either one argument per pair (`"Account#Sales","Period#Jan"`) or one argument with
  semicolons (`"Account#Sales;Period#Jan"`).
- Avoid `#` and `;` in names: they are delimiters.
- Names or aliases from **one** alias table; set that table on the connection. An
  alias equal to another member's name returns an invalid member.
- Dimensions left out take the default POV. Name every dimension that matters
  (scenario, version, currency, view) instead of trusting defaults.
- Build the POV string from cells for maintainable sheets:
  `=HsGetValue($B$1, "Account#"&$A5&";Period#"&C$3&";Scenario#Actual")`.

### HsGetSheetInfo items

1 connected, 2 sheet type, 3 server, 4 application, 5 cube, 6 URL, 7 provider,
8 provider URL, 9 friendly name, 10 alias table, 12 description, 13 last retrieved.

## When they calculate

- Smart View functions run on **Smart View Refresh** (active sheet) or **Refresh All**
  (workbook), not on Excel recalculation. A typed formula is validated on refresh.
- `HsSetValue` runs on **Submit Data**, before the `HsGetValue` cells, which only run if
  the submit succeeds. Treat any sheet with `HsSetValue` as a data-entry sheet: refresh
  is safe, submit writes to the server.
- Opened offline (sign-in dismissed), cells keep their last values.
- In VBA: `HypMenuVRefresh`, `HypMenuVRefreshAll`, `HypRetrieveAllWorkbooks`.
  Option 121 (`HSV_REFRESH_SELECTED_DEPENDENT_FUNCTIONS`) runs dependent functions
  first.

## Errors in cells

| Shows | Means |
|---|---|
| `#NAME?` | Excel does not know the function: Smart View missing or its function add-in link broken |
| `#NO CONNECTION` | Not connected or not signed in, or a bad connection token |
| `#Cube Not Selected` | Connected, but not at cube level |
| `#INVALID MEMBER` / `#INVALID DIMENSION` | Unknown member or dimension (check alias table and spelling) |
| `#INVALID` | Invalid metadata intersection |
| `#NO ACCESS` | No security access to the intersection |
| `#NO DATA` | No data (shown with the missing label) |
| `#LOCKED` / `#READ ONLY` | Cell locked or read-only (relevant for `HsSetValue`) |
| `#NEEDS REFRESH` | Refresh needed |

### Fixing `#NAME?`

Excel stores the full path of the functions add-in (`HsTbar.xla` in the Smart View
`bin` folder) in each workbook. On a PC where Smart View lives elsewhere, links break.
Use Smart View ribbon > Functions > **Fix Links**, then Refresh. Oracle recommends the
same install folder for everyone sharing workbooks. If Smart View itself is not
loaded, see [troubleshooting.md](troubleshooting.md#3-name-in-cells-with-hsgetvalue).

## Working with function sheets through the bridge

### Audit a sheet of HsGetValue formulas

```powershell
sv.ps1 hs-audit "<Book.xlsx>" "<Sheet>"            # whole used range
sv.ps1 hs-audit "<Book.xlsx>" "<Sheet>" "C6:N41"   # one block
sv.ps1 formulas "<Book.xlsx>" "<Sheet>" "C6:C10"   # raw formulas, if you need them
```

`hs-audit` finds every Smart View function in the range, **evaluates its arguments
with Excel** (so a POV built as `"Account#"&$A6&";Period#"&C$5` shows the real
members), and reports:

- `connections`: each connection argument, its kind (private, default, sheet with
  `HsActive`, shared `WSFN`, URL `URLFN`) and, for private names, whether it exists on
  this PC. A private name missing here is the classic reason a workbook works for the
  author and shows `#NAME?` or `#NO CONNECTION` for everyone else.
- `dimensions`: every dimension named in the POVs, in how many cells, how many
  distinct members and a sample. **`missingInCells` > 0 means some cells leave that
  dimension to the default POV**: the usual cause of numbers that do not match an ad
  hoc grid.
- `errors`: cells showing `#...` with their evaluated arguments, so you can see the
  member that fails.
- `writeCalls`: cells with `HsSetValue` / `HsSetText`. The sheet is a data-entry
  sheet: refreshing is safe, Submit writes. Tell the user.
- `brokenAddinLinks`: formulas that carry a hard-coded path to `HsTbar.xla`. Fix with
  Smart View > Functions > Fix Links (the user clicks it).

Read the audit before proposing fixes: it answers most "why does this cell say X"
questions without touching data.

### Refresh

Cell functions compute only on a Smart View refresh:

```powershell
sv.ps1 call Refresh "<Book.xlsx>" "<Sheet>"     # one sheet (HypMenuVRefresh)
sv.ps1 call RefreshAll "<Book.xlsx>"            # every sheet of that workbook
```

`Retrieve` (HypRetrieve) is for ad hoc grids; on a sheet with only functions use
`Refresh`. Refreshing brings figures into the sheet: only when the user asks.

### Check one value without touching the sheet

```powershell
sv.ps1 call CellValue "" Grid "Account#Sales|Entity#E100|Period#Jan|Years#FY26|Scenario#Actual|Version#Final"
```

`CellValue` (HypCell) reads one intersection through the sheet's connection and
writes nothing. It is the quickest way to tell whether a wrong `HsGetValue` is wrong
because of its POV or because of the data. It returns data: use it on request.

### Build a function sheet

*Tested on FCCS (21-oct-2026):* do **not** associate a function sheet with a connection
(no `UIConnect` on it): Smart View then treats its labels as an ad hoc grid and fails
with "Grid is Invalid". The formulas carry their own connection; `Refresh` is enough.
Excel limits each **text literal inside a formula to 255 characters**: a full FCCS POV
written as one quoted string is longer. Split it with `&`, or keep the POV in a cell
and reference it (`=HsGetValue($A$1,$A$3)`). Long formulas are fine otherwise.

Generate formulas instead of typing them, so every cell follows one pattern:

- Put the connection in one cell (for example `B1`), ideally as
  `URLFN|https://<host>/HyperionPlanning/SmartView||<APP>|<CUBE>` so the workbook works on
  any PC, or a private name every user has.
- Put the members of rows and columns in label cells and build the POV with `&`:
  `=HsGetValue($B$1,"Entity#"&$A6&";Account#"&$B6&";Period#"&C$5&";Years#"&$B$2&";Scenario#Actual;Version#Final;Currency#Local")`.
- Name **every** dimension of the cube in the POV (see the dimension list in the cube
  profile, [structure-discovery.md](structure-discovery.md)). Omitted dimensions take
  the default POV and change silently with it.
- On the scratch workbook the bridge can write them: `WriteTsv` accepts formula text in
  English syntax with commas (`=HsGetValue(...)`), even on a localized Excel. It refuses
  `HsSetValue` and `HsSetText`.
- After writing, `Refresh`, then `hs-audit` to confirm there are no error cells.

### Convert an ad hoc grid into functions

Build the formula for each data cell from its intersection: `CellPOV` returns every
dimension and member behind a cell. Replace the literal members with references to the
row and column labels so the result stays maintainable. Smart View's own *Copy Data
Points* / *Paste Data Points* does the same from the UI, with literal members.

## When an HsGetValue does not match the ad hoc grid

Check in this order:

1. **Missing dimensions.** `hs-audit` → `missingInCells`. Every dimension the ad hoc
   grid shows (rows, columns and POV) must be in the formula.
2. **Different members for the same idea.** Names vs aliases (one alias table per
   connection), `Years` vs `Year`, `FY26` vs `2026`, a parent vs its level-0 member.
3. **FCCS system dimensions.** Besides Account, Entity, Period, Years and Scenario,
   FCCS cubes carry View (Periodic, YTD, QTD), Consolidation, Data Source, Movement,
   Intercompany and Currency, plus Multi-GAAP and custom dimensions when enabled. A
   formula that omits View or Consolidation reads whatever the default POV points to.
4. **Stale values.** The formula has not been refreshed since the data changed (Excel
   recalculation does not fetch).
5. **Different connection.** The cell points to another environment or cube than the
   ad hoc sheet (`hs-audit` → `connections`; ad hoc sheet → `SheetInfo`).
6. **Security.** `#NO ACCESS` or empty values for intersections the user cannot read.

## Choosing between functions and ad hoc

- Functions suit fixed-format reports (a template where each cell is a known
  intersection) and mixing several cubes on one sheet.
- Ad hoc grids suit exploration and lists, and are much faster for large ranges:
  thousands of `HsGetValue` cells mean thousands of lookups.
- Never mix: `HsGetValue` formulas inside an ad hoc grid are wiped or misread on refresh.

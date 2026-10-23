# Preparing metadata changes (reclassifications, history, aliases)

Claude prepares; the user executes. Submitting a dimension sheet and refreshing the
database change the application for everyone, so Claude never does either.

## Workflow

1. Explore inside the future "before" sheet (no extra sheet): open it with
   `panel.ps1 open-dimension <APP> <Dimension> <url> <Sheet>`, zoom it
   (see [structure-discovery.md](structure-discovery.md)) and read the current rows of
   every member involved: the member, its old and new parent, siblings that matter, and
   any precedent (for example an existing `_Hist` member).
2. Exactly **two connected dimension sheets**, nothing else for this purpose (no extra
   working sheet, no plain-Excel copies): open the dimension twice with
   `panel.ps1 open-dimension <APP> <Dimension> <url> <Sheet>`, for example `Dim_1234_Before`
   and `Dim_1234_After`. In each, type only the affected members in column A and
   `Retrieve`, so they show the current rows. Then write the change only in the "after"
   sheet and do **not** refresh it again (a refresh brings back the server values).
   Connected sheets let the user see what Smart View would reject (locked or read-only
   intersections) before submitting. List the changed cells in the chat.
3. List for the user what the default columns do not show and must be checked: UDAs,
   attribute associations, formulas, security on the new parent, rules or forms that
   reference the member or its parent, and whether the member exists in other cubes.
4. The user applies the change (in the dimension sheet or in the web Dimension Editor),
   submits and refreshes the database. Claude may write the change into the dimension
   sheet when asked, listing every cell it touched; it never submits.
5. After the user's refresh, reopen the dimension and confirm the rows match "After".

If the user has not said the new parent (or new values), ask before writing the "after"
sheet; do not write placeholders. To discard prepared sheets the user no longer wants,
`sv.ps1 drop-sheet <Before> <After>` (scratch workbook only; nothing was submitted, so
nothing changes on the server).

## Does the member have data? (check before every move)

Use the tool, once per cube (input BSO and reporting ASO when there is one):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/member-data.ps1 -Profile "<environment profile>" -Dimension <Dim> -Member <name>
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/member-data.ps1 -Url <url> -App <APP> -Cube <CUBE> -CubeType BSO|ASO -Dimension <Dim> -Member <name>
```

With `-Profile`, it takes url, application, cube (the reporting cube if the profile has
one, else the work cube), cube type and `dataCheckPov` from `connections.json`.

It returns `hasData`, the number of non-zero cells and the Scenario / Version / Year
combinations with data, no amounts, and deletes its working sheet.

**On a BSO cube with no reporting cube, add `-Pov`.** Every dimension that is not explored
stays at its root, and in BSO cubes the roots of Account, Period, Currency, data-type
dimensions and the like are usually label-only or not aggregated: the check then returns 0
for a member full of data. *Tested on EPM Cloud Planning (Oct 2026):* an entity with data
gave 0 with everything at the root and hundreds of non-zero cells with
`-Pov "Period#YearTotal;Account#<stored parent>;Currency#Local;..."`. Without `-Pov`, a 0
comes back with a `warning`: do not report "no data" from it. Take the members from the
user or from a form that shows data, and list their children first (zoom one level with
navigate-without-data) if you need names. Answering "does it have data?" counts as a request for data.
Find the cube names with `panel.ps1 cubes <APP> <url>`. Only checking needs this one
sheet; the two before/after dimension sheets are for preparing the change. What the tool
does, for reference:

Before moving or renaming a member, find the Scenario / Version / Year combinations where it
has data. Tested on EPM Cloud Planning (Oct 2026):

- One sheet. Build it from the default grid: `PivotToGrid` Scenario and Version onto the
  existing row-dimension cell, Year and the member's dimension onto the column-dimension
  cell; write the member in its header cell; zoom Scenario, Version and Year to bottom
  level (`ZoomIn` level 2). Every other dimension stays at its top.
- Refresh with data **without** row suppression: with suppression on, the refresh failed
  (rc 1000) on these cubes.
- Count with Excel formulas that return text, never with masked reads: "no data" shows
  as 0 in one cube and as an empty cell in another. For example
  `="nonzero="&SUMPRODUCT(ISNUMBER(D3:AA2544)*(D3:AA2544<>0))`, and per row one `IF` per
  column joined with `&`:
  `=IF(AND(ISNUMBER(F3),N(F3)<>0),F$2&",","")&IF(AND(ISNUMBER(G3),N(G3)<>0),G$2&",","")&...`
  to list the columns with data. Do **not** use array formulas (`TEXTJOIN(IF(range...))`):
  written from the bridge they get implicit intersection and return empty, a false
  "no data". If the global count is above zero but the per-row lists are empty, the
  check is broken: never conclude "no data" from it. Report only those combinations,
  never amounts.
- Look at children, not only at the top: a Scenario or Version root can show nothing
  while its children have data. Qualified rows such as `[Parent].[Child]` are other
  views of the same data, not extra data.
- Which cube: the input cube (BSO) holds all history, but its upper levels may not be
  aggregated (rules often leave some dimensions at level 0), so zeros at the top do not
  prove there is no data. The reporting cube (ASO) aggregates on the fly and confirms it,
  for everything pushed to it up to the last aggregation. Check both; if they disagree,
  say that an aggregation or push is pending.

## Keeping history when a member moves

- Moving a stored member to another parent moves its whole history with it: past periods
  then aggregate under the new parent.
- To keep the past under the old parent, the pattern is: **rename** the existing member to
  `<name>_Hist` in the web Dimension Editor (the rename keeps the data) and create a new
  `<name>` under the new parent. Do the rename in the web editor, not from Smart View, and
  rename to a name that does **not** already exist: renaming onto an existing member
  makes the refresh drop the history (field report, 2026).
- Never create the `_Hist` member beforehand from Smart View.

## Alias-only changes

Changing an alias has no data impact; it does change what users and `HsGetValue` formulas
that use aliases resolve to. Check formulas that use the old alias.

# Understanding a cube's structure

Read this before building any ad hoc grid, writing a POV, or answering "where does
this member live". A grid built on assumptions fails in the classic ways (zeros, out of
place members, wrong alias resolving to another member). A grid built from the real
outline refreshes the first time.

## Contents

1. What to find out
2. Sources, from most to least complete
3. Route A: EPM Admin Extension dimension sheets (administrators)
4. Route B: member queries through the bridge
5. Route C: ad hoc exploration without data (works everywhere)
6. Route D: metadata files the user provides
7. The cube profile
8. From profile to grid

## 1. What to find out

For the **cube**: every dimension and its type (account, time, period, scenario,
version, entity, currency, view or generic, attribute), the alias tables, which cubes
exist in the application (input block storage vs reporting aggregate storage), the
substitution variables (current year, current month...) and user variables.

For each **dimension**: the hierarchy (parent-child), its depth, the level-0 members,
the generation and level names if defined, shared members (a member appearing under
two parents), and alternate hierarchies.

For the **members** that matter for the task: name and aliases, storage (stored,
dynamic calc, label only, never share, shared), consolidation operator (`+ - * / ~ ^`),
data type, time balance for accounts, formula, UDAs, attribute associations.

And where **data actually lives**: which Version, Scenario, Currency, View or data type
members the business rules write to. That is not in the outline; it is in the rules
(their `FIX` statements) or in a working report.

## 2. Sources, from most to least complete

| Route | Gives | Needs |
|---|---|---|
| A. EPM Admin Extension dimension sheet | Every member with every property | Service administrator role, the extension, a click by the user |
| B. Member queries via bridge | Children, descendants, levels, generations, search, member properties, formulas, UDAs, attributes | **Essbase connections only** (Oracle documents none of them for Planning, FCCS or HFM) |
| C. Ad hoc exploration without data | Hierarchy, names, aliases, level/indent | Any connection; no figures retrieved |
| D. Metadata exports | Everything, offline | The user provides the file (metadata export job or snapshot) |

Use the cheapest route that answers the question. On **Planning, FCCS or Tax
Reporting** the API has no member queries: use C for hierarchies and A or D for
properties. On **Essbase**, B answers most questions directly.

## 3. Route A: EPM Admin Extension dimension sheets

With the EPM Admin Extension installed (`sv.ps1 version` lists it) and an
administrator user, the Smart View panel shows the application's dimensions and can
open one in a sheet with a row per member and a column per property (parent, aliases,
data storage, consolidation, formula, UDAs, attributes, plan types...).

**Automated (tested on Planning, 21/22-oct-2026):** `panel.ps1` opens the panel if it is closed
(`ensure`) and points it at a server by typing the URL in the Private Connections box
(`server <url>`); `open-dimension` does both when given the URL. Then run

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/panel.ps1 open-dimension <APP> <Dimension> https://<host>/HyperionPlanning/SmartView
```

It expands `<APP> > Dimensions` (English or Spanish labels), selects the dimension and
presses Enter on it, on a **blank** scratch sheet (a sheet with content makes Smart View
ask what to do with it). Then `ZoomIn` on A2 with level 1 gives every member with its
properties. `panel.ps1 tree` shows what the panel currently has. Synthetic mouse clicks
are ignored by the panel; Enter works.

Manual fallback:

1. Ask the user to open it from the Smart View panel: application > **Dimensions** >
   **double-click the dimension**. A sheet opens with the dimension's top member and its
   predefined property columns; from there the hierarchy can be expanded. This is
   UI-only. To go straight to one branch, the user can type that member on the sheet
   and zoom, or Claude zooms it through the bridge on that sheet.
2. Read it with `sv.ps1 read "<workbook>" "<sheet>"` (increase maxRows for large
   dimensions, or read in ranges).
3. **Treat that sheet as read-only.** It has its own submit, which changes the
   application's metadata. Never edit it, never trigger submit or refresh-database
   actions; close it without saving when done, or let the user decide.

## 4. Route B: member queries through the bridge (Essbase)

All take a sheet connected to an **Essbase** cube. Values for predicates and options
are in [vba-api.md](vba-api.md).

```powershell
sv.ps1 call FindMember "" Explore "Sales"                      # dimension, alias, generation and level names
sv.ps1 call Children "" Explore "Total_Entity"                  # direct children
sv.ps1 call Parent "" Explore "E100"
sv.ps1 call QueryMembers "" Explore "Total_Entity" 2 1          # descendants, member names only
sv.ps1 call QueryMembers "" Explore "Account" 3 4 "Account"     # bottom level, names and aliases
sv.ps1 call QueryMembers "" Explore "Sales*" 12 1 "Account"     # wildcard search
sv.ps1 call MemberInfo "" Explore "EBITDA"                      # every property Smart View exposes
sv.ps1 call OutlineMemberInfo "" Explore "Account" "EBITDA" 2  # formula
sv.ps1 call OutlineMemberInfo "" Explore "Account" "EBITDA" 3  # UDAs
sv.ps1 call SubstitutionVariables "" Explore "<APP>" "<CUBE>" ""
sv.ps1 call ListUserVariables "<friendlyName>"
```

On other providers expect `SS_NOSUPPORT_PROVIDER` (-41) or empty results: switch to
route C, and record in the cube profile what works so you do not retry.

## 5. Route C: ad hoc exploration without data

This works on any provider and never brings figures back.

1. `sv.ps1 call Scratch Explore true` (fresh sheet), then
   `sv.ps1 call UIConnect "" Explore "<friendlyName>"`.
2. Set options on the `Explore` sheet only (`SetOption <idx> <value> "" Explore`,
   unified indices in vba-api.md):
   - navigate without data **on** (`HSV_NAVIGATE_WITHOUTDATA` = 33): no data retrieved;
   - suppress missing and zero rows **off** (6, 7) and columns off (41, 42): with no
     data everything would be suppressed;
   - member display = 1, name and alias (`HSV_MEMBER_DISPLAY` = 16);
   - indentation = 1, sub-items (`HSV_INDENTATION` = 5), plus Excel indentation on
     (46) so `read` reports the depth in its first column (if 46 is not available,
     depth shows as leading spaces in the member cell).
   Read the current value first (`GetOption <idx> "" Explore`) so you can report what
   you changed.
3. `sv.ps1 call Retrieve "" Explore`. On an empty connected sheet Smart View lays out a
   default grid: one or more dimensions on rows and on columns (it varies by
   application), the rest on the POV. Read it before pivoting.
   Set the sheet's alias table to `None` first (`SetAliasTable "" Explore None`) if you
   need member names: with an alias table, labels and even dimension names are aliases.
4. `sv.ps1 call GetDimensions "" Explore` lists **all dimensions** and where each sits.
   This is the authoritative dimension list for that cube.
5. To walk one dimension (tested method, FCCS 21-oct-2026): retrieve the default grid,
   then **swap**: write the target dimension's root in the first row-dimension cell (and
   its alias cell) and the displaced row dimension where the target was, and `Retrieve`
   again. A free-form grid typed on an empty connected sheet failed (rc 1000). Reuse one
   sheet for all dimensions instead of one per dimension. Then
   `ZoomIn` on that cell with level 0 (next level) for large dimensions, 1 (all
   levels) for small ones or 2 (bottom level only), and `read` the result. The indent column gives the depth; with "member
   name and alias" display the next column holds the alias.
   *Tested on Planning with a floating POV toolbar (EPM Cloud, 21-oct-2026):* the POV
   dimensions are not cells of the sheet, so the swap cannot find them; `PivotToGrid
   "" <sheet> <Dimension> A3` brings one onto the rows (or `C1` onto the column headers)
   and works. `HypSetPOV` returned -3 and `HypGetPOV` returned nothing there: put
   scenario, version and currency in the grid as header rows instead (PivotToGrid to the
   column header cell, then write the members).
   The cell passed to `PivotToGrid` must hold the label of a dimension already on that
   axis; any other cell gives 57705 "Invalid selection". A dimension already on the grid
   cannot be pivoted onto it again (rc -15): write the member in its cell instead.
   To walk the next dimension on the same sheet: `PivotToGrid` it onto the row
   dimension cell, `PivotToPOV` the one it displaced, `Retrieve`. If zoomed rows of the
   previous dimension remain below, start a fresh sheet (`Scratch Explore true`):
   `WriteTsv` cannot clear cells.
   `ZoomIn` level 2 (bottom) keeps the root row and, on dimensions with alternate
   hierarchies (for example Period with `_ac`, `_Stored`, `_LTM` members), lists their
   level-0 members too.
6. For large dimensions (thousands of members), explore top-down: zoom one level,
   decide which branch matters, `KeepOnly` it, zoom again. Do not try to dump the whole
   dimension through the grid.
   *Tested on FCCS (21-oct-2026):* with member display 1 the alias lands in the column
   next to each member, and without option 46 the depth shows as 5 leading spaces per
   level in the member cell (indent column stays 0). `GetDimensions` on the default
   grid listed all 14 dimensions with types (1 column, 0 row, 2 POV).
7. `CellPOV` on any data cell returns the full member intersection behind it, which is
   a quick way to confirm how a grid is being read.

When done, leave the user's options as they were (sheet options live only on the
scratch sheet; global options must be restored if you changed them).

## 6. Route D: metadata files the user provides

A metadata export from the application (an export job in the web interface) or the
dimension files inside an application snapshot have every property for every member.
If the user has one, read it directly: it is complete, offline and costs no server
calls. Do not ask for credentials to fetch it; the user downloads it.

## 7. The cube profile

Not to be confused with the **environment profile** in `connections.json`
([connections.md](connections.md)), which says which cube to use and how to connect.
The cube profile describes the cube's dimensions and traps.
Write what you learn to `%LOCALAPPDATA%\smart-view-skill\profiles\<friendlyName or app_cube>.md`
(outside the skill folder; it describes client metadata). Read it first next time and
refresh it when it is older than the last metadata change the user mentions.

```markdown
# <friendlyName>
Captured: <date> · Provider: <Planning|Essbase|...> · Cube: <name> (BSO input | ASO reporting)
Alias tables: Default, <others> · Functions that work here: QueryMembers yes, MemberInfo no, ...

| Dimension | Type | Default axis | Root | Depth | Level-0 count | Notes |
|---|---|---|---|---|---|---|
| Account | account | row | Account | 7 | ~1,200 | dynamic calc ratios under KPIs |
| Period | time | column | YearTotal | 3 | 12 | Q1..Q4, BegBalance |
| Scenario | scenario | POV | Scenario | 2 | 6 | Actual, Budget, Forecast... |
| Version | version | POV | Version | 2 | 4 | data lives in Final (rules write there) |
| Currency | currency | POV | Currency | 2 | 3 | rules write Local |

## Where data lives
- Scenario/Version/Currency/View the rules write to, and the source of that claim.

## Traps found
- Shared members, duplicate aliases across dimensions, numeric member names,
  members whose alias equals another member's name, attribute dimensions.

## Key hierarchies
- Entity: Total_Entity > regions > countries > legal entities (level 0).
```

Mark each line with how you know it (outline query, dimension sheet, user said, rule
code). Unverified assumptions go in a separate "to confirm" list.

## 8. From profile to grid

0. Settle the member of **every** dimension before building, and know why each one is
   there: from what the user asked, from the source they gave (a ticket, a report, a file),
   or from the cube profile. Never fill a dimension by habit. If a member cannot be
   justified, ask. Typical traps: a month after today has no Actual (it lives in Forecast
   or Budget); version, currency and view decide which number you get; a dimension left
   at its root may have no data.
1. List every dimension of the cube and give each exactly **one** axis: rows, columns
   or a header row. Prefer header rows over the POV for scenario, version, currency,
   view and year.
2. Pick members from the profile, not from memory. Validate each one exists and is in
   the dimension you think (`FindMember`).
3. Choose the level on purpose: level-0 members for reconciliation with a source file,
   parents for review. Remember dynamic calc parents are computed on retrieval.
4. Use member **names** with the alias table set to None, unless the user works with
   aliases and they are unique.
5. Numeric member names (years, account codes, entity codes) must be entered as text
   (prefix with an apostrophe) or Smart View reads them as data.
6. Write the grid on the scratch sheet (`WriteTsv`), keep navigate without data on,
   `Retrieve`, check `rc` and read the sheet back. When the layout is right and the user
   wants numbers, turn navigate without data off and refresh.
7. Sanity-check before reporting: zeros or blanks where a balance should exist mean a
   wrong intersection until proven otherwise (scenario, version, currency, future month,
   a POV dimension left at a root without data). Fix and refresh before saying it works.
8. Hand the user either the scratch sheet or the exact layout for their own workbook,
   following [adhoc-design.md](adhoc-design.md).

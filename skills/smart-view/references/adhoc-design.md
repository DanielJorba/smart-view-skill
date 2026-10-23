# Designing ad hoc sheets that refresh the first time

An ad hoc refresh works by parsing the contiguous block of cells as a grid: member
labels on the left and on top, data in the body. Every design rule below follows from
that. If Smart View cannot read a cell as a member, a data cell or an empty cell, the
refresh fails, misreads the grid or wipes the cell.

## The rule

**Inside the grid there are only members, each one in its place.** No titles, no
dimension labels, no notes, no formulas, no check columns. Anything else goes on
another sheet that reads from the grid sheet, which is never refreshed itself.

## Layout that works

```
        A          B           C        D        E     ...
  1                                     Final    Final         <- Version   (one header row per column dimension,
  2                                     Actual   Actual        <- Scenario   member repeated over every data column)
  3                                     Local    Local         <- Currency
  4                                     FY26     FY26          <- Year
  5                                     Jan      Feb           <- Period
  6   E100       Sales_7000  Product_X  1.234    1.410
  7   E100       Sales_7001  Product_X  ...
     ^ row dimensions, one column each   ^ data starts in the first column after them
```

The `<- ...` notes are for the reader of this file only. Never write them in the sheet.

- **Column header rows:** one row per column dimension. Repeating the member over every
  data column is the safe choice: it always parses, even after inserting columns.
- **Top-left corner empty:** the cells above the row-dimension columns and left of the
  column headers (A1:C5 above) stay blank. Typing "Entity" or "Account" there is the
  most common reason a hand-built grid refuses to refresh.
- **Row dimensions:** one column per dimension, data from the next column on.
- **Put the "POV" dimensions in the grid as header rows** (Scenario, Version, Currency,
  View...) instead of leaving them on the sheet POV. They stay visible and survive
  copying the sheet, and they do not depend on the POV default, which on a new sheet
  is the dimension root and usually has no data.
- **Each dimension on exactly one axis.** A dimension can be on rows, on columns or on
  the POV, never in two places. Otherwise Smart View answers `Member X is out of place`.
  This also bites when a member you put on rows belongs to a dimension you already
  placed in the header, for example a KPI-type member that actually lives in the
  Market dimension.

## Member names and alias tables

- Labels follow the sheet's alias table, **dimension names included**: with `Default` an
  application can show "Producto" or a company's long name where the member name is
  `Product` or a short code. To work with names, set the sheet to `None` with
  `sv.ps1 call SetAliasTable "" <sheet> None` (sheet level: the user's connections are
  not touched). `member-data.ps1` does it on its own.
- Default grids differ by application: one or several dimensions on each axis. Read the
  grid before pivoting instead of assuming its shape.

- When you type member **names** (not aliases), set the connection alias table to
  **None** first. With the Default alias table on, a typed name can resolve to a
  different member whose alias happens to match, and the row silently returns another
  intersection (field report: a measure resolved to a look-alike member and showed
  zeros in every scenario).
- *Tested on Planning (21-oct-2026):* a connection with an alias table such as
  `Code - Alias` (aliases like `CODE - Description`) rejects a typed description on
  refresh (rc 1000). Check the alias table with `SheetInfo` and switch it
  (`SetAliasTable`) to the table whose names you are typing.
- In VBA: `HypSetAliasTable(sheet, "None")` for the sheet, or
  `HypSetConnAliasTable(friendlyName, "None")` for the connection. The exact name of the
  "no alias" table can vary by provider; check the alias list in the Smart View panel.

## Numeric member names

Member names that look like numbers (years such as `2026`, account or entity codes such
as `7000` or `001`) must be typed as **text** (a leading apostrophe, or the cell
formatted as text before typing). Otherwise Excel stores a number and Smart View may
read it as data instead of a member label, or drop leading zeros *(field practice;
confirm with your provider)*.

## Know the outline first

Build grids from the real outline, not from memory: dimension list, axis per dimension,
exact member names. See [structure-discovery.md](structure-discovery.md).

## Currency and other "where does the rule write" dimensions

If data is calculated by business rules, put on the grid the members **where the rule
writes**: the rule's `FIX` statement tells you which Version, Currency, View or data
type it targets. A typical symptom of getting this wrong: quantities show up but
amounts in currency come back as zero, because Currency was left at the root or at
the reporting currency while the rule writes local currency.

## Checks, differences and notes

- Best: put checks on a separate sheet that references the grid cells.
- Acceptable: outside the grid, separated by at least one fully blank row or column.
  Zoom, pivot and keep-only operations can expand the grid into that space, so keep a
  generous margin.
- Smart View has an option to preserve formulas and comments in ad hoc operations
  (`HSV_PRESERVE_FORMULA_COMMENT` in the option enum). Test it on a copy before relying
  on it; behaviour varies with the operation and the provider.

## Saving

Smart View stores the sheet's connection and POV as sheet metadata when you refresh
and then save the workbook. A grid that refreshed but was never saved afterwards comes
back without its POV on reopening.

## Diagnosing a new ad hoc sheet that returns zeros or blanks

1. Look at the POV toolbar of the failing sheet. If it shows dimension names (the
   roots) instead of concrete members, that is the cause. Compare with a working sheet
   in the same workbook.
2. Refresh another sheet that works, **with the same connection**. If it returns data,
   the connection and the year are fine and the problem is the POV or the grid.
3. With the POV fixed and still zeros: confirm the data actually exists. Has the rule
   that fills those accounts been run for that scenario and period?
4. Check the alias table issue above if only some rows are zero.

## Formatted reports (EPM Cloud Reports) in Excel

- Insert them from the Smart View panel with **Insert Formatted Report**, update them
  with **Refresh Report**.
- **Edit POV creates a new tab and deletes the old one.** Anything on that tab, and any
  formula on another sheet pointing to it, is lost. Keep checks on a different sheet
  and use one workbook per POV (for example one per year).
- Changing the POV of these reports from VBA has no direct API: see
  [automation.md](automation.md).

## Building the workbook from outside Excel

If a script writes the xlsx file (openpyxl, editing the XML), Excel must not have the
file open; otherwise write into the running instance (COM) or create a new file. Never
rebuild a workbook that a user has open with unsaved Smart View changes.

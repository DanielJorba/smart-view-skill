# Troubleshooting Smart View

Field-tested fixes, ordered by how often each cause shows up. Each entry: symptom,
cause, fix. Several are reports from real support cases rather than Oracle
documentation; those are marked *(field report)*.

## Contents

1. First questions for any vague ticket
2. Empty response from provider, then Invalid MRU item
3. `#NAME?` in cells with HsGetValue
4. A new ad hoc sheet returns zeros or blanks
5. `Member X is out of place`
6. Refresh deletes notes, formulas or checks
7. Refresh loses the Excel formatting
8. Timeouts on large refresh or submit; a rule looks stuck "In Progress"
9. A user does not see data they should have access to
10. A report or form does not open from a task list
11. Generic "Smart View does not work", session errors, empty panel
12. Titles scrambled when exporting a report workbook to PDF
13. Things that look like Smart View problems but are not
14. Collecting evidence

## 1. First questions for any vague ticket

Ask before diagnosing; most tickets arrive as "Smart View doesn't work":

- What exactly were you doing, and what is the exact error text? (A screenshot.)
- Which connection and which workbook or form? Does it happen with everything or
  with one artifact?
- Does it work in the **browser** with the same user? If not, it is permissions or the
  service, not Excel.
- Does it work for a colleague? Did it work before, and what changed (security,
  metadata, an update, a new laptop, VPN)?
- Which Smart View version (Help > About) and which Office bitness?

From a screenshot, extract the full POV before anything else: entity, scenario,
version, period, currency, the other dimensions (total or level 0) and whether the
connection is to the input cube or to a reporting (ASO) cube.

## 2. Empty response from provider, then Invalid MRU item

**Symptom.** Opening a connection or a recent item shows *Empty response from provider*
followed by *Invalid MRU item* (offering to remove the entry). Reinstalling Smart View
changes nothing.

**Cause.** Both messages are one problem: the server returns nothing (cause) and Smart
View marks the recent item as invalid (consequence). Typical reasons:
- The connection points to the instance **base URL** instead of the Smart View
  endpoint. For EPM Cloud Planning-based services (Planning, FCCS, Tax Reporting) the
  URL ends in `/HyperionPlanning/SmartView`.
- The user lost access to the application or artifact (recent security change).
- Smart View settings (connections, recent items, options) live in the user registry,
  `HKCU\Software\Hyperion Solutions\HyperionSmartView`, and **survive uninstalling**.
  That is why reinstalling does not help.

**Fix, in order.**
1. Sign in to EPM in the browser with the same user. If that fails, it is access or
   service, not Excel.
2. Create a **private connection by hand** with the URL ending in
   `/HyperionPlanning/SmartView`, instead of reusing the recent item.
3. Clear recent items: Smart View Options > Advanced > number of recently used items
   to 0, OK, then set it back.
4. Still failing: close Excel completely (no `EXCEL.EXE` left in Task Manager),
   **export the registry key as a backup**, delete it, reopen Excel. Smart View recreates
   it clean. The user loses saved connections and options, so keep the export.
5. If reinstalling, download the installer from the EPM instance itself (Downloads
   page), not an old copy from a network share.

### Related: "Unable to connect to the provider" / "The server name could not be resolved"

- *Server name could not be resolved*: the host does not exist or is misspelled (or
  DNS/VPN). Check the host against the instance URL the user opens in the browser.
- *Unable to connect to the provider* (with a Details button): the host answers but not
  as a Smart View provider, usually a URL without `/HyperionPlanning/SmartView`, or a
  proxy blocking it. The skill's scripts complete the path for `*.oraclecloud.com` hosts.
- Smart View may repeat the message several times; the skill's dialog watcher reads and
  closes them and reports the text.

## 3. `#NAME?` in cells with HsGetValue

**Cause** (by frequency): Smart View not loaded or its function add-in path broken;
the connection name used as first argument does not exist on this machine; wrong
alias table.

**Fix.**
1. Check that Smart View is loaded (ribbon tab present; File > Options > Add-ins >
   COM Add-ins: *Oracle Smart View for Office*). If the add-in is disabled, re-enable
   it from *Disabled Items*.
2. Repair the function links: Smart View ribbon > Functions > **Fix Links**. Excel
   stores the full path of the functions add-in (`HsTbar.xla`, in the Smart View `bin`
   folder) in every workbook, so a workbook built on a PC with Smart View in another
   folder breaks elsewhere. If Fix Links is not enough: File > Options > Add-ins >
   Manage *Excel Add-ins* > untick the Smart View functions add-in and add it again
   from the local installation folder (by default `C:\Oracle\SmartView\bin`).
3. The first argument of `HsGetValue` is a **connection name** that must exist on the
   user's machine. Create the private connection with exactly that name (Smart View
   panel > right-click the cube > Add to private connections / Rename), using the same
   text in name and description.
4. Set the alias table the workbook expects (right-click the cube > Set Alias Table).
5. Formulas sometimes read hidden columns: unhide them to see what they expect.
6. After fixing, use Smart View **Refresh**; a plain Excel recalculation does not fetch.

See [cell-functions.md](cell-functions.md).

## 4. A new ad hoc sheet returns zeros or blanks

Almost always the sheet POV points to dimension roots. See the diagnosis steps in
[adhoc-design.md](adhoc-design.md#diagnosing-a-new-ad-hoc-sheet-that-returns-zeros-or-blanks).
Also check that the data exists (rule run?) and the alias-table trap.

## 5. `Member X is out of place`

A dimension appears on two axes, or the member belongs to a dimension already placed
elsewhere (often in a header row or the POV). Move every member of a dimension to a
single axis. See [adhoc-design.md](adhoc-design.md).

## 6. Refresh deletes notes, formulas or checks

The ad hoc refresh treats the whole contiguous block as grid. Keep only members and
data in it; move checks to another sheet or behind a blank row and column. Optional:
the "preserve formulas and comments in ad hoc operations" option, tested on a copy.

## 7. Refresh loses the Excel formatting

Smart View Options > Formatting > enable **Use Excel Formatting** and save as default
options (in VBA, option index `HSV_EXCEL_FORMATTING`). Related options: *Retain numeric
formatting*, *Move formatting on operations*, *Adjust column width*.

## 8. Timeouts on large refresh or submit; a rule looks stuck "In Progress"

**Fix.**
1. Run Smart View Health Check (Help > Health Check) and apply the recommended values.
2. Raise the network timeout in Smart View options (support teams commonly use around
   900 seconds) and keep connections alive if the option is offered *(field report)*.
3. If it persists, Windows WinINet timeouts: under
   `HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings` the DWORD values
   `ReceiveTimeout` / `SendTimeout` (milliseconds) can be raised, for example to
   480000. Changing the registry needs the user's or IT's approval; export the key first.

A rule that Smart View shows as "In Progress" may have finished on the server: the
client timed out, the job did not. Check the job status in the web interface (Jobs /
Job Console) before relaunching it, or you may run it twice.

## 9. A user does not see data they should have access to

1. Security: in Access Control, check the user's groups and the access granted on the
   members involved (read/write by dimension).
2. **User variables** without a value: forms and reports that use user variables return
   nothing until each variable has a member. Smart View panel > right-click the
   instance > User Preferences > User Variables.
3. If it is a new ad hoc sheet, check the POV first (entry 4).
4. A permission loss can also show as *Empty response from provider* (entry 2).

## 10. A report or form does not open from a task list

Check in order: a connection to the instance exists (private or shared); the cube
connections the report uses exist with the **friendly name** it expects; user
variables are set; the alias table matches what the report expects (User Preferences >
Application Settings > Alias Table).

## 11. Generic "Smart View does not work", session errors, empty panel

1. Health Check (Help > Health Check), apply recommended values. The settings that fail
   most often are the browser emulation mode and the embedded browser component
   (WebView2) used for sign-in *(field report)*.
2. Session or token errors with the panel already empty or disconnected: reconnect
   from the private connection (the arrow next to it). Disconnect All is not needed.
   It is fixed when the tree appears.
3. Add-in disabled or broken install: it is the desktop team's job (software center,
   admin rights), not an EPM issue.

## 12. Titles scrambled when exporting a report workbook to PDF

Report titles are vector EMF pictures regenerated on each refresh, and Excel's PDF
exporter mishandles their character spacing, even at 100% scale. Rasterize titles
before export and restore them after: see [automation.md](automation.md#pdf-export).

## 13. Things that look like Smart View problems but are not

- **Refresh Database / cube refresh very slow** after changing a dense dimension
  (typically Account or Period in a block storage cube): restructuring dense
  dimensions rewrites blocks and is slow by nature. Do not kill or relaunch it.
- **A total does not add up** in the grid: look at the outline (consolidation operators
  such as ignore `~`, dynamic calc members, time balance), not at Smart View.
- **Zeros where a rule should have written:** check the rule's job status and its FIX
  before blaming the sheet.

## 14. Collecting evidence

- Smart View can log messages to a file (Options > Advanced; option indices
  `HSV_LOGMESSAGE_DISPLAY`, `HSV_ROUTE_LOGMESSAGE_TO_FILE`,
  `HSV_ROUTE_LOGMESSAGE_FILE_LOCATION`). Turn it on, reproduce, attach the log.
- Health Check output, Smart View version and Office version and bitness.
- In VBA, `HypGetLastError` right after the failing call gives code, message and
  description.

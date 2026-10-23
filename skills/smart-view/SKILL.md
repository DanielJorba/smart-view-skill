---
name: smart-view
description: Operate and master Oracle Smart View for Office, the Excel add-in for Oracle EPM Cloud (Planning, FCCS, Tax Reporting), Essbase and Hyperion. Connects to the user's running Excel through a read-only bridge to sign in once, refresh, open forms and ad hoc grids, zoom, pivot, set the POV, search members, read hierarchies, attributes and member properties, and map a cube's dimension structure before building an ad hoc grid that refreshes the first time. Also writes and debugs VBA macros with the Smart View API (HypConnect, HypRetrieve, HypSetPOV...), audits, builds and fixes HsGetValue sheets, troubleshoots errors (Empty response from provider, Invalid MRU, #NAME?, timeouts, zeros, out-of-place members) and checks Smart View updates. Use it whenever the user mentions Smart View, ad hoc grids, POV, EPM, Planning, FCCS or Essbase data in Excel, Hyp* or Hs* functions, smartview.bas, or wants to explore a cube's dimensions or members, even if they do not name the add-in.
compatibility: Knowledge parts work anywhere. Operating Smart View needs Claude Code running on the same Windows PC as desktop Excel 64-bit with Smart View and PowerShell 5.1; the SmartViewBridge.xlam add-in ships prebuilt in bridge/ and installs with scripts/sv.ps1 install.
---

# Oracle Smart View for Office

Smart View is an Excel add-in that talks to an EPM or Essbase server. Everything it
shows comes from a **connection**, every Smart View sheet is **associated** with one
connection, and every grid is read against a **POV** (point of view) for the
dimensions that are not on rows or columns. Keep that model in mind: most problems and
most good designs follow from it.

This skill does two kinds of work:

1. **Operate Smart View** in the user's own Excel through the bridge
   (`scripts/sv.ps1` + `SmartViewBridge.xlam`): connect, refresh, open, explore and read.
2. **Know Smart View**: design ad hoc grids, write VBA, explain errors, plan updates.

## Rules that protect the user

These exist because Smart View usually points at production systems, where a mistake
is immediate and shared by everyone.

- **Never change cube data or server artifacts.** No submit, no calculate, no business
  rules, no calc scripts, no consolidate or translate, no copy version, no approvals,
  no substitution-variable changes, no cell text or attachments, no metadata submit
  from a dimension sheet. The bridge does not wrap any of that and refuses those menu
  commands. If the user wants one of these, explain exactly what to click and let
  them do it. In VBA you write for the user, put such calls in a clearly named,
  separate procedure and say what it does.
- **No `HsSetValue` or `HsSetText`.** They write to the server on Submit. The bridge
  refuses to write them; in workbooks you design for the user, flag them clearly.
- **Prepare freely, never execute.** Claude may edit sheets to prepare a change: a
  dimension sheet (new parent, alias, storage, moving members to `_Hist`), an ad hoc
  grid, values on a sheet for the user to review. It lists every cell it changed and
  leaves the execution to the user: Submit (data or metadata), Refresh Database,
  rules and calculations are never run by Claude.
- **Dimension changes: before and after, side by side.** When preparing a metadata change,
  keep a sheet with the members as they are now and another with the proposed ones, so
  the user sees exactly what changes, unless the user says otherwise.
- **Changes to data cells are preparation, so say so loudly.** A later Submit by the user
  sends whatever is there: list the changed cells and never leave unannounced edits. Build and explore in the
  bridge's **scratch workbook**; touch a user workbook only to read it or when the
  user explicitly asks for a layout change there.
- **Never touch the user's existing private connections.** Do not modify, rename,
  remove or change the alias table of a connection the user already has: their
  `HsGetValue` formulas and macros depend on its exact name and settings. Create new
  connections only, with a new name, after `ConnectionExists` says it is free. Change
  alias tables per sheet (`SetAliasTable`), never per connection.
- **Passwords are never yours.** Do not ask for them, accept them, store them or put
  them in code. The user signs in once in Smart View's own dialog; after that the
  session is reused by every sheet connected to the same server.
- **Figures stay hidden.** Every read through the bridge returns numbers as `<n>`. Work
  with structure, members, error texts and checks computed in Excel that return text
  (OK / DIF). Never try to read amounts any other way (Excel COM, files, screenshots).
  Only if the user has set `SMARTVIEW_SKILL_ALLOW_FIGURES=1` themselves and explicitly
  asks to see amounts in that message, add `-ShowFigures` to that call. Never set or
  change that variable.
- **Structure first, data only on request.** Exploring dimensions and building grids
  works with *navigate without data* on, so no figures leave the system. Retrieve
  actual numbers only when the user asks for them, and remember they may be
  confidential client data.
- **Do not invent API functions or options.** The declared functions, return codes
  and option indices are listed in `references/vba-api.md`. If a behaviour is not
  confirmed there, say so and test it on the scratch workbook.

## Route by task

| The user wants to... | Read |
|---|---|
| Connect, refresh, open forms or ad hoc grids, zoom, pivot, POV, read a sheet | `references/operations.md` (+ `references/bridge-setup.md` the first time) |
| Understand a cube: dimensions, hierarchies, members, properties, attributes | `references/structure-discovery.md` **(read before building any ad hoc grid)**; `scripts/panel.ps1` opens dimension sheets from the panel |
| Build an ad hoc grid that refreshes the first time | `references/adhoc-design.md` |
| Prepare a metadata change (move a member, `_Hist` history, alias) | `references/metadata-changes.md` |
| Does a member have data? (before moving it) | `scripts/member-data.ps1` (see `references/metadata-changes.md`) |
| Which cubes does an application have? | `scripts/panel.ps1 cubes <APP> <url>` (leaves the panel's Home page and expands the server on its own; if it still fails, ask the user to click Private Connections once) |
| Open a report from the panel (Reports folder) | `scripts/panel.ps1 open-report <APP> "Reports/<folder>/<report>" <url>` (see `references/operations.md`) |
| **First time in an environment: which cube to use** (ask once, save, reuse) | `references/connections.md` > "The cube to work on"; `sv.ps1 private-connections` |
| Save connection links to each environment, sign in once | `references/connections.md` |
| Write or fix a VBA macro using Smart View | `references/vba-api.md` |
| Audit, build, refresh or fix HsGetValue sheets and other cell functions | `references/cell-functions.md` |
| Explain an error or odd behaviour | `references/troubleshooting.md` |
| Automate beyond the API (dialogs, scheduled runs, PDF export, Excel COM) | `references/automation.md` |
| Know the installed version, update Smart View or its extensions | `references/updates.md` |

## Bridge quick start

All commands print JSON (except `read`, which prints TSV). Run from the skill folder:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 status
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 call Scratch Explore true
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 call UIConnect "" Explore "MyApp_Plan1"
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 call Retrieve "" Explore
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 read "" Explore
```

- An empty workbook argument (`""`) means the scratch workbook. Pass a workbook name
  only to read a user's workbook or when the user asked you to work there. Calling
  `powershell -File` from inside Windows PowerShell 5.1 drops empty `""` arguments and
  shifts the rest: from a PowerShell prompt call the script directly
  (`& .\scripts\sv.ps1 call Retrieve '' Explore`); from Bash or cmd the commands above
  work as written.
- Check `ok` and `rc` on every call. When `ok` is false, `error` carries Smart View's
  own message (from `HypGetLastError`); read it before retrying anything.
- **Smart View URL of an EPM Cloud instance = instance host + `/HyperionPlanning/SmartView`**
  (`https://<host>/HyperionPlanning/SmartView`). The base URL, `/epmcloud` or the browser
  URL fail with "Unable to connect to the provider" or "Empty response from provider".
  The scripts fix an `*.oraclecloud.com` URL on their own; write it right anyway in
  everything you hand the user (profiles, macros, `HsGetValue` connection strings).
- To connect when you know URL, application and cube but no private connection name,
  use the inline string `"<url>|<url host>|<APP>|<CUBE>"` as the connection argument
  (tested on EPM Cloud). Do not create private connections and never touch the user's.
- Connections made with the inline string do not show in the Smart View panel. After
  connecting, also run `panel.ps1 server <url>` so the user sees the server and its
  applications in the panel (it opens the panel and leaves its Home page on its own).
- EPM Cloud connects with `UIConnect` (Smart View shows its own sign-in the first
  time); `Connect` is for on-premises sources only.
- `status` first. If `bridgeLoaded` is false, ask the user and run `sv.ps1 install`
  (prebuilt add-in, nothing to build; details in `references/bridge-setup.md`).
  If `excelRunning` is false, say so and run `sv.ps1 start-excel`: it launches Excel as
  the user would (visible, Smart View loaded). Never create Excel through COM
  (`New-Object -ComObject Excel.Application`): that instance is hidden and Smart View
  does not load in it.
- Workbooks: `sv.ps1 open <path>` opens a saved workbook **read-only** (add `edit` only
  when the user wants to work in it); `sv.ps1 new-workbook` adds an empty one for the
  user. The skill never saves workbooks, and closes only the ones it opened or created
  (`sv.ps1 close <workbook>`, without saving; it refuses any other): the user closes theirs. Opening a workbook with
  Smart View sheets may bring up the sign-in; the dialogs are reported like in `call`.
- Pop-ups are handled in real time: every `call` runs with a dialog watcher and returns
  what appeared under `dialogs` (text, options, what was done). Known messages are
  dismissed after reading; the "worksheet contains data" question is answered with
  `-Reuse` (default cancel; `contents` keeps the user's members when reconnecting a
  sheet; `clear` only on the scratch workbook); sign-in windows and server-changing
  confirmations are never touched, only brought to the front (or flashed in the taskbar
  when Windows refuses) so the user sees them. Report every dialog to the user.
- Keep the scratch workbook tidy: reuse one working sheet per purpose instead of creating
  a new one per step, never recreate a sheet other sheets' formulas point to, and at the
  end of a task remove the working sheets with `sv.ps1 drop-sheet <sheet> [...]` (scratch
  workbook only) and leave only what the user needs to see.
- **Results the user keeps are connected Smart View sheets**, never values pasted into
  an unconnected sheet: a member list (for example a dimension down to level 0) is an ad
  hoc grid built and zoomed on a connected sheet (`ZoomIn` level 2 for level 0), or a
  dimension sheet opened from the panel. Give that sheet a clear name, leave it
  connected so the user can refresh, zoom and pivot, and delete the working sheets
  around it. A plain copy only when the user asks for one (and say it is not connected).
- With figures hidden, count "has data" with an Excel formula that returns text (for
  example `="nonzero="&SUMPRODUCT(ISNUMBER(B7:B169)*(B7:B169<>0))`, written with
  `WriteTsv` in English syntax and read back with `read` on that one cell): masked reads
  turn zeros into `<n>` too, so they cannot tell an empty intersection from a filled one.
- Excel is shared with the user. Tell them before you start, do not close their
  workbooks, and expect the active sheet to change while you work.

## What has no API

Some Smart View features exist only in the UI: task lists, reports with non-default
prompt values (opening a report with its defaults is automated: `panel.ps1 open-report`,
see `references/operations.md`), the dimension editor of the EPM Admin Extension, approvals and
the Member Selection dialog. For those, tell the user the exact clicks, wait for them
to do it, then read the resulting sheet with `sv.ps1 read`. `references/operations.md`
lists what is API, what is UI-only and what is off limits.

## Answering without the bridge

When the user wants code, a design or an explanation rather than an action:

- **VBA:** a complete module with `Option Explicit`, constants for connection and sheet
  names at the top, every `Hyp*` return code checked, errors read with
  `HypGetLastError`, no `MsgBox` inside loops that may run unattended. State the
  prerequisites: `smartview.bas` imported into the workbook, the connection saved,
  the user signed in. Recommend testing on a copy.
- **Ad hoc layouts:** give the exact cells, following `references/adhoc-design.md`.
- **Errors:** most likely cause first, then the ordered fix, then how to confirm.

Reply in the user's language.

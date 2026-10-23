# smart-view: an Agent Skill for Oracle Smart View for Office

A skill that lets Claude operate and master **Oracle Smart View for Office**, the Excel
add-in for Oracle EPM Cloud (Planning, FCCS, Tax Reporting), Essbase and Hyperion.

- **Operate Smart View in your own Excel** through a read-only bridge: sign in once,
  refresh sheets, open forms and ad hoc grids, zoom, pivot, change the POV, read any
  Smart View sheet.
- **Understand a cube before touching it**: list its dimensions, walk hierarchies
  without retrieving data, read member properties (Essbase), and keep a cube profile
  so ad hoc grids are built from the real outline and refresh the first time.
- **Keep HsGetValue sheets under control**: audit which connection, dimensions and
  members each formula really uses, find error cells and broken links, refresh,
  build function grids and spot-check single values.
- **Know Smart View**: VBA API reference checked against Oracle's documentation and
  the installed `smartview.bas`, cell functions (`HsGetValue`...), ad hoc design
  rules, troubleshooting of the classic errors, and version and update guidance.

## Safety by design

- Nothing changes cube data or server artifacts. Submit, calculate, business rules,
  copy version, approvals and similar are not wrapped, and the menu wrapper refuses
  them in six languages.
- No password is ever requested, stored or passed. You sign in in Smart View's own
  window.
- Exploration runs with *navigate without data*, so structure work brings back no
  figures. Data is retrieved only when you ask.
- The bridge works on its own scratch workbook unless you name one of yours.
- Figures are hidden (`<n>`) in everything Claude reads. To let Claude see amounts on
  your own machine, set the user environment variable `SMARTVIEW_SKILL_ALLOW_FIGURES=1`
  yourself; Claude then shows them only when you explicitly ask in that message.

## Requirements

- **Full features (operating Smart View):** Claude Code (CLI or the desktop app's Code
  tab) on Windows, desktop Excel 64-bit with Smart View installed, PowerShell 5.1.
  Claude must run on the same PC as Excel: skills uploaded to the Claude apps run in a
  cloud sandbox and cannot reach your Excel.
- **Knowledge only (VBA, HsGetValue, troubleshooting, ad hoc design):** any Claude client
  with skills.

## Install

Claude Code:

```text
/plugin marketplace add DanielJorba/smart-view-skill
/plugin install smart-view@smart-view-skill
```

First use: ask Claude anything about Smart View. It checks Excel and, if the bridge is
missing, asks in the chat before running `scripts/sv.ps1 install`, which copies the
prebuilt add-in to your Excel add-ins folder, loads it, creates the figures switch at 0
(figures hidden) and an empty local profiles file. Nothing else to configure. You still:

- sign in to each EPM instance in Smart View's own window (Claude never sees passwords);
- answer, the first time on each environment, which cube to work on (Claude saves it in
  `%LOCALAPPDATA%\smart-view-skill\connections.json`, on your PC only);
- set `SMARTVIEW_SKILL_ALLOW_FIGURES=1` yourself only if you want Claude to see amounts.

Corporate PCs may block PowerShell scripts or Office add-ins by policy; then ask your IT
team. Excel and Smart View in English or Spanish are tested; other languages may miss
some dialog texts.

Claude apps (knowledge only): zip the `skills/smart-view` folder and upload it in
Customize > Skills.

## Layout

```text
skills/smart-view/
  SKILL.md                 entry point
  references/              one file per topic, loaded on demand
  bridge/SmartViewBridge.bas   read-only VBA wrappers returning JSON (source)
  bridge/SmartViewBridge.xlam  the same, prebuilt as an Excel add-in
  scripts/sv.ps1           CLI the agent uses to call the bridge
  scripts/build-bridge.ps1 maintainers: rebuild the add-in after changing the .bas
```

## Disclaimer

Not affiliated with or endorsed by Oracle. Oracle, Essbase, Hyperion and Smart View
are trademarks of Oracle and/or its affiliates. This project does not include or
redistribute Oracle files: the bridge declares the Smart View DLL functions it uses
itself, from the published signatures, for interoperability.

## License

MIT, see [LICENSE](LICENSE). Provided as is, without warranty: test on a non-production
instance first.

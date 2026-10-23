# Bridge setup

The bridge has two parts, both shipped with the skill:

- `bridge/SmartViewBridge.xlam`: a prebuilt Excel add-in with read-only VBA wrappers
  (`SVB_*`) over the Smart View API that return JSON. Its source is
  `bridge/SmartViewBridge.bas`, which declares the Smart View DLL functions itself, so
  no Oracle file is needed or redistributed.
- `scripts/sv.ps1`: attaches to the user's running Excel and calls those wrappers
  through `Application.Run`.

Why an add-in: Smart View functions run inside Excel's process, so they can only be
called from VBA in that same Excel.

## Install (every user, once)

1. `sv.ps1 status`. If `excelRunning` is false, ask the user to open Excel.
   If `smartViewLoaded` is false, Smart View is disabled: the user re-enables it in
   File > Options > Add-ins > COM Add-ins (or *Disabled Items*).
2. If `bridgeLoaded` is false, **ask the user** before installing: it copies the
   add-in to their AddIns folder and loads it in Excel, which is a persistent Excel
   setting. On a yes:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 install
   ```

   It copies `bridge/SmartViewBridge.xlam` to `%APPDATA%\Microsoft\AddIns\` (a
   location Excel trusts by default), removes the downloaded-file mark, and ticks it
   in Excel's add-in list. The only other change is the figures switch below (created at
   0 if missing). No Trust Center setting, no registry edit by hand.
3. `sv.ps1 status` again: `bridgeLoaded: true` and `bridgePing.ok: true`.
4. Tell the user how figures are set (`figures` in the output). The install creates the
   user environment variable `SMARTVIEW_SKILL_ALLOW_FIGURES` with value 0 if it does not
   exist (amounts hidden) and never overwrites it. To let Claude see amounts, the user
   changes it to 1 themselves; Claude never changes it.

Nobody has to import modules or open the VBA editor. If `install` reports that the
prebuilt add-in is missing, the skill copy is incomplete: get the full repository.

## Updating the bridge

When a new skill version ships a new `SmartViewBridge.xlam`, run `sv.ps1 install` again
(it unloads the old copy, replaces it and loads the new one). Smart View updates do
not require rebuilding the bridge: it calls the DLL entry points directly, whose
signatures are stable across releases.

## Uninstall

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 uninstall
```

Unloads the add-in and deletes the installed copy. Optional local data:
`%LOCALAPPDATA%\smart-view-skill\` (connection profiles and cube profiles).

## Maintainers: rebuilding the add-in after changing the .bas

Only needed when `bridge/SmartViewBridge.bas` changes. Then commit the new `.xlam`.

- Automatic: `scripts/build-bridge.ps1`, if the maintainer's Excel already trusts
  access to the VBA project object model. The script never changes that setting.
- Manual (2 minutes, no setting change):
  1. In Excel, create a new blank workbook.
  2. Alt+F11. In the VBA editor: File > Import File... > `bridge/SmartViewBridge.bas`.
  3. Debug > Compile VBAProject. It should finish silently.
  4. File > Save As > *Excel Add-in (\*.xlam)* > `SmartViewBridge.xlam`, saved in the
     skill's `bridge/` folder (not in the AddIns folder Excel proposes).
  5. Close the blank workbook without saving, then `sv.ps1 install`.

The add-in is compiled VBA plus its source: anyone can open it in the VBA editor and
compare it with `SmartViewBridge.bas`.

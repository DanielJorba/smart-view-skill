# Smart View versions and updates

## What is installed

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/sv.ps1 version
```

Returns the product version of the Smart View DLL (`HsAddin.dll`), its folder, the
path of `smartview.bas` and the installed extensions (for example
`Oracle.SmartView.EPMAdminExtension`). In Excel: Smart View ribbon > Help > About.
Also note the Office version and bitness (File > Account > About Excel), which matter
for the installer and for VBA declarations.

## What is current

Do not state the latest version from memory: Smart View ships several releases a year.

- Search the web for the current *Oracle Smart View for Office* release and its
  *Readme / New Features* page on docs.oracle.com.
- For EPM Cloud, the instance itself offers the installer it supports: user menu >
  **Downloads** > Smart View. That version is the safe target for that environment.
- Read the readme's *supported platforms* and *known issues* before recommending an
  upgrade, especially the Office versions supported and any change in sign-in
  (embedded browser components).

Compare the installed version with that target and report: installed, available,
what the newer release adds that matters to the user, and known issues.

Facts from the readmes that decide whether to update (checked Oct 2026; re-check, they
change with each release):

- **Support window:** Oracle certifies only the **current and the previous** release
  against EPM Cloud. Two releases behind means unsupported, even if it still works.
- **Versions are YY.100 and YY.200** (two a year), plus patch builds. Upgrade paths are
  listed in each readme; jumping several releases is normally allowed but check it.
- **Office:** from 25.200, only Office 2021, 2024 and Microsoft 365; the last release for
  Office 2016/2019 is 25.100. Check the user's Office before recommending an upgrade.
- **WebView2** (Microsoft's embedded browser) is required from 25.200 for the sign-in
  and other UI; configured in `properties.xml` in the Smart View `cfg` folder. Several
  known issues are tied to it (extension updates, Disconnect All).
- Some new features also need a minimum EPM Cloud monthly version (the readme says
  which, for example "EPM 26.06+").

## Updating Smart View (the user runs it)

1. Download the installer from the EPM instance Downloads page or Oracle's official
   download page. Never from an old copy on a network share. Downloading needs the
   user's go-ahead; installing is the user's action.
2. Close **all** Office applications, including Outlook: Smart View loads there too,
   and an open Outlook makes the installer fail or ask for a reboot.
3. Run the installer. Depending on how it was installed (per user or for all users)
   it may need administrator rights; in managed companies it goes through the IT
   software center.
4. Open Excel, check Help > About, run **Health Check** (Help > Health Check) and apply
   the recommended settings.
5. Connections, options and recent items live in the user registry
   (`HKCU\Software\Hyperion Solutions\HyperionSmartView`) and survive the update.
6. The skill's bridge needs no rebuild ([bridge-setup.md](bridge-setup.md#updating-the-bridge)).
   Workbooks with their own macros keep working with their old `smartview.bas`, but
   re-import it if the macros need newer functions.

## Extensions

Extensions (EPM Admin Extension, Narrative Reporting, Task Manager and others) are
installed and updated from Smart View: Options > Extensions, where Smart View can check
for updates, new installs and uninstalls offered by the connected services. Restart
Excel after updating an extension. The EPM Admin Extension is what lets
administrators open and edit dimensions in Excel; see
[structure-discovery.md](structure-discovery.md).

## When an update is the fix

Suspect the version when a feature documented for EPM Cloud is missing, when sign-in
fails after an EPM Cloud monthly update (embedded browser changes), or when Health
Check flags the version. Otherwise follow
[troubleshooting.md](troubleshooting.md) first: reinstalling is rarely the fix,
because the settings survive it.

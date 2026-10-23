# Connections: save each environment once, sign in once

## How Smart View keeps connections

- **Private connections**: saved per user, each with a **friendly name**. Macros,
  cell functions (`HsGetValue`) and the bridge refer to connections by that name.
- **Shared connections**: a list published by the server (or an XML file) that every
  user sees; read-only for the user.
- **Quick connect URLs**: the recently typed URLs, stored in the user registry under
  `HKCU\Software\Hyperion Solutions\HyperionSmartView\QuickConnect`.
- Connections, options and recent items live in the user registry and survive
  reinstalling Smart View.

## Connection URLs

- EPM Cloud Planning-based services (Planning, FCCS, Tax Reporting, Enterprise
  Profitability and others on the same platform):
  `https://<instance-host>/HyperionPlanning/SmartView`. The most common mistake is
  using the base URL of the instance, which answers *Empty response from provider*.
- Other providers (Essbase, Narrative Reporting, on-premises EPM) use their own
  endpoints: take them from the product documentation or from the administrator, and
  verify by connecting once.

## The skill's connection profiles

To let the user "save the link to each cloud", keep one profile per environment in
`%LOCALAPPDATA%\smart-view-skill\connections.json` (outside the skill, so it is never
published). `sv.ps1 install` creates it empty (`[]`) and never overwrites it; Claude adds
a profile the first time it works on an environment. Never store credentials there.

```json
[
  {
    "name": "Client A - Planning TEST",
    "environment": "test",
    "url": "https://<instance-host>/HyperionPlanning/SmartView",
    "provider": "Planning",
    "application": "<APP>",
    "cube": "<PLAN_TYPE>",
    "friendlyName": "ClientA_TEST_<APP>_<PLAN_TYPE>",
    "inline": "https://<instance-host>/HyperionPlanning/SmartView|<server>|<APP>|<PLAN_TYPE>",
    "workCube": "<PLAN_TYPE>",
    "cubeType": "BSO",
    "cubeTypeConfirmed": true,
    "reportingCube": null,
    "getValuesConnection": "<private connection name>",
    "aliasTable": "None",
    "dataCheckPov": "Period#<stored member>;Account#<stored parent>;Currency#<member>",
    "notes": "Answers given by the user on <date>."
  }
]
```

Fields are read by the scripts (`member-data.ps1 -Profile <name>`), so keep the types:
`environment` is `test` or `prod` (ask if the user did not say; do not guess from the
host); `cubeType` is `BSO`, `ASO` or `null`, with doubts in `cubeTypeConfirmed: false`
and `notes`, never in the value; `getValuesConnection` is just the connection name;
`friendlyName` is the private connection to connect with, or `null` when connecting with
the inline string; `reportingCube` and `dataCheckPov` may be `null`. Comments go in
`notes`.

### The cube to work on: ask once, then always use it

Applications often have several cubes (an input BSO cube, a reporting ASO cube, others
that are obsolete), and which one to use is a decision of the team, not something the
panel shows. So, **the first time** you work on an environment:

1. Look for its profile in `connections.json` (`sv.ps1 connections`). If it is there, use
   `workCube`, `getValuesConnection` and `dataCheckPov` without asking again.
2. If it is not: list the cubes (`panel.ps1 cubes <APP> <url>`) and the user's private
   connections (`sv.ps1 private-connections`), and ask the user in one message:
   - which cube to work on (and, if there is one, the reporting cube);
   - whether it is BSO or ASO, if they know (a name ending in RPT is usually ASO; ask
     anyway);
   - which private connection their `HsGetValue` formulas use;
   - for a BSO cube without a reporting cube: stored members to fix in data checks
     (`dataCheckPov`, passed to `member-data.ps1 -Pov`).
3. Save the answers in the profile (that file lives outside the skill and is never
   published) and say so. Change them only when the user says so.

`sv.ps1 private-connections` reads the user's private connections (name, URL,
application, cube, provider, alias table) from the registry without touching them and
never decodes the stored credential. Use it to tell which environment a private
connection points to: a connection named like a production cube may point to a test
instance.

- `environment` is `test` or `prod`. Use test for anything experimental and say it out
  loud before touching prod, even read-only.
- `friendlyName` follows one convention so names are predictable in macros and
  `HsGetValue` formulas: `<Client>_<ENV>_<App>_<Cube>`.
- `sv.ps1 connections` lists the profiles plus the quick-connect URLs Smart View
  already remembers, which is a good way to fill the file the first time.

## Creating the Smart View private connection from a profile

Only create **new** connections. Existing ones belong to the user and their
`HsGetValue` formulas depend on them: never update, rename, remove them or change
their alias table (`HypUpdateConnection`, `HypResetFriendlyName`,
`HypRemoveConnection`, `HypSetConnAliasTable`).

```powershell
sv.ps1 call ConnectionExists "ClientA_TEST_APP_Plan1"
sv.ps1 call CreateConnection "<provider type>" "<url>" "<server>" "<APP>" "<PLAN_TYPE>" "ClientA_TEST_APP_Plan1" "Client A test"
```

*Tested on EPM Cloud Planning (21-oct-2026):* `CreateConnection "Planning" <url> <instance host> <APP> <CUBE> <friendlyName> <description>`
works and the connection then appears in the Smart View panel (Private Connections
drop-down). Connections made with the inline string are not saved and do not appear.

The exact provider type string and server value that `HypCreateConnectionEx` expects
depend on the provider: see `references/vba-api.md`. If creation fails, the user can
create it once from the Smart View panel (Private Connections > type the URL > pick
application and cube > Add to private connections, using the friendly name from the
profile), and everything else works the same.

## Signing in once

1. Associate the first sheet: `sv.ps1 call UIConnect "" <sheet> "<friendlyName>"`.
   For EPM Cloud this is the only documented way (`HypConnect` does not work against
   cloud sources). `<friendlyName>` can be a saved private connection or an inline
   string `"URL|server|app|cube"`, which needs no saved connection at all.
2. Oracle documents that `HypUIConnect` shows Smart View's **own sign-in** for cloud
   sources (often a browser-based or single sign-on page). The command waits while the
   user signs in. Tell the user it is coming and that the password is typed by them,
   in that window. If a session for that server already exists, the sign-in may
   complete on its own: confirm the behaviour in your environment.
   *Tested on EPM Cloud (Smart View 25.200, 21-oct-2026):* the inline string needs a
   non-empty server part (empty gives error 57705, "Argument srv ... is an empty string
   or NULL"), and a wrong application name gives "Application not found". Smart View
   shows these as its own modal dialogs: the call waits until the user closes them,
   and afterwards the call may end in a VBA error. Prefer a **private connection
   created once by the user in the Smart View panel** (they pick application and cube
   from the list, so nothing is guessed) and connect by its friendly name.
   Connecting by a private connection name worked first time and without a new
   sign-in when the user already had a session. `SheetInfo` then returns the real
   server (the instance host), application, cube and provider (`FCCS`): save those in
   the profile.
   Test environments are often reused for other applications: never trust a stored
   application name for a test pod; read it from the panel or `SheetInfo`.
   Pods of the same Oracle identity domain share the sign-in: one login covered every
   pod of the same domain tested (Oct 2026).
3. After that, other sheets and connections to the **same server** connect without
   asking, until Excel closes or the session expires (EPM Cloud sessions time out
   after a period of inactivity).
4. If the call fails with a not-connected or session error, ask the user to connect
   once from the Smart View panel and retry.

Passwords never go through the bridge: the connect wrappers pass an empty password
on purpose and `SVB_ConnectionInfo` discards the password slot.

## Default connection

`HsGetValue` and other cell functions can use the default connection when the
connection argument is empty. The default is set in the Smart View panel (right-click
a connection > Set as Default Connection) or with `HypSetAsDefault`. Prefer explicit
friendly names in workbooks other people will open.

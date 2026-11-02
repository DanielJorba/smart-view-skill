# EPM Cloud through REST (epm.py)

`scripts/epm.py` talks to the REST APIs of Oracle EPM Cloud pods. It can only read:
every request passes a guard that allows GET on the EPM REST paths and a single POST,
`exportdataslice`, which returns a grid and changes nothing. Run it with
`python "<skill dir>/scripts/epm.py" <command> ...`.

The REST layer is the fast one for structure: once a dimension is cached, `tree` and
`find` answer in a second, where Smart View needs a zoom per level. Use Smart View for
the grids the user keeps; open them on the same cloud with `pod:<name>` (below).

Local data lives in `%LOCALAPPDATA%\smart-view-skill\`, next to the Smart View
profiles: `pods.json` (URLs, application, cubes, test or prod; no secrets), `cache\`
(dimension hierarchies) and `logs\calls.log` (one line per call: time, pod, method,
path, HTTP status; no payloads or figures).

Pods registered before the merge (in `%LOCALAPPDATA%\epm-cloud-skill\`) are copied
here on first use, and credentials saved by the former epm-cloud skill keep working.

## Rules that protect the user

Pods hold client data and production pods are shared by everyone who works on them.

- **Read only, always.** Never try to change a pod: no rules, data loads, metadata
  imports, refresh, jobs, file uploads or deletions, by this client or any other way
  (curl, Invoke-RestMethod, EPM Automate, the browser). If the user wants a change,
  prepare it and give the exact steps; the user executes it.
- **Passwords:** see SKILL.md. The user runs `cred-set` in their own terminal.
- **Figures:** the same switch as Smart View, `SMARTVIEW_SKILL_ALLOW_FIGURES` (see
  SKILL.md). Without it `slice` prints `<n>` and refuses `--out`.
- **Say which pod you read.** Each pod is `test` or `prod`. Prefer test; on prod keep
  grids small and avoid heavy reads while the user is closing a period. Do not run
  several large slices in parallel.
- **User names stay hidden.** The client never prints the saved user, and `files`,
  `cat`, `get` and `dijob` replace e-mail addresses (and Oracle's `user_domain_` file
  prefixes) with `<user>`. Do not try to recover them.
- **One client at a time.** Do not carry member names, structures or figures from one
  client's pod into another client's work.
- **Do not invent endpoints or members.** The verified REST calls are in
  `rest-endpoints.md`. Member names come from `tree` and `find`, not from memory.

## Route by task

| The user wants to... | Do |
|---|---|
| Open the pod's cube in Smart View | `sv <pod> [--cube C]` prints the URL and inline connection; the bridge takes `pod:<name>[/<CUBE>]` directly |
| Set up a pod for the first time | `pod-add`, then the user runs `cred-set` (see Setup) |
| Check the connection, list applications | `ping <pod>`, `apps <pod>` |
| Which cubes, BSO or ASO | `cubes <pod>` |
| Which dimensions a cube has | `dims <pod> --cube <cube>` (also caches them for `slice` checks) |
| See a hierarchy or where a member hangs | `tree <pod> <Dim> [Member] --depth N` |
| Search a member by part of its name or alias | `find <pod> <Dim> <text>` |
| Member properties (formula, storage, parent) | `member <pod> <Dim> <Member>` (needs Service Administrator) |
| Read figures, check a tie, does a member have data | `slice` (see Reading data) |
| Substitution variables (current month, year...) | `subvars <pod> [--cube C]` (application level first, then each cube) |
| Which rules or jobs exist; did a job finish | `jobdefs <pod> [--type RULES]`, `job <pod> <id>` |
| Which files or snapshots are in the pod | `files <pod>` |
| Which data integrations or pipelines ran, and what they did | see Data Integration below |
| Read a log or text file of the pod | `cat <pod> outbox/logs/<id>.log` |
| Anything else that is a documented GET | `get <pod> <path>` (paths that look like actions are refused) |

## Setup (once per pod)

1. `pod-add <pod> --url <base url> --app <APP> --cube <CUBE> [--cube <CUBE2>] --env test|prod`.
   The URL can be pasted with any path (for example the Smart View URL); only the base
   is kept. Pick short names such as `client-test`.
2. The user runs `python "<skill dir>/scripts/epm.py" cred-set <pod>` in their own
   terminal and types user and password. It tests the sign-in at the end.
   - OCI pods (`*.ocs.oraclecloud.com`): user without identity domain. Classic pods:
     `identitydomain.user` may be needed.
   - Basic authentication does not work for users with MFA or for SSO-only users:
     they need native credentials in the pod (or OAuth, not supported yet).
3. `status` lists pods, which ones have a credential, and the figures switch.

## Reading data (`slice`)

```
python epm.py slice <pod> --cube <CUBE> \
  --pov "Year=FY26;Scenario=Actual;Version=Working;Currency=Local" \
  --rows "Account=ILvl0Descendants(PL)" --rows "Entity=E1,E2" \
  --cols "Period=Jan,Feb,Mar"
```

- Every dimension of the cube must be on the POV, rows or columns. Run `dims` first;
  `slice` then warns about dimensions left out (the cube picks their member, which is
  rarely what you want).
- Several `--rows` (or `--cols`) cross their dimensions. Members are comma separated;
  member functions work (`ILvl0Descendants(X)`, `IDescendants(X)`, `Children(X)`).
  Quote names that contain commas.
- Empty rows and columns are suppressed by default (`--no-suppress` keeps them). The
  user only gets cells their own security lets them read.
- `--out file.xlsx --open` writes the whole grid with figures (sheet `Data`, sheet `POV`)
  and opens it for the user. It needs the figures switch at `1`; otherwise it stops. With `--out`, nothing is printed unless `--print`.
- `--grid payload.json` sends a full `exportdataslice` payload for layouts the flags
  cannot express (several row segments).
- Printing stops at `--max-rows` (200); use `--out` for more.
- `--digest` prints a salted fingerprint per row, per column and for the whole grid
  instead of the grid (add `--print` for the masked grid too). Equal fingerprints mean
  equal values, so with figures masked you can still check that a copy matches its
  source (two columns), or whether data changed overnight (same command before and
  after). The salt lives in `digest.key` on this PC, so fingerprints only compare on the
  same PC. An all-empty column has its own constant fingerprint.
- Planning job ids are sequential: `job <pod> <id>` on the ids after a known one lists
  what ran since (rule name and status, not the runtime prompts or times).

## Data Integration (integrations, pipelines)

Oracle has no read-only REST call that lists the integrations, locations or pipelines
defined in a pod. What can be read is what ran:

1. `files <pod>`: every Data Integration process leaves `outbox/logs/<process id>.log`
   (the pod keeps them about 7 days). Their dates show when things run.
2. `dijob <pod> <id>`: process type (`PIPELINE_EXECUTE`, `PIPELINE_JOB_EXECUTE`, data
   load rules...) and status.
3. `cat <pod> outbox/logs/<id>.log`: the pipeline log names the pipeline, its stages and
   every job with its type (businessRule, businessRuleset, epmPlatformJob, integration)
   and parameters. Log times are UTC.
4. `jobdefs <pod> --type IMPORT_DATA` (and `EXPORT_DATA`) lists Planning's own file jobs,
   which are not Data Integration.

To see integrations that are defined but did not run, the user opens Data Integration
in the pod, or exports its setup (that creates a file in the pod, so the user runs it).

## Errors

| Message | Meaning |
|---|---|
| `HTTP 401` | Wrong user or password, user with MFA, or SSO-only user. The user re-runs `cred-set`. |
| `HTTP 403` | The role does not allow it. `member`, `jobdefs`, `job` and `files` need Service Administrator. |
| `HTTP 404` | Application, cube, dimension or member name wrong (check with `cubes`, `dims`, `find`). |
| `HTML page (probably an SSO sign-in page)` | The user signs in through SSO only; REST needs native credentials. |
| `Blocked: ...` | The guard refused a write or a path outside the EPM REST APIs. Do not look for a way around it. |
| `No credential for pod` | The user has not run `cred-set` for that pod on this PC. |

Git Bash rewrites arguments that start with `/` into Windows paths
(`C:/Program Files/Git/HyperionPlanning/...`). `get` repairs that, but you can also
write the path without the leading slash.

Large dimensions: `tree` downloads the whole hierarchy once and caches it; use
`--refresh` after the user changes metadata. Start `tree` from a member and keep
`--depth` low to avoid flooding the output.

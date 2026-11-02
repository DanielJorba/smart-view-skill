# REST calls used by epm.py

Checked against Oracle's "REST APIs for Oracle Fusion Cloud EPM" on 2026-10-29 and
tested live the same day on a Planning pod (OCI), with a user whose role was not
checked: `ping`, `cubes`, `dims`, `tree`, `find`, `subvars`, `slice` and `get`. Base
URL: the pod root, for example `https://<pod>.epm.<region>.ocs.oraclecloud.com`.
Authentication: HTTP Basic in the `Authorization` header. On OCI pods the user goes
without identity domain; users with MFA cannot use Basic authentication.

| Command | Method and path | Role | Oracle page |
|---|---|---|---|
| `ping`, `apps` | `GET /HyperionPlanning/rest/v3/applications` | any | Planning REST APIs |
| `cubes` | `GET /HyperionPlanning/rest/v3/applications/{app}/plantypes` | any | get_plan_types.html (new in 26.04) |
| `dims` | `GET .../applications/{app}/plantypes/{cube}/dimensions?limit=-1` | any | get_dim_plan_types.html (26.04) |
| `tree`, `find` | `GET .../applications/{app}/plantypes/{cube}/dimensions/{dim}[?aliasTableName=]` | any | get_dim_details.html (26.04) |
| `member` | `GET .../applications/{app}/dimensions/{dim}/members/{member}` | Service Administrator | get_member.html |
| `slice` | `POST .../applications/{app}/plantypes/{cube}/exportdataslice` | any (only cells the user can read) | export_dataslices.html |
| `subvars` | `GET .../applications/{app}/substitutionvariables` (all, `planType` "ALL" for application level) | any | not on Oracle's index page; works (tested 2026-10-29 on a Planning pod) |
| `subvars` | `GET .../applications/{app}/plantypes/{cube}/substitutionvariables` | any | planning_get_subst_variables_defined_at_plan_type_level_5.html |
| `jobdefs` | `GET .../applications/{app}/jobdefinitions[?q={"jobType":"RULES"}]` | Service Administrator | get_job_definitions.html |
| `job` | `GET .../applications/{app}/jobs/{jobId}` | Service Administrator | retrieve_job_status.html |
| `files` | `GET /interop/rest/11.1.2.3.600/applicationsnapshots` | Service Administrator or Migrations Administer | list_files.html |
| `cat` | `GET /interop/rest/11.1.2.3.600/applicationsnapshots/{url-encoded name}/contents` | Service Administrator | Download a File (tested 2026-10-29 with `outbox%2Flogs%2F<id>.log`) |
| `dijob` | `GET /aif/rest/V1/jobs/{jobId}` | Service Administrator, Power User | fdmee_jobstatus_di.html (tested 2026-10-29: returns `processType`, `jobStatus`, `logFileName`, `executedBy`; no rule or pipeline name) |

Pages live under
`https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/prest/`.

## Response shapes worth knowing

- `exportdataslice` returns `pov` (list of members), `columns` (one list per column
  dimension, one entry per column) and `rows` (`headers` and `data`, values as strings;
  empty cells come as `""` or `" "`).
- Job status `status`: -1 in progress, 0 success, 1 error, 2 cancel pending, 3 cancelled,
  4 invalid parameter. Field names vary between examples (`descriptiveStatus` or
  `jobStatus`).
- Plan type `cubeType`: 0 BSO, 136 ASO.
- Dimension details return the whole hierarchy as nested `children`, with `name`,
  `alias`, `dataStorage`, `level`, `generation`, `parentName`, `path` and `usedIn`.

## Documentation inconsistencies seen

- Get Member: the sample URL omits `/members/`; the template and the self link include it.
- exportdataslice: `suppressMissingColumns` is described as suppressing rows (copy error).
- Retrieve Job Status: the table says `applicationName`, the template `{application}`.

## Not used yet (candidates)

- Export Data (`POST .../plantypes/{cube}/exportdata`, 26.04): larger grids with segments.
- Export Form Data (26.04): reads the slice of a form by name.
- Get Application Summary (26.04): AI-oriented summary of an application; Oracle warns
  the format can change between releases.
- Get Pipeline Details (`GET /aif/rest/V1/pipeline?pipelineName=<code>`, get_pipeline.html):
  stages, jobs and variables of one pipeline; needs the pipeline code, there is no list.

- Migration v2 (26.04): List Artifacts in a Category and Export Categories Artifacts
  (`selectedArtifacts`). The export writes a snapshot in the pod, so the user runs it.
- Platform REST APIs announced for 26.11: 141 for Planning and FreeForm (applications,
  cubes, dimensions and members, Calculation Manager, forms, dashboards, jobs,
  substitution variables). Read the reference when 26.11 reaches the test pods; the read
  ones may replace calls here, the write ones (create dimensions, batch save of members)
  stay out of the guard.

Data Integration has no documented GET that lists integrations, locations or pipelines
(checked fdmee_rest_apis.html on 2026-10-29 and 2026-10-30); the rest of that chapter
runs or exports. `PUT /aif/rest/V1/objects/connections` (Update Connection) writes
connection secrets: never. The Data Exchange page itself reads lists from undocumented
`/aif/ui/model/...` endpoints with the browser session. They are not added to `epm.py`.

Each needs a test on a test pod before it is added to the guard or the commands.

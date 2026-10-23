# Smart View VBA API reference

Function names, argument names, constants and codes are facts taken from Oracle's
*Smart View for Office Developer's Guide* (EPM Cloud library, October 2026; matches the
26.100 on-premises guide) and from the `smartview.bas` shipped with Smart View 25.200.
Explanations are original. When the guide and `smartview.bas` disagree, **the installed
`smartview.bas` wins**: it is what the DLL actually exports.

Docs: https://docs.oracle.com/en/cloud/saas/enterprise-performance-management-common/svdvg/

## Contents

1. Setup
2. Conventions
3. Return codes and errors
4. Provider support at a glance
5. Connections and sign-in
6. Ad hoc grid
7. POV and data cells
8. Forms
9. Options
10. Menus and refresh
11. Member and outline queries (Essbase only)
12. Functions that change data (recognise, do not run)
13. Template: refresh listed sheets with a log
14. Every function declared in smartview.bas 25.200

## 1. Setup

- All declarations live in `smartview.bas`, in the Smart View `bin` folder (default
  install `C:\Oracle\SmartView\bin\`; `sv.ps1 version` prints the real path). Import it
  into the workbook's VBA project (VBA editor > File > Import File). Do not hand-write
  `Declare` lines or copy them from web samples: Oracle's samples contain errors, and
  the installed file matches the installed DLL and Office bitness (`PtrSafe` for 64-bit).
- Every declaration uses `Lib "HsAddin"`. Smart View must be loaded in that Excel.
- `smartview.bas` also defines the constants (`HSV_*`, `HYP_*`, `SS_*`), the enums
  `SmartViewErrors`, `HYP_SVC_OPTIONS_INDEX`, `TYPE_OF_CONTENTS_IN_SHEET`,
  `DIMENSION_TYPE`, the types `DOC_Info` and `LRO_Info`, and a helper
  `GetReturnCodeMessage(sts As Long) As String`.
- **Keep `Application.EnableEvents = True`** around Smart View calls; Oracle states the
  API relies on Excel events.
- Do not redistribute `smartview.bas`; tell users to import their own copy. (The skill's
  bridge declares the few functions it needs itself, so it ships no Oracle file.)

## 2. Conventions

- Arguments are `Variant`, mostly `ByVal`; outputs are `ByRef`.
- `vtSheetName`: a sheet name, or the keyword `Empty` (not the string `"Empty"`) for the
  active sheet. Activating the target sheet before the call avoids surprises.
- Range arguments take a `Range` object, e.g. `Worksheets("Grid").Range("A5")`.
- Many arguments fall back to a default when passed `Empty` or `Null`.

## 3. Return codes and errors

`0` = success. `1` and `2` = success with a condition (submit only). Negative =
client-side problem. `4` = provider or generic error.

| Code | Constant | Meaning |
|---|---|---|
| 4 | SS_ERR_ERROR | Provider-specific or generic error |
| 2 | SS_NO_GRID_ON_SHEET_BUT_FUNCTIONS_SUBMITTED | No grid, but cell functions were processed |
| 1 | SS_SHEET_NOT_CONNECTED_BUT_FUNCTIONS_SUBMITTED | Not connected, but cell functions were processed |
| 0 | SS_OK | Success |
| -4 | SS_NOT_CONNECTED | Sheet not connected |
| -8 | SS_NOUNDO_INFO | Nothing to undo |
| -9 | SS_CANCELED | Cancelled (often the user closed a dialog) |
| -14 | SS_DIALOG_ERROR | A required dialog could not be shown |
| -15 | SS_INVALID_PARAM | Invalid argument |
| -16 | SS_CALCULATING | A calculation is running |
| -18 | SS_FORMULAPRESERVE | Blocked by formula-preservation mode |
| -19 | SS_INTERNALSSERROR | Operation not possible on this sheet |
| -20 / -21 | SS_INVALID_SHEET / SS_NOACTIVESHEET | Sheet cannot be determined |
| -23 | SS_INVALIDSELECTION | Invalid selection argument |
| -39 | SS_SHEET_PROTECTED | Sheet protected |
| -41 | SS_NOSUPPORT_PROVIDER | Not supported for this provider |
| -42 | SS_INVALID_ALIAS | Invalid alias table |
| -43 | SS_CONN_NOT_FOUND | Connection name not found |
| -47 | SS_CONN_ALREADY_EXISTS | Connection already exists |
| -55 | SS_OPERATION_FAILED | Operation failed |
| -56 | SS_CANNOT_ASSOCIATE_SHEET_WITH_CONNECTION | Sheet cannot be bound to the connection |
| -57 | SS_REFRESH_SHEET_NEEDED | Refresh the sheet first |
| -58 / -62 | SS_NO_GRID_OBJECT_ON_SHEET / SS_NO_SVC_CONTENT_ON_SHEET | No grid / no Smart View content |
| -59 | SS_NO_CONNECTION_ASSOCIATED | Sheet has no connection |
| -60 | SS_NON_DATA_CELL_PASSED | A member cell was passed where a data cell is needed |
| -61 | SS_DATA_CELL_IS_NOT_WRITABLE | Read-only data cell |
| -64 | SS_OP_FAILED_AS_CHART_IS_SELECTED | A chart is selected |
| -65 | SS_EXCEL_IN_EDIT_MODE | Excel is in cell-edit mode |
| -68 | SS_SMART_VIEW_DISABLED | Smart View disabled |
| -69 | SS_VBA_DEPRECATED | Deprecated call or use (e.g. `HypSetPOV` on a form) |
| -70 | SS_OPERATION_NOT_SUPPORTED_IN_MULTIGRID_MODE | Not on multiple-grid sheets |
| -71 | SS_INVALID_MEMBER | Invalid member name |
| -73 | SS_AMBIGUOUS_MENU | Ambiguous menu label (`HypExecuteMenu`) |

The full list (-1 to -73) is the `SmartViewErrors` enum in `smartview.bas`.

After any non-zero code, call:

```vba
Dim code As Variant, msg As Variant, desc As Variant
HypGetLastError code, msg, desc          ' the server's own message, usually the useful one
' HypGetErrorInfo(code, text) or GetReturnCodeMessage(rc) turn a code into text
```

Functions with other return conventions: `HypConnected` (-1 true / 0 false),
`HypConnectionExists` (Boolean), `HypIsAncestor`, `HypIsDescendant`, `HypIsParent`,
`HypIsChild`, `HypIsUDA`, `HypIsAttribute` (-1 / 0 / error code), `HypIsCellWritable`
(Boolean), `HypCell` (value or error string), `HypGetGlobalOption` and
`HypGetSheetOption` (the value, or a negative code).

## 4. Provider support at a glance

| Area | EPM Cloud Planning, FCCS, Tax Reporting | Essbase (cloud, 21c+) | On-prem Planning / HFM |
|---|---|---|---|
| Connections, disconnect, alias table | Yes (`HypUIConnect`, not `HypConnect`) | Yes | Yes (`HypConnect` works) |
| Retrieve, zoom, keep/remove only, pivot, undo | Yes, ad hoc grids | Yes | Planning, HFM: ad hoc |
| Forms (`HypOpenForm`, `HypSetPages`, ad hoc on form) | Yes | n/a | Planning, HFM |
| Member queries, member information, outline info | **No** | Yes | No |
| Substitution variables (`HypGetSubstitutionVariable`) | **No** | Yes | No |
| MDX (`HypExecuteQuery`) | **No** | Yes | No |
| Business rules (`HypListCalcScriptsEx`) | Listed, but running them is off limits | Yes | Planning |

For Planning-family structure questions use ad hoc exploration or the EPM Admin
Extension instead ([structure-discovery.md](structure-discovery.md)).

## 5. Connections and sign-in

```vba
HypUIConnect(vtSheetName, vtUserName, vtPassword, vtFriendlyName) As Long   ' cloud and on-prem
HypConnect(vtSheetName, vtUserName, vtPassword, vtFriendlyName) As Long     ' on-prem only
HypConnected(vtSheetName) As Variant
HypConnectionExists(vtFriendlyName) As Variant
HypCreateConnectionEx(vtProviderType, vtServerName, vtApplicationName, vtDatabaseName, _
    vtFormName, vtProviderURL, vtFriendlyName, vtUserName, vtPassword, vtDescription, _
    vtReserved1, vtReserved2) As Long
HypUpdateConnection(vtProviderType, vtServerName, vtApplicationName, vtDatabaseName, _
    vtProviderURL, vtFriendlyName, vtUserName, vtPassword, vtDescription, _
    bCreateNewIfConnectionDoesnotExist As Boolean) As Long
HypDisconnect(vtSheetName, bLogoutUser As Boolean) As Long
HypDisconnectEx(vtFriendlyName) As Long
HypDisconnectAll() As Long
HypSetActiveConnection(vtFriendlyName) As Long        ' bind the active sheet to an open connection
HypSetConnAliasTable(vtFriendlyName, vtAliasTableName) As Long   ' "None" = no alias table
HypSetAliasTable(vtSheetName, vtAliasTableName) As Long
HypSetAsDefault(vtFriendlyName) As Long
HypGetConnectionInfo(...) As Long                     ' has a password slot: never log it
```

- **EPM Cloud:** Oracle documents `HypConnect` as not working against cloud sources.
  `HypUIConnect` works on an existing private connection and **always shows Smart
  View's sign-in** for cloud sources. Design macros so the user signs in; do not pass
  passwords.
- `vtFriendlyName` accepts a private connection name or an inline string
  `"URL|server|app|db"`.
- Provider constants: `HYP_PLANNING = "Planning"`, `HYP_ESSBASE = "Essbase"`,
  `HYP_FINANCIAL_MANAGEMENT`, `HYP_RA`. Using `HYP_PLANNING` with
  `https://<host>/HyperionPlanning/SmartView` for FCCS or Tax Reporting follows from the
  provider lists but Oracle shows no cloud sample: test it.
- For Planning connections to appear in the Smart View panel, Oracle says
  `HypCreateConnectionEx` needs `vtFormName` (a form path) and `vtProviderURL`.
- Non-interactive authentication functions exist (`HypSetBasicAuthToken`,
  `HypSetCustomHeader` with an OAuth bearer token, `HypSetSSO`,
  `HypLoginEssbaseCloudBasicAuth`). They put credentials or tokens in code; this skill
  does not use them. Unattended jobs belong on the server (EPM Automate, REST).

URL formats: Planning-family cloud `https://<host>/HyperionPlanning/SmartView`;
Profitability and Cost Management `https://<host>/aps/SmartView`; Essbase cloud and
21c+ `https://<host>/essbase/smartview`; on-prem Planning
`http(s)://<server>:<port>/HyperionPlanning/SmartView`.

## 6. Ad hoc grid

```vba
HypRetrieve(vtSheetName) As Long                                 ' refresh the sheet
HypRetrieveRange(vtSheetName, vtRange, vtFriendlyName) As Long   ' refresh one grid range; clears Undo
HypRetrieveAllWorkbooks() As Long                                ' every open workbook
HypZoomIn(vtSheetName, vtSelection, vtLevel, vtAcross) As Long   ' vtAcross unused
HypZoomOut(vtSheetName, vtSelection) As Long
HypKeepOnly(vtSheetName, vtSelection) As Long                    ' member cells only
HypRemoveOnly(vtSheetName, vtSelection) As Long
HypPivot(vtSheetName, vtStart, vtEnd) As Long                    ' single-cell ranges
HypPivotToGrid(vtSheetName, vtDimensionName, vtSelection) As Long
HypPivotToPOV(vtSheetName, vtSelection) As Long
HypUndo(vtSheetName) As Long
HypRedo(vtSheetName) As Long
HypIsSmartViewContentPresent(vtSheetName, ByRef vtType As TYPE_OF_CONTENTS_IN_SHEET) As Boolean
HypGetSheetInfo(vtSheetName, ByRef vtItemNames, ByRef vtItemValues) As Long
HypGetDimensions(vtSheetName, ByRef vtMemberNames, ByRef vtType) As Long   ' types: DIMENSION_TYPE
```

`HypZoomIn` levels: 0 next level, 1 all levels, 2 bottom level, 3 siblings, 4 same
level, 5 same generation, 6 formulas (3 to 6 documented for Essbase). `Empty` uses the
current option.

`TYPE_OF_CONTENTS_IN_SHEET`: 0 empty, 1 ad hoc, 2 form, 3 interactive report.
`DIMENSION_TYPE`: 0 row, 1 column, 2 POV, 3 page, 5 user variable.

## 7. POV and data cells

```vba
HypSetPOV(vtSheetName, ParamArray MemberList()) As Long     ' ad hoc: "Dimension#Member" items
HypSetPages(vtSheetName, ParamArray MemberList()) As Long   ' forms: "Dimension#Member" items
HypGetPOV(vtSheetName, ByRef vtDimensionNames, ByRef vtMemberNames, ByRef vtType) As Long
HypGetBackgroundPOV(vtFriendlyName, ByRef vtDimensionNames, ByRef vtMemberNames) As Long
HypSetBackgroundPOV(vtFriendlyName, ParamArray MemberList()) As Long
HypGetPagePOVChoices(vtSheetName, vtDimensionName, ByRef vtNames, ByRef vtDescs) As Long
HypGetDimMbrsForDataCell(vtSheetName, vtCellRange, ByRef vtServer, ByRef vtApp, _
    ByRef vtCube, ByRef vtForm, ByRef vtDimensionNames, ByRef vtMemberNames) As Long
HypCell(vtSheetName, ParamArray MemberList()) As Variant    ' value of one intersection, not written to the sheet
HypIsCellWritable(vtSheetName, vtCellRange) As Boolean
```

- Example: `HypSetPOV(Empty, "Year#FY26", "Scenario#Actual")`. On a form, `HypSetPOV`
  returns -69: use `HypSetPages`.
- `HypGetPOV` is declared in `smartview.bas` 25.200 but missing from Oracle's guide.
  It works as a reader of the sheet POV; if it fails on a provider, fall back to
  `HypGetDimMbrsForDataCell` on any data cell.

## 8. Forms

```vba
HypListDocuments(vtSheetName, vtUserName, vtPassword, vtConnInfo, vtCompletePath, ByRef vtDocs As DOC_Info) As Long
HypOpenForm(vtSheetName, vtFolderPath, vtFormName, ByRef vtDimensionList() As Variant, ByRef vtMemberList() As Variant) As Long  ' arrays unused
HypPerformAdhocOnForm(vtSheetName, vtFormName) As Long      ' full path, e.g. "/Forms/<folder>/<form>"
HypIsFreeForm(vtSheetName) As Boolean
```

`DOC_Info` fields: `numDocs`, `docTypes`, `docNames`, `docDescriptions`,
`docPlanTypes`, `docAttributes`. Type strings `DOC_FORM`, `DOC_FOLDER`.

## 9. Options

Use the unified API (`HSV_*` indices). With a sheet name the option applies to that
sheet; with `""` it changes the user's default options, which persist. Prefer sheet
options. If a macro must change a default option, read it first with `HypGetOption`
and restore it in the cleanup path that also runs after an error, and tell the user.

```vba
HypSetOption(vtItem, vtOption, vtSheetName) As Long
HypGetOption(vtItem, ByRef vtRet, vtSheetName) As Long
```

| Idx | Constant | Values |
|---|---|---|
| 1 | HSV_ZOOMIN | 0 next, 1 all, 2 bottom, 3 siblings, 4 same level, 5 same generation, 6 formulas |
| 2 | HSV_INCLUDE_SELECTION | Boolean |
| 5 | HSV_INDENTATION | 0 none, 1 sub-items, 2 totals |
| 6 / 7 | HSV_SUPPRESSROWS_MISSING / _ZEROS | Boolean |
| 9 | HSV_SUPPRESSROWS_NOACCESS | Boolean |
| 10 | HSV_SUPPRESSROWS_REPEATEDMEMBERS | Boolean |
| 12 | HSV_ANCESTOR_POSITION | 0 top, 1 bottom |
| 13 | HSV_MISSING_LABEL | Text, e.g. `#Missing` |
| 16 | HSV_MEMBER_DISPLAY | 0 name, 1 name and description (alias), 2 description only |
| 18 | HSV_SUBMITZERO | Boolean (submit-related; leave alone) |
| 21 | HSV_PRESERVE_FORMULA_COMMENT | Boolean: keep formulas and comments in ad hoc |
| 30 | HSV_EXCEL_FORMATTING | Boolean |
| 33 | HSV_NAVIGATE_WITHOUTDATA | Boolean: zoom, pivot, keep/remove only without fetching data |
| 36 | HSV_ADJUSTCOLUMNWIDTH | Boolean |
| 41 / 42 | HSV_SUPPRESSCOLUMNS_MISSING / _ZEROS | Boolean |
| 46 | HSV_USE_EXCEL_INDENTATION | Boolean (with 5) |
| 102 | HSV_UNDO_ENABLE | Boolean |
| 104 | HSV_LOGMESSAGE_DISPLAY | 0 info, 1 warnings, 2 errors, 3 none, 4 extended |
| 105 / 113 | HSV_ROUTE_LOGMESSAGE_TO_FILE / _FILE_LOCATION | Boolean / path |
| 112 | HSV_NUMBER_OF_MRU_ITEMS | Number (15 or fewer) |
| 121 | HSV_REFRESH_SELECTED_DEPENDENT_FUNCTIONS | Boolean |

The full enum `HYP_SVC_OPTIONS_INDEX` (1 to 125, with gaps) is in `smartview.bas`.

Legacy API: `HypSetSheetOption` / `HypGetSheetOption` use items 1 to 16 with the same
meaning as unified 1 to 16. `HypSetGlobalOption` / `HypGetGlobalOption` use a
**different numbering** (for example global item 9 = navigate without data, item 1 =
Excel formatting). Prefer the unified API to avoid mixing them up.

## 10. Menus and refresh

```vba
HypMenuVRefresh() As Long        ' active sheet, like the ribbon button
HypMenuVRefreshAll() As Long     ' all connected sheets of the active workbook
HypExecuteMenu(vtSheetName, vtMenuName) As Long
```

- `HypMenuV*` functions behave exactly like the ribbon command, with the same
  preconditions (and the same dialogs).
- `HypExecuteMenu` runs a ribbon item by its **visible label in the Office language**
  (a Spanish Office needs Spanish labels). Prefix the ribbon when a label repeats:
  `"Smart View->Refresh"`. Only action items work; an ambiguous label returns -73.
- Commands that open a modal dialog block VBA until the dialog closes
  ([automation.md](automation.md)).

## 11. Member and outline queries (Essbase only)

```vba
HypFindMember(vtSheetName, vtMemberName, vtAliasTable, ByRef vtDimensionName, _
    ByRef vtAliasName, ByRef vtGenerationName, ByRef vtLevelName) As Long
HypGetMemberInformation(vtSheetName, vtMemberName, vtPropertyName, ByRef vtValue, ByRef vtValueStrings) As Long
HypGetMemberInformationEx(vtSheetName, vtMemberName, ByRef vtNames, ByRef vtValues, ByRef vtValueStrings) As Long
HypGetChildren(vtSheetName, vtMemberName, intChildCount As Integer, ByRef vtChildArray) As Long  ' <=0 = all
HypGetParent(vtSheetName, vtMemberName, ByRef vtParentName) As Long
HypGetAncestor(vtSheetName, vtMemberName, vtLayerType, intLayerNumber As Integer, ByRef vtAncestor) As Long  ' "Gen" or "Level"
HypQueryMembers(vtSheetName, vtMemberName, vtPredicate, vtOption, vtDimensionName, vtInput1, vtInput2, ByRef vtMemberArray) As Long
HypOtlGetMemberInfo(vtSheetName, vtDimensionName, vtMemberName, vtPredicate, ByRef vtMemberArray) As Long
HypIsUDA(vtSheetName, vtDimensionName, vtMemberName, vtUDAString) As Variant
HypIsAttribute(vtSheetName, vtDimensionName, vtMemberName, vtAttribute) As Variant
HypGetSubstitutionVariable(vtSheetName, vtApp, vtDb, vtVariableName, ByRef vtNames, ByRef vtValues) As Long
HypExecuteQuery(vtSheetName, vtMDXQuery) As Long
```

- `HypQueryMembers` predicates: 1 children, 2 descendants, 3 bottom level, 4 siblings,
  5 same level, 6 same generation, 7 parent, 8 dimension, 9 named generation, 10 named
  level, 11 search, 12 wildcard search (trailing `*`), 13 user attribute, 14 ancestors,
  15 DTS member, 16 dimension user attributes. Options: 1 members only, 2 aliases only,
  4 members and aliases.
- `HypOtlGetMemberInfo` predicates: 1 comment, 2 formula, 3 UDA, 4 attribute.
- Member property names (`HYP_MI_*`): `Name`, `Dim`, `Level`, `Generation`,
  `ParentMbrName`, `ChildMbrName`, `Consolidation`, `IsTwoPassCalcMbr`, `IsExpenseMbr`,
  `TimeBalanceOption`, `ShareOption`, `StorageCategory`, `ChildCount`, `DefaultAlias`,
  `AliasNames`, `AliasTables`, `AttributeDims`, `AttributeMbrs`, `Formula`, `Comment`,
  `Udas`, `FormatString` and others in `smartview.bas`.
- Oracle's samples for `HypIsAncestor` / `HypIsDescendant` contradict the parameter
  names on argument order: test against a known outline before relying on them.

## 12. Functions that change data (recognise, do not run)

`HypSubmitData`, `HypSubmitSelectedDataCells`, `HypSubmitSelectedRangeWithoutRefresh`,
`HypMenuVSubmitData`, `HypSetCellsDirty`, `HypExecuteCalcScript`, `HypExecuteCalcScriptEx`
(business rules), `HypExecuteCalcScriptEx2`, `HypExecuteCalcScriptString`,
`HypMenuVBusinessRules`, `HypMenuVCalculation`, `HypCalculate`, `HypConsolidate*`,
`HypForce*`, `HypTranslate`, `HypMenuVAdjust`, `HypSetSubstitutionVariable`,
`HypAddLRO`, `HypUpdateLRO`, `HypDeleteLROs`, `HypDeleteAllLROs`, `HypDeleteCalc`,
`HypMenuVSyncBack`, `HypMenuVSupportingDetails`, and any sheet with `HsSetValue` /
`HsSetText` cells (they write on Submit).

When a user asks for a macro that needs one of these, write it in a separate,
clearly named procedure, explain what it changes on the server, and recommend running
it first on a test environment. Never run it through the bridge.

Not data changes, despite the names: `HypMenuVCopyDataPoints` / `HypMenuVPasteDataPoints`
copy the intersection of cells and paste it as `HsGetValue` formulas (local sheet only).

Housekeeping functions that change the user's setup (not data): `HypRemoveConnection`,
`HypDeleteMetaData` (strips Smart View metadata from a sheet or workbook),
`HypDeleteAllMRUItems`, `HypSetGlobalOption`. Use only on request.

## 13. Template: refresh listed sheets with a log

```vba
Option Explicit

' Refreshes the Smart View sheets listed in Control!A2:A and writes the result in B:C.
' Prerequisites: smartview.bas imported in this project; the user is signed in
' (for EPM Cloud, connect once interactively before running).
Private Const CONTROL_SHEET As String = "Control"

Public Sub RefreshListedSheets()
    Dim ctl As Worksheet, ws As Worksheet, r As Long, shName As String, rc As Long
    Set ctl = ThisWorkbook.Worksheets(CONTROL_SHEET)
    For r = 2 To ctl.Cells(ctl.Rows.Count, 1).End(xlUp).Row
        shName = CStr(ctl.Cells(r, 1).Value)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(shName)
        On Error GoTo 0
        If ws Is Nothing Then
            ctl.Cells(r, 2).Value = "n/a"
            ctl.Cells(r, 3).Value = "Sheet not found"
        Else
            ws.Activate
            rc = HypRetrieve(shName)
            ctl.Cells(r, 2).Value = rc
            If rc = 0 Then
                ctl.Cells(r, 3).Value = "OK " & Format$(Now, "yyyy-mm-dd hh:nn")
            Else
                ctl.Cells(r, 3).Value = LastSmartViewError()
            End If
        End If
    Next r
    ctl.Activate
End Sub

Private Function LastSmartViewError() As String
    Dim code As Variant, msg As Variant, desc As Variant
    HypGetLastError code, msg, desc
    LastSmartViewError = Trim$(CStr(code) & " " & CStr(msg) & " " & CStr(desc))
End Function
```

To change the POV first: `rc = HypSetPOV(shName, "Year#" & ctl.Range("E1").Value,
"Period#" & ctl.Range("E2").Value)` before `HypRetrieve`, checking `rc` the same way.
Dimensions that sit in header rows of the grid are changed by writing the cell, not
with `HypSetPOV`.

## 14. Every function declared in smartview.bas 25.200

Use this list to answer "does this function exist?". If it is not here, it does not
exist in that version; check the user's own `smartview.bas` for newer releases.

HypAddLRO, HypCalculate, HypCalculateContribution, HypCaptureFormatting, HypCell,
HypConnect, HypConnectToAPS, HypConnected, HypConnectionExists, HypConsolidate,
HypConsolidateAll, HypConsolidateAllWithData, HypCopyMetaData, HypCreateConnection,
HypCreateConnectionEx, HypCreateRangeGrid, HypDeleteAllLROs, HypDeleteAllMRUItems,
HypDeleteCalc, HypDeleteLROs, HypDeleteMetaData, HypDisconnect, HypDisconnectAll,
HypDisconnectEx, HypDisconnectFromAPS, HypDisplayToLinkView, HypExecuteCalcScript,
HypExecuteCalcScriptEx, HypExecuteCalcScriptEx2, HypExecuteCalcScriptString,
HypExecuteDrillThroughReport, HypExecuteMDXEx, HypExecuteMenu, HypExecuteQuery,
HypFindMember, HypFindMemberEx, HypForceCalculate, HypForceCalculateContribution,
HypForceTranslate, HypFreeDataPoint, HypGetActiveMember, HypGetAncestor,
HypGetBackgroundPOV, HypGetCalcScript, HypGetCellRangeForMbrCombination, HypGetChildren,
HypGetColCount, HypGetColItems, HypGetConnectionInfo, HypGetCurrentAPSURL,
HypGetDataPoint, HypGetDatabaseNote, HypGetDimMbrsForDataCell, HypGetDimensions,
HypGetDrillThroughReports, HypGetErrorInfo, HypGetGlobalOption, HypGetLastError,
HypGetLinkMacro, HypGetMemberInformation, HypGetMemberInformationEx, HypGetMembers,
HypGetNameRangeList, HypGetOption, HypGetOverrideFlag, HypGetPOV, HypGetPOVCount,
HypGetPOVItems, HypGetPagePOVChoices, HypGetParent, HypGetRowCount, HypGetRowItems,
HypGetSharedConnectionsURL, HypGetSheetInfo, HypGetSheetOption, HypGetSourceGrid,
HypGetSubstitutionVariable, HypGetVersion, HypHideRibbonMenu, HypHideRibbonMenuReset,
HypInvalidateSSO, HypIsAncestor, HypIsAttribute, HypIsCellWritable, HypIsChild,
HypIsConnectedToAPS, HypIsConnectedToSharedConnections, HypIsDataModified,
HypIsDescendant, HypIsExpense, HypIsFreeForm, HypIsParent, HypIsSmartViewContentPresent,
HypIsUDA, HypKeepOnly, HypListApplications, HypListCalcScripts, HypListCalcScriptsEx,
HypListDatabases, HypListDocuments, HypListLROs, HypListUserVariables,
HypLoginEssbaseCloudBasicAuth, HypMenuVAbout, HypMenuVAdjust, HypMenuVBusinessRules,
HypMenuVCalculation, HypMenuVCascadeNewWorkbook, HypMenuVCascadeSameWorkbook,
HypMenuVCellText, HypMenuVCollapse, HypMenuVConnect, HypMenuVCopyDataPoints,
HypMenuVExpand, HypMenuVFunctionBuilder, HypMenuVInstruction, HypMenuVKeepOnly,
HypMenuVLRO, HypMenuVMemberInformation, HypMenuVMemberSelection, HypMenuVMigrate,
HypMenuVOptions, HypMenuVPOVManager, HypMenuVPasteDataPoints, HypMenuVPivot,
HypMenuVQueryDesigner, HypMenuVRedo, HypMenuVRefresh, HypMenuVRefreshAll,
HypMenuVRefreshOfflineDefinition, HypMenuVRemoveOnly, HypMenuVRulesOnForm,
HypMenuVRunReport, HypMenuVSelectForm, HypMenuVShowHelpHtml, HypMenuVSubmitData,
HypMenuVSupportingDetails, HypMenuVSyncBack, HypMenuVTakeOffline, HypMenuVUndo,
HypMenuVVisualizeinExcel, HypMenuVVisualizeinHVE, HypMenuVZoomIn, HypMenuVZoomOut,
HypMigrateConnectionToDataSourceMgr, HypModifyConnection, HypModifyRangeGridName,
HypOpenFlexForm, HypOpenForm, HypOtlGetMemberInfo, HypPerformAdhocOnForm, HypPivot,
HypPivotToGrid, HypPivotToPOV, HypPreserveFormatting, HypQueryMembers, HypRedo,
HypRemoveCapturedFormats, HypRemoveConnection, HypRemoveOnly, HypRemovePreservedFormats,
HypResetFriendlyName, HypRetrieve, HypRetrieveAllWorkbooks, HypRetrieveLRO,
HypRetrieveNameRange, HypRetrieveRange, HypSetActiveConnection, HypSetActiveMember,
HypSetAliasTable, HypSetAsDefault, HypSetBackgroundPOV, HypSetBasicAuthToken,
HypSetCellsDirty, HypSetColItems, HypSetConnAliasTable, HypSetConnectionInfo,
HypSetCurrentUserAPSURL, HypSetDimensions, HypSetGlobalOption, HypSetLinkMacro,
HypSetMembers, HypSetMenu, HypSetOption, HypSetOverrideFlag, HypSetPOV, HypSetPOVItems,
HypSetPages, HypSetRowItems, HypSetSSO, HypSetSharedConnectionsURL, HypSetSheetOption,
HypSetSubstitutionVariable, HypSetUserVariable, HypShowPanel, HypShowPov,
HypSubmitData, HypSubmitSelectedDataCells, HypSubmitSelectedRangeWithoutRefresh,
HypTranslate, HypUIConnect, HypUndo, HypUpdateConnection, HypUpdateLRO,
HypUseLinkMacro, HypZoomIn, HypZoomOut.

Newer guides also document `HypSetCustomHeader`, absent from this file.

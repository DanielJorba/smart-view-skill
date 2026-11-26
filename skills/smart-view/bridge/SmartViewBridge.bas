Attribute VB_Name = "SmartViewBridge"
'==============================================================================
' SmartViewBridge
' Read-only JSON wrappers over the Oracle Smart View VBA API, called by
' scripts/sv.ps1 through Application.Run.
'
' Requirements
'   - Smart View for Office loaded in the same Excel.
'   - This module saved alone in an add-in (bridge/SmartViewBridge.xlam). It
'     declares the DLL functions it uses itself, so no Oracle file is needed.
'
' Design rules
'   - Nothing here changes cube data or server artifacts. Submit, calculate,
'     business rules, calc scripts, consolidate/translate, substitution
'     variable changes, linked reporting objects and approvals are deliberately
'     not wrapped. SVB_ExecuteMenu refuses menu names that do those things.
'   - No password is ever accepted or returned. Sign-in happens in Smart View's
'     own dialog.
'   - When no workbook name is given, wrappers work on the scratch workbook
'     (created on demand), never on the user's workbooks.
'   - Every public function returns a JSON string:
'     {"ok":bool,"rc":n,"rcText":"...","error":{...}|null,"data":...}
'==============================================================================
Option Explicit

Private Const BRIDGE_VERSION As String = "0.2.1"
Private Const SCRATCH_TAG As String = "SmartViewBridgeScratch"

'------------------------------------------------------------------------------
' Smart View DLL entry points used by this module. Written for this bridge from
' the published signatures (Oracle Smart View Developer's Guide); Oracle's
' smartview.bas is not needed. Private, so they never clash with a workbook
' that imports smartview.bas. Every argument is a Variant, so PtrSafe alone
' works on 32-bit and 64-bit Office.
'------------------------------------------------------------------------------
Private Type SvbDocInfo
    numDocs As Long
    docTypes As Variant
    docNames As Variant
    docDescriptions As Variant
    docPlanTypes As Variant
    docAttributes As Variant
End Type

Private Declare PtrSafe Function HypCell Lib "HsAddin" (ByVal vtSheetName As Variant, ParamArray MemberList() As Variant) As Variant
Private Declare PtrSafe Function HypMenuVRefresh Lib "HsAddin" () As Long
Private Declare PtrSafe Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtFriendlyName As Variant) As Long
Private Declare PtrSafe Function HypConnected Lib "HsAddin" (ByVal vtSheetName As Variant) As Variant
Private Declare PtrSafe Function HypConnectionExists Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Variant
Private Declare PtrSafe Function HypCreateConnectionEx Lib "HsAddin" (ByVal vtProviderType As Variant, ByVal vtServerName As Variant, ByVal vtApplicationName As Variant, ByVal vtDatabaseName As Variant, ByVal vtFormName As Variant, ByVal vtProviderURL As Variant, ByVal vtFriendlyName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtDescription As Variant, ByVal vtReserved1 As Variant, ByVal vtReserved2 As Variant) As Long
Private Declare PtrSafe Function HypDisconnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal bLogoutUser As Boolean) As Long
Private Declare PtrSafe Function HypExecuteDrillThroughReport Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelectionRange As Variant, ByVal vtID As Variant, ByVal vtName As Variant, ByVal vtURL As Variant, ByVal vtURLTemplate As Variant, ByVal vtType As Variant) As Long
Private Declare PtrSafe Function HypExecuteMenu Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMenuName As Variant) As Long
Private Declare PtrSafe Function HypExecuteQuery Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMDXQuery As Variant) As Long
Private Declare PtrSafe Function HypFindMember Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByVal vtAliasTable As Variant, ByRef vtDimensionName As Variant, ByRef vtAliasName As Variant, ByRef vtGenerationName As Variant, ByRef vtLevelName As Variant) As Long
Private Declare PtrSafe Function HypGetAncestor Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByVal vtLayerType As Variant, ByVal intLayerNumber As Integer, ByRef vtAncestor As Variant) As Long
Private Declare PtrSafe Function HypGetBackgroundPOV Lib "HsAddin" (ByVal vtFriendlyName As Variant, ByRef vtDimensionNames As Variant, ByRef vtMemberNames As Variant) As Long
Private Declare PtrSafe Function HypGetChildren Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByVal intChildCount As Integer, ByRef vtChildNameArray As Variant) As Long
Private Declare PtrSafe Function HypGetConnectionInfo Lib "HsAddin" (ByRef vtServerName As Variant, ByRef vtUserName As Variant, ByRef vtPassword As Variant, ByRef vtApplicationName As Variant, ByRef vtDatabaseName As Variant, ByRef vtFriendlyName As Variant, ByRef vtURL As Variant, ByRef vtProviderType As Variant) As Long
Private Declare PtrSafe Function HypGetDatabaseNote Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtDBNote As Variant) As Long
Private Declare PtrSafe Function HypGetDimMbrsForDataCell Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtCellRange As Variant, ByRef vtServerName As Variant, ByRef vtAppName As Variant, ByRef vtCubeName As Variant, ByRef vtFormName As Variant, ByRef vtDimensionNames As Variant, ByRef vtMemberNames As Variant) As Long
Private Declare PtrSafe Function HypGetDimensions Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtMemberNames As Variant, ByRef vtType As Variant) As Long
Private Declare PtrSafe Function HypGetDrillThroughReports Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelectionRange As Variant, ByRef vtIDs As Variant, ByRef vtNames As Variant, ByRef vtURLs As Variant, ByRef vtURLTemplates As Variant, ByRef vtTypes As Variant) As Long
Private Declare PtrSafe Function HypGetErrorInfo Lib "HsAddin" (ByVal vtErrCode As Variant, ByRef vtError As Variant) As Long
Private Declare PtrSafe Function HypGetGlobalOption Lib "HsAddin" (ByVal vtItem As Long) As Variant
Private Declare PtrSafe Function HypGetLastError Lib "HsAddin" (ByRef vtErrorCode As Variant, ByRef vtErrorMessage As Variant, ByRef vtErrorDescription As Variant) As Long
Private Declare PtrSafe Function HypGetMemberInformation Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByVal vtPropertyName As Variant, ByRef vtPropertyValue As Variant, ByRef vtPropertyValueStrings As Variant) As Long
Private Declare PtrSafe Function HypGetMemberInformationEx Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByRef vtPropertyNames As Variant, ByRef vtPropertyValues As Variant, ByRef vtPropertyValueStrings As Variant) As Long
Private Declare PtrSafe Function HypGetMembers Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtDimensionName As Variant, ByRef vtMbrNameChoices As Variant, ByRef vtMbrDescChoices As Variant) As Long
Private Declare PtrSafe Function HypGetOption Lib "HsAddin" (ByVal vtItem As Variant, ByRef vtRet As Variant, ByVal vtSheetName As Variant) As Long
Private Declare PtrSafe Function HypGetPOV Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtDimensionNames As Variant, ByRef vtMemberNames As Variant, ByRef vtType As Variant) As Long
Private Declare PtrSafe Function HypGetPagePOVChoices Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtDimensionName As Variant, ByRef vtMbrNameChoices As Variant, ByRef vtMbrDescChoices As Variant) As Long
Private Declare PtrSafe Function HypGetParent Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByRef vtParentName As Variant) As Long
Private Declare PtrSafe Function HypGetSheetInfo Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtItemNames As Variant, ByRef vtItemValues As Variant) As Long
Private Declare PtrSafe Function HypGetSheetOption Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtItem As Variant) As Variant
Private Declare PtrSafe Function HypGetSubstitutionVariable Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtApplicationName As Variant, ByVal vtDatabaseName As Variant, ByVal vtVariableName As Variant, ByRef vtVariableNames As Variant, ByRef vtVariableValues As Variant) As Long
Private Declare PtrSafe Function HypIsAttribute Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtDimensionName As Variant, ByVal vtMemberName As Variant, ByVal vtUDAString As Variant) As Variant
Private Declare PtrSafe Function HypIsCellWritable Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtCellRange As Variant) As Boolean
Private Declare PtrSafe Function HypIsSmartViewContentPresent Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtTypeOfContentsInSheet As Long) As Boolean
Private Declare PtrSafe Function HypIsUDA Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtDimensionName As Variant, ByVal vtMemberName As Variant, ByVal vtUDAString As Variant) As Variant
Private Declare PtrSafe Function HypKeepOnly Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelection As Variant) As Long
Private Declare PtrSafe Function HypListDocuments Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtConnInfo As Variant, ByVal vtCompletePath As Variant, ByRef vtDocs As SvbDocInfo) As Long
Private Declare PtrSafe Function HypListUserVariables Lib "HsAddin" (ByVal vtFriendlyName As Variant, ByRef vtDimensionNames As Variant, ByRef vtUserVariableNames As Variant, ByRef vtMemberNames As Variant) As Long
Private Declare PtrSafe Function HypMenuVRefreshAll Lib "HsAddin" () As Long
Private Declare PtrSafe Function HypOpenForm Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtFolderPath As Variant, ByVal vtFormName As Variant, ByRef vtDimensionList() As Variant, ByRef vtMemberList() As Variant) As Long
Private Declare PtrSafe Function HypOtlGetMemberInfo Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtDimensionName As Variant, ByVal vtMemberName As Variant, ByVal vtPredicate As Variant, ByRef vtMemberArray As Variant) As Long
Private Declare PtrSafe Function HypPerformAdhocOnForm Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtFormName As Variant) As Long
Private Declare PtrSafe Function HypPivot Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtStart As Variant, ByVal vtEnd As Variant) As Long
Private Declare PtrSafe Function HypPivotToGrid Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtDimensionName As Variant, ByVal vtSelection As Variant) As Long
Private Declare PtrSafe Function HypPivotToPOV Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelection As Variant) As Long
Private Declare PtrSafe Function HypQueryMembers Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtMemberName As Variant, ByVal vtPredicate As Variant, ByVal vtOption As Variant, ByVal vtDimensionName As Variant, ByVal vtInput1 As Variant, ByVal vtInput2 As Variant, ByRef vtMemberArray As Variant) As Long
Private Declare PtrSafe Function HypRedo Lib "HsAddin" (ByVal vtSheetName As Variant) As Long
Private Declare PtrSafe Function HypRemoveOnly Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelection As Variant) As Long
Private Declare PtrSafe Function HypRetrieve Lib "HsAddin" (ByVal vtSheetName As Variant) As Long
Private Declare PtrSafe Function HypRetrieveRange Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtRange As Variant, ByVal vtFriendlyName As Variant) As Long
Private Declare PtrSafe Function HypSetActiveConnection Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long
Private Declare PtrSafe Function HypSetAliasTable Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtAliasTableName As Variant) As Long
Private Declare PtrSafe Function HypSetConnAliasTable Lib "HsAddin" (ByVal vtFriendlyName As Variant, ByVal vtAliasTableName As Variant) As Long
Private Declare PtrSafe Function HypSetGlobalOption Lib "HsAddin" (ByVal vtItem As Long, ByVal vtGlobalOption As Variant) As Long
Private Declare PtrSafe Function HypSetOption Lib "HsAddin" (ByVal vtItem As Variant, ByVal vtOption As Variant, ByVal vtSheetName As Variant) As Long
Private Declare PtrSafe Function HypSetPOV Lib "HsAddin" (ByVal vtSheetName As Variant, ParamArray MemberList() As Variant) As Long
Private Declare PtrSafe Function HypSetPages Lib "HsAddin" (ByVal vtSheetName As Variant, ParamArray MemberList() As Variant) As Long
Private Declare PtrSafe Function HypSetSheetOption Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtItem As Variant, ByVal vtOption As Variant) As Long
Private Declare PtrSafe Function HypUIConnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtFriendlyName As Variant) As Long
Private Declare PtrSafe Function HypUndo Lib "HsAddin" (ByVal vtSheetName As Variant) As Long
Private Declare PtrSafe Function HypZoomIn Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelection As Variant, ByVal vtLevel As Variant, ByVal vtAcross As Variant) As Long
Private Declare PtrSafe Function HypZoomOut Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtSelection As Variant) As Long


'------------------------------------------------------------------------------
' JSON helpers
'------------------------------------------------------------------------------
Private Function JStr(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String, code As Long, parts() As String
    If IsNull(v) Or IsEmpty(v) Then JStr = "null": Exit Function
    If IsError(v) Then JStr = """#ERROR""": Exit Function
    s = CStr(v)
    If Len(s) = 0 Then JStr = """""": Exit Function
    ReDim parts(1 To Len(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34: parts(i) = "\"""
            Case 92: parts(i) = "\\"
            Case 10: parts(i) = "\n"
            Case 13: parts(i) = "\r"
            Case 9: parts(i) = "\t"
            Case 0 To 31: parts(i) = "\u" & Right$("000" & Hex$(code), 4)
            Case Else: parts(i) = ch
        End Select
    Next i
    JStr = """" & Join(parts, "") & """"
End Function

Private Function ArrayRank(ByRef v As Variant) As Long
    Dim r As Long, tmp As Long
    On Error GoTo Done
    Do
        r = r + 1
        tmp = UBound(v, r)
    Loop
Done:
    ArrayRank = r - 1
End Function

Private Function JArr(ByRef v As Variant) As String
    Dim rank As Long, i As Long, j As Long, parts() As String, cells() As String
    rank = ArrayRank(v)
    If rank = 0 Then JArr = "[]": Exit Function
    If rank = 1 Then
        If UBound(v) < LBound(v) Then JArr = "[]": Exit Function
        ReDim parts(LBound(v) To UBound(v))
        For i = LBound(v) To UBound(v)
            parts(i) = JVal(v(i))
        Next i
        JArr = "[" & Join(parts, ",") & "]"
    Else
        ReDim parts(LBound(v, 1) To UBound(v, 1))
        For i = LBound(v, 1) To UBound(v, 1)
            ReDim cells(LBound(v, 2) To UBound(v, 2))
            For j = LBound(v, 2) To UBound(v, 2)
                cells(j) = JVal(v(i, j))
            Next j
            parts(i) = "[" & Join(cells, ",") & "]"
        Next i
        JArr = "[" & Join(parts, ",") & "]"
    End If
End Function

Private Function JVal(ByVal v As Variant) As String
    If IsArray(v) Then JVal = JArr(v): Exit Function
    Select Case VarType(v)
        Case vbEmpty, vbNull: JVal = "null"
        Case vbBoolean: JVal = IIf(v, "true", "false")
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JVal = Trim$(Str$(v))
        Case vbError: JVal = """#ERROR"""
        Case Else: JVal = JStr(v)
    End Select
End Function

Private Function JObj(ParamArray kv() As Variant) As String
    ' JObj("key1", jsonValue1, "key2", jsonValue2, ...) - values must already be JSON.
    Dim i As Long, parts() As String, n As Long
    If UBound(kv) < 0 Then JObj = "{}": Exit Function
    ReDim parts(0 To (UBound(kv) - 1) \ 2)
    For i = 0 To UBound(kv) - 1 Step 2
        parts(n) = JStr(kv(i)) & ":" & kv(i + 1)
        n = n + 1
    Next i
    JObj = "{" & Join(parts, ",") & "}"
End Function

'------------------------------------------------------------------------------
' Result envelope
'------------------------------------------------------------------------------
Private Function RcText(ByVal rc As Long) As String
    Dim txt As Variant
    On Error GoTo NoText
    If HypGetErrorInfo(rc, txt) = 0 Then RcText = CStr(txt)
    Exit Function
NoText:
    RcText = ""
End Function

Private Function LastErrorJson() As String
    Dim code As Variant, msg As Variant, desc As Variant
    On Error Resume Next
    HypGetLastError code, msg, desc
    On Error GoTo 0
    LastErrorJson = JObj("code", JVal(code), "message", JVal(msg), "description", JVal(desc))
End Function

Private Function Result(ByVal rc As Long, Optional ByVal dataJson As String = "null") As String
    Dim errJson As String
    If rc = 0 Then errJson = "null" Else errJson = LastErrorJson()
    Result = JObj("ok", IIf(rc = 0, "true", "false"), "rc", CStr(rc), _
                  "rcText", JStr(IIf(rc = 0, "", RcText(rc))), "error", errJson, "data", dataJson)
End Function

Private Function Okay(ByVal dataJson As String) As String
    Okay = JObj("ok", "true", "rc", "0", "rcText", """""", "error", "null", "data", dataJson)
End Function

Private Function Refused(ByVal reason As String) As String
    Refused = JObj("ok", "false", "rc", "null", "rcText", """""", _
                   "error", JObj("bridge", JStr(reason)), "data", "null")
End Function

Private Function VbaError() As String
    VbaError = JObj("ok", "false", "rc", "null", "rcText", """""", _
                    "error", JObj("vbaNumber", CStr(Err.Number), "vba", JStr(Err.Description)), _
                    "data", "null")
End Function

'------------------------------------------------------------------------------
' Workbook and sheet resolution
'------------------------------------------------------------------------------
Private Function ScratchWorkbook(ByVal createIfMissing As Boolean) As Workbook
    Dim wb As Workbook, p As Object
    For Each wb In Application.Workbooks
        For Each p In wb.CustomDocumentProperties
            If p.Name = SCRATCH_TAG Then
                Set ScratchWorkbook = wb
                Exit Function
            End If
        Next p
    Next wb
    If createIfMissing Then
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        wb.CustomDocumentProperties.Add Name:=SCRATCH_TAG, LinkToContent:=False, Type:=4, Value:="1"
        Set ScratchWorkbook = wb
    End If
End Function

Private Function IsScratch(ByVal wb As Workbook) As Boolean
    Dim p As Object
    For Each p In wb.CustomDocumentProperties
        If p.Name = SCRATCH_TAG Then IsScratch = True: Exit Function
    Next p
End Function

Private Function ResolveSheet(ByVal wbName As Variant, ByVal shName As Variant) As Object
    Dim wb As Workbook, sh As Object
    If Len(CStr(wbName)) = 0 Then
        Set wb = ScratchWorkbook(True)
    Else
        Set wb = Application.Workbooks(CStr(wbName))
    End If
    If Len(CStr(shName)) = 0 Then
        Set sh = wb.ActiveSheet
    Else
        Set sh = wb.Sheets(CStr(shName))
    End If
    wb.Activate
    sh.Activate
    Set ResolveSheet = sh
End Function

Private Function SplitList(ByVal s As Variant) As Variant
    ' "a|b|c" -> array. Empty string -> empty array.
    If Len(CStr(s)) = 0 Then SplitList = Array() Else SplitList = Split(CStr(s), "|")
End Function

Private Function ContentTypeName(ByVal t As Long) As String
    Select Case t
        Case 0: ContentTypeName = "empty"
        Case 1: ContentTypeName = "adhoc"
        Case 2: ContentTypeName = "form"
        Case 3: ContentTypeName = "interactive-report"
        Case Else: ContentTypeName = CStr(t)
    End Select
End Function

'------------------------------------------------------------------------------
' Diagnostics and scratch workbook
'------------------------------------------------------------------------------
Public Function SVB_Ping() As String
    On Error GoTo Fail
    Dim wb As Workbook, scratchName As String, activeName As String
    Set wb = ScratchWorkbook(False)
    If Not wb Is Nothing Then scratchName = wb.Name
    If Not ActiveWorkbook Is Nothing Then activeName = ActiveWorkbook.Name
    SVB_Ping = Okay(JObj("bridge", JStr(BRIDGE_VERSION), "excel", JStr(Application.Version), _
                         "activeWorkbook", JStr(activeName), "scratchWorkbook", JStr(scratchName)))
    Exit Function
Fail:
    SVB_Ping = VbaError()
End Function

Public Function SVB_Workbooks() As String
    ' Open workbooks and their sheets (no activation, nothing changes).
    On Error GoTo Fail
    Dim wb As Workbook, sh As Object, parts() As String, sheets() As String, i As Long, j As Long
    If Application.Workbooks.Count = 0 Then SVB_Workbooks = Okay("[]"): Exit Function
    ReDim parts(1 To Application.Workbooks.Count)
    For Each wb In Application.Workbooks
        i = i + 1
        ReDim sheets(1 To wb.Sheets.Count)
        j = 0
        For Each sh In wb.Sheets
            j = j + 1
            sheets(j) = JStr(sh.Name)
        Next sh
        parts(i) = JObj("name", JStr(wb.Name), "scratch", JVal(IsScratch(wb)), _
                        "saved", JVal(wb.Saved), "sheets", "[" & Join(sheets, ",") & "]")
    Next wb
    SVB_Workbooks = Okay("[" & Join(parts, ",") & "]")
    Exit Function
Fail:
    SVB_Workbooks = VbaError()
End Function

Public Function SVB_LastError() As String
    SVB_LastError = Okay(LastErrorJson())
End Function

Public Function SVB_Scratch(Optional ByVal shName As Variant = "", Optional ByVal clearSheet As Variant = False) As String
    ' Ensure the scratch workbook (and a sheet named shName) exists. clearSheet=True
    ' deletes and recreates that sheet, which also drops its Smart View metadata.
    On Error GoTo Fail
    Dim wb As Workbook, sh As Worksheet, alerts As Boolean
    Set wb = ScratchWorkbook(True)
    If Len(CStr(shName)) > 0 Then
        On Error Resume Next
        Set sh = wb.Worksheets(CStr(shName))
        On Error GoTo Fail
        If Not sh Is Nothing And CBool(clearSheet) Then
            If wb.Worksheets.Count = 1 Then wb.Worksheets.Add After:=sh
            alerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            sh.Delete
            Application.DisplayAlerts = alerts
            Set sh = Nothing
        End If
        If sh Is Nothing Then
            Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            sh.Name = CStr(shName)
        End If
        wb.Activate
        sh.Activate
    End If
    SVB_Scratch = Okay(JObj("workbook", JStr(wb.Name), "sheet", JStr(IIf(sh Is Nothing, "", shName))))
    Exit Function
Fail:
    Application.DisplayAlerts = True
    SVB_Scratch = VbaError()
End Function

Public Function SVB_ReadTsv(Optional ByVal wbName As Variant = "", Optional ByVal shName As Variant = "", _
                            Optional ByVal address As Variant = "", Optional ByVal maxRows As Variant = 2000) As String
    ' Values as TSV (not JSON). Column 0 is the indent level of the first column,
    ' which is how Smart View shows hierarchy when Excel indentation is used.
    On Error GoTo Fail
    Dim sh As Object, rng As Range, v As Variant, n As Long, c As Long, r As Long
    Dim lines() As String, cells() As String, s As String
    Set sh = ResolveSheet(wbName, shName)
    If Len(CStr(address)) = 0 Then Set rng = sh.UsedRange Else Set rng = sh.Range(CStr(address))
    n = rng.Rows.Count
    If n > CLng(maxRows) Then n = CLng(maxRows)
    Set rng = rng.Resize(n)
    v = rng.Value2
    ReDim lines(1 To n)
    If Not IsArray(v) Then
        lines(1) = rng.Cells(1, 1).IndentLevel & vbTab & CStr(v)
    Else
        For r = 1 To n
            ReDim cells(0 To UBound(v, 2))
            cells(0) = CStr(rng.Cells(r, 1).IndentLevel)
            For c = 1 To UBound(v, 2)
                If IsError(v(r, c)) Then
                    s = "#ERROR"
                Else
                    s = CStr(v(r, c))
                End If
                s = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
                cells(c) = s
            Next c
            lines(r) = Join(cells, vbTab)
        Next r
    End If
    SVB_ReadTsv = "#range" & vbTab & rng.Address(False, False) & vbTab & "#sheet" & vbTab & sh.Name & _
                  vbTab & "#workbook" & vbTab & sh.Parent.Name & vbLf & Join(lines, vbLf)
    Exit Function
Fail:
    SVB_ReadTsv = VbaError()
End Function

Public Function SVB_WriteTsv(ByVal shName As Variant, ByVal topLeft As Variant, ByVal tsv As Variant) As String
    ' Writes labels into a sheet of the SCRATCH workbook only (to build ad hoc grids).
    ' Rows separated by LF, cells by TAB. Prefix numeric member names with an apostrophe.
    On Error GoTo Fail
    Dim wb As Workbook, sh As Worksheet, rows As Variant, cols As Variant
    Dim r As Long, c As Long, maxC As Long, arr() As Variant
    If InStr(1, CStr(tsv), "HsSetValue", vbTextCompare) > 0 Or InStr(1, CStr(tsv), "HsSetText", vbTextCompare) > 0 Then
        SVB_WriteTsv = Refused("HsSetValue and HsSetText write to the server on Submit; not allowed.")
        Exit Function
    End If
    Set wb = ScratchWorkbook(True)
    Set sh = wb.Worksheets(CStr(shName))
    rows = Split(Replace(CStr(tsv), vbCr, ""), vbLf)
    For r = 0 To UBound(rows)
        cols = Split(rows(r), vbTab)
        If UBound(cols) + 1 > maxC Then maxC = UBound(cols) + 1
    Next r
    If maxC = 0 Then SVB_WriteTsv = Refused("Nothing to write."): Exit Function
    ReDim arr(1 To UBound(rows) + 1, 1 To maxC)
    For r = 0 To UBound(rows)
        cols = Split(rows(r), vbTab)
        For c = 0 To UBound(cols)
            arr(r + 1, c + 1) = cols(c)
        Next c
    Next r
    Dim fx As Variant
    fx = arr
    For r = 1 To UBound(rows) + 1
        For c = 1 To maxC
            If Left$(CStr(arr(r, c)), 1) = "=" Then arr(r, c) = Empty
        Next c
    Next r
    sh.Range(CStr(topLeft)).Resize(UBound(rows) + 1, maxC).Value = arr
    For r = 1 To UBound(rows) + 1
        For c = 1 To maxC
            If Left$(CStr(fx(r, c)), 1) = "=" Then sh.Range(CStr(topLeft)).Cells(r, c).Formula = fx(r, c)
        Next c
    Next r
    SVB_WriteTsv = Okay(JObj("workbook", JStr(wb.Name), "sheet", JStr(sh.Name), _
                             "range", JStr(sh.Range(CStr(topLeft)).Resize(UBound(rows) + 1, maxC).Address(False, False))))
    Exit Function
Fail:
    SVB_WriteTsv = VbaError()
End Function

'------------------------------------------------------------------------------
' Connections (no passwords, ever)
'------------------------------------------------------------------------------
Public Function SVB_ConnectionExists(ByVal friendlyName As Variant) As String
    On Error GoTo Fail
    SVB_ConnectionExists = Okay(JVal(HypConnectionExists(friendlyName)))
    Exit Function
Fail:
    SVB_ConnectionExists = VbaError()
End Function

Public Function SVB_CreateConnection(ByVal providerType As Variant, ByVal providerUrl As Variant, _
                                     ByVal serverName As Variant, ByVal appName As Variant, _
                                     ByVal cubeName As Variant, ByVal friendlyName As Variant, _
                                     Optional ByVal description As Variant = "") As String
    ' Creates a private connection without credentials.
    On Error GoTo Fail
    SVB_CreateConnection = Result(HypCreateConnectionEx(providerType, serverName, appName, cubeName, "", _
                                  providerUrl, friendlyName, "", "", description, "", ""))
    Exit Function
Fail:
    SVB_CreateConnection = VbaError()
End Function

Public Function SVB_Connect(ByVal wbName As Variant, ByVal shName As Variant, ByVal friendlyName As Variant, _
                            Optional ByVal userName As Variant = "") As String
    ' HypConnect: on-premises sources only (Oracle documents it as not working
    ' against cloud sources). For EPM Cloud use SVB_UIConnect.
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Connect = Result(HypConnect(sh.Name, userName, "", friendlyName))
    Exit Function
Fail:
    SVB_Connect = VbaError()
End Function

' Cloud and on-premises. friendlyName may also be an inline "URL|server|app|cube".
' For cloud sources Smart View shows its own sign-in; the user types the password there.
Public Function SVB_UIConnect(ByVal wbName As Variant, ByVal shName As Variant, ByVal friendlyName As Variant, _
                              Optional ByVal userName As Variant = "") As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_UIConnect = Result(HypUIConnect(sh.Name, userName, "", friendlyName))
    Exit Function
Fail:
    SVB_UIConnect = VbaError()
End Function

Public Function SVB_Disconnect(ByVal wbName As Variant, ByVal shName As Variant, _
                               Optional ByVal logoutUser As Variant = False) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Disconnect = Result(HypDisconnect(sh.Name, CBool(logoutUser)))
    Exit Function
Fail:
    SVB_Disconnect = VbaError()
End Function

Public Function SVB_Connected(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Connected = Okay(JVal(HypConnected(sh.Name)))
    Exit Function
Fail:
    SVB_Connected = VbaError()
End Function

Public Function SVB_ConnectionInfo() As String
    ' Connection of the active sheet. The password slot is discarded unread.
    On Error GoTo Fail
    Dim srv As Variant, usr As Variant, pwd As Variant, app As Variant, db As Variant
    Dim fn As Variant, url As Variant, prov As Variant, rc As Long
    rc = HypGetConnectionInfo(srv, usr, pwd, app, db, fn, url, prov)
    pwd = Empty
    SVB_ConnectionInfo = Result(rc, JObj("server", JVal(srv), "user", JVal(usr), "application", JVal(app), _
                                         "cube", JVal(db), "friendlyName", JVal(fn), "url", JVal(url), _
                                         "provider", JVal(prov)))
    Exit Function
Fail:
    SVB_ConnectionInfo = VbaError()
End Function

Public Function SVB_SetActiveConnection(ByVal friendlyName As Variant) As String
    On Error GoTo Fail
    SVB_SetActiveConnection = Result(HypSetActiveConnection(friendlyName))
    Exit Function
Fail:
    SVB_SetActiveConnection = VbaError()
End Function

Public Function SVB_SetConnAliasTable(ByVal friendlyName As Variant, ByVal aliasTable As Variant) As String
    On Error GoTo Fail
    SVB_SetConnAliasTable = Result(HypSetConnAliasTable(friendlyName, aliasTable))
    Exit Function
Fail:
    SVB_SetConnAliasTable = VbaError()
End Function

Public Function SVB_SetAliasTable(ByVal wbName As Variant, ByVal shName As Variant, ByVal aliasTable As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_SetAliasTable = Result(HypSetAliasTable(sh.Name, aliasTable))
    Exit Function
Fail:
    SVB_SetAliasTable = VbaError()
End Function

Public Function SVB_ListUserVariables(ByVal friendlyName As Variant) As String
    On Error GoTo Fail
    Dim dims As Variant, vars As Variant, mbrs As Variant, rc As Long
    rc = HypListUserVariables(friendlyName, dims, vars, mbrs)
    SVB_ListUserVariables = Result(rc, JObj("dimensions", JVal(dims), "variables", JVal(vars), "members", JVal(mbrs)))
    Exit Function
Fail:
    SVB_ListUserVariables = VbaError()
End Function

Public Function SVB_SubstitutionVariables(ByVal wbName As Variant, ByVal shName As Variant, _
                                          Optional ByVal appName As Variant = "", Optional ByVal cubeName As Variant = "", _
                                          Optional ByVal variableName As Variant = "") As String
    On Error GoTo Fail
    Dim sh As Object, names As Variant, values As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetSubstitutionVariable(sh.Name, appName, cubeName, variableName, names, values)
    SVB_SubstitutionVariables = Result(rc, JObj("names", JVal(names), "values", JVal(values)))
    Exit Function
Fail:
    SVB_SubstitutionVariables = VbaError()
End Function

'------------------------------------------------------------------------------
' Sheet content and ad hoc grid operations
'------------------------------------------------------------------------------
Public Function SVB_ContentType(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, t As Long, present As Boolean
    Set sh = ResolveSheet(wbName, shName)
    present = HypIsSmartViewContentPresent(sh.Name, t)
    SVB_ContentType = Okay(JObj("smartViewContent", JVal(present), "type", JStr(ContentTypeName(CLng(t)))))
    Exit Function
Fail:
    SVB_ContentType = VbaError()
End Function

Public Function SVB_SheetInfo(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, names As Variant, values As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetSheetInfo(sh.Name, names, values)
    SVB_SheetInfo = Result(rc, JObj("names", JVal(names), "values", JVal(values)))
    Exit Function
Fail:
    SVB_SheetInfo = VbaError()
End Function

Public Function SVB_Retrieve(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Retrieve = Result(HypRetrieve(sh.Name))
    Exit Function
Fail:
    SVB_Retrieve = VbaError()
End Function

Public Function SVB_RetrieveRange(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant, _
                                  ByVal friendlyName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_RetrieveRange = Result(HypRetrieveRange(sh.Name, sh.Range(CStr(address)), friendlyName))
    Exit Function
Fail:
    SVB_RetrieveRange = VbaError()
End Function

Public Function SVB_RefreshAll(ByVal wbName As Variant) As String
    ' Smart View "Refresh All" on one workbook (all its connected sheets).
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, "")
    SVB_RefreshAll = Result(HypMenuVRefreshAll())
    Exit Function
Fail:
    SVB_RefreshAll = VbaError()
End Function

Public Function SVB_ZoomIn(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant, _
                           Optional ByVal level As Variant = 1) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_ZoomIn = Result(HypZoomIn(sh.Name, sh.Range(CStr(address)), CLng(level), False))
    Exit Function
Fail:
    SVB_ZoomIn = VbaError()
End Function

Public Function SVB_ZoomOut(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_ZoomOut = Result(HypZoomOut(sh.Name, sh.Range(CStr(address))))
    Exit Function
Fail:
    SVB_ZoomOut = VbaError()
End Function

Public Function SVB_KeepOnly(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_KeepOnly = Result(HypKeepOnly(sh.Name, sh.Range(CStr(address))))
    Exit Function
Fail:
    SVB_KeepOnly = VbaError()
End Function

Public Function SVB_RemoveOnly(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_RemoveOnly = Result(HypRemoveOnly(sh.Name, sh.Range(CStr(address))))
    Exit Function
Fail:
    SVB_RemoveOnly = VbaError()
End Function

Public Function SVB_Pivot(ByVal wbName As Variant, ByVal shName As Variant, ByVal startAddress As Variant, _
                          ByVal endAddress As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Pivot = Result(HypPivot(sh.Name, sh.Range(CStr(startAddress)), sh.Range(CStr(endAddress))))
    Exit Function
Fail:
    SVB_Pivot = VbaError()
End Function

Public Function SVB_PivotToPOV(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_PivotToPOV = Result(HypPivotToPOV(sh.Name, sh.Range(CStr(address))))
    Exit Function
Fail:
    SVB_PivotToPOV = VbaError()
End Function

Public Function SVB_PivotToGrid(ByVal wbName As Variant, ByVal shName As Variant, ByVal dimensionName As Variant, _
                                ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_PivotToGrid = Result(HypPivotToGrid(sh.Name, dimensionName, sh.Range(CStr(address))))
    Exit Function
Fail:
    SVB_PivotToGrid = VbaError()
End Function

Public Function SVB_Undo(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Undo = Result(HypUndo(sh.Name))
    Exit Function
Fail:
    SVB_Undo = VbaError()
End Function

Public Function SVB_Redo(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Redo = Result(HypRedo(sh.Name))
    Exit Function
Fail:
    SVB_Redo = VbaError()
End Function

' HypGetPOV is exported by the Smart View DLL but not documented by Oracle.
Public Function SVB_GetPOV(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, dims As Variant, mbrs As Variant, types As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetPOV(sh.Name, dims, mbrs, types)
    SVB_GetPOV = Result(rc, JObj("dimensions", JVal(dims), "members", JVal(mbrs), "types", JVal(types)))
    Exit Function
Fail:
    SVB_GetPOV = VbaError()
End Function

Public Function SVB_SetPOV(ByVal wbName As Variant, ByVal shName As Variant, ByVal members As Variant) As String
    ' members: items separated by "|", each in the format HypSetPOV expects
    ' (see references/vba-api.md). Up to 16 items.
    On Error GoTo Fail
    Dim sh As Object, m As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    m = SplitList(members)
    Select Case UBound(m) + 1
        Case 1: rc = HypSetPOV(sh.Name, m(0))
        Case 2: rc = HypSetPOV(sh.Name, m(0), m(1))
        Case 3: rc = HypSetPOV(sh.Name, m(0), m(1), m(2))
        Case 4: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3))
        Case 5: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4))
        Case 6: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5))
        Case 7: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6))
        Case 8: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7))
        Case 9: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8))
        Case 10: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9))
        Case 11: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10))
        Case 12: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11))
        Case 13: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12))
        Case 14: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12), m(13))
        Case 15: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12), m(13), m(14))
        Case 16: rc = HypSetPOV(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12), m(13), m(14), m(15))
        Case Else
            SVB_SetPOV = Refused("Pass between 1 and 16 POV items separated by |.")
            Exit Function
    End Select
    SVB_SetPOV = Result(rc)
    Exit Function
Fail:
    SVB_SetPOV = VbaError()
End Function

Public Function SVB_SetPages(ByVal wbName As Variant, ByVal shName As Variant, ByVal members As Variant) As String
    ' Page (POV) members of a FORM, "Dimension#Member" items separated by "|". Up to 8.
    On Error GoTo Fail
    Dim sh As Object, m As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    m = SplitList(members)
    Select Case UBound(m) + 1
        Case 1: rc = HypSetPages(sh.Name, m(0))
        Case 2: rc = HypSetPages(sh.Name, m(0), m(1))
        Case 3: rc = HypSetPages(sh.Name, m(0), m(1), m(2))
        Case 4: rc = HypSetPages(sh.Name, m(0), m(1), m(2), m(3))
        Case 5: rc = HypSetPages(sh.Name, m(0), m(1), m(2), m(3), m(4))
        Case 6: rc = HypSetPages(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5))
        Case 7: rc = HypSetPages(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6))
        Case 8: rc = HypSetPages(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7))
        Case Else
            SVB_SetPages = Refused("Pass between 1 and 8 page items separated by |.")
            Exit Function
    End Select
    SVB_SetPages = Result(rc)
    Exit Function
Fail:
    SVB_SetPages = VbaError()
End Function

Public Function SVB_BackgroundPOV(ByVal friendlyName As Variant) As String
    ' POV stored for a connection (POV Manager).
    On Error GoTo Fail
    Dim dims As Variant, mbrs As Variant, rc As Long
    rc = HypGetBackgroundPOV(friendlyName, dims, mbrs)
    SVB_BackgroundPOV = Result(rc, JObj("dimensions", JVal(dims), "members", JVal(mbrs)))
    Exit Function
Fail:
    SVB_BackgroundPOV = VbaError()
End Function

Public Function SVB_GetDimensions(ByVal wbName As Variant, ByVal shName As Variant) As String
    ' Dimensions of the sheet's grid and where each sits (row, column, POV, page).
    On Error GoTo Fail
    Dim sh As Object, names As Variant, types As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetDimensions(sh.Name, names, types)
    SVB_GetDimensions = Result(rc, JObj("dimensions", JVal(names), "types", JVal(types)))
    Exit Function
Fail:
    SVB_GetDimensions = VbaError()
End Function

Public Function SVB_GetMembers(ByVal wbName As Variant, ByVal shName As Variant, ByVal dimensionName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, names As Variant, descs As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetMembers(sh.Name, dimensionName, names, descs)
    SVB_GetMembers = Result(rc, JObj("names", JVal(names), "descriptions", JVal(descs)))
    Exit Function
Fail:
    SVB_GetMembers = VbaError()
End Function

Public Function SVB_PagePOVChoices(ByVal wbName As Variant, ByVal shName As Variant, ByVal dimensionName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, names As Variant, descs As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetPagePOVChoices(sh.Name, dimensionName, names, descs)
    SVB_PagePOVChoices = Result(rc, JObj("names", JVal(names), "descriptions", JVal(descs)))
    Exit Function
Fail:
    SVB_PagePOVChoices = VbaError()
End Function

Public Function SVB_CellPOV(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    ' Full member intersection behind one data cell.
    On Error GoTo Fail
    Dim sh As Object, srv As Variant, app As Variant, cube As Variant, frm As Variant
    Dim dims As Variant, mbrs As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetDimMbrsForDataCell(sh.Name, sh.Range(CStr(address)), srv, app, cube, frm, dims, mbrs)
    SVB_CellPOV = Result(rc, JObj("server", JVal(srv), "application", JVal(app), "cube", JVal(cube), _
                                  "form", JVal(frm), "dimensions", JVal(dims), "members", JVal(mbrs)))
    Exit Function
Fail:
    SVB_CellPOV = VbaError()
End Function

Public Function SVB_IsCellWritable(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_IsCellWritable = Okay(JVal(HypIsCellWritable(sh.Name, sh.Range(CStr(address)))))
    Exit Function
Fail:
    SVB_IsCellWritable = VbaError()
End Function

Public Function SVB_ExecuteQuery(ByVal wbName As Variant, ByVal shName As Variant, ByVal mdx As Variant) As String
    ' Runs an MDX SELECT and places the result on the sheet (provider support varies).
    On Error GoTo Fail
    Dim sh As Object
    If Not UCase$(LTrim$(CStr(mdx))) Like "SELECT*" And Not UCase$(LTrim$(CStr(mdx))) Like "WITH*" Then
        SVB_ExecuteQuery = Refused("Only MDX SELECT queries are allowed.")
        Exit Function
    End If
    Set sh = ResolveSheet(wbName, shName)
    SVB_ExecuteQuery = Result(HypExecuteQuery(sh.Name, mdx))
    Exit Function
Fail:
    SVB_ExecuteQuery = VbaError()
End Function

Public Function SVB_DatabaseNote(ByVal wbName As Variant, ByVal shName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, note As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetDatabaseNote(sh.Name, note)
    SVB_DatabaseNote = Result(rc, JVal(note))
    Exit Function
Fail:
    SVB_DatabaseNote = VbaError()
End Function

'------------------------------------------------------------------------------
' Forms, documents and drill-through
'------------------------------------------------------------------------------
Public Function SVB_ListDocuments(ByVal wbName As Variant, ByVal shName As Variant, _
                                  ByVal connInfo As Variant, Optional ByVal folderPath As Variant = "") As String
    On Error GoTo Fail
    Dim sh As Object, docs As SvbDocInfo, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypListDocuments(sh.Name, "", "", connInfo, folderPath, docs)
    SVB_ListDocuments = Result(rc, JObj("count", CStr(docs.numDocs), "types", JVal(docs.docTypes), _
                                        "names", JVal(docs.docNames), "descriptions", JVal(docs.docDescriptions), _
                                        "planTypes", JVal(docs.docPlanTypes), "attributes", JVal(docs.docAttributes)))
    Exit Function
Fail:
    SVB_ListDocuments = VbaError()
End Function

Public Function SVB_OpenForm(ByVal wbName As Variant, ByVal shName As Variant, ByVal folderPath As Variant, _
                             ByVal formName As Variant) As String
    ' Opens a Planning form on the sheet. Read it; never type into its data cells.
    On Error GoTo Fail
    Dim sh As Object, dims() As Variant, mbrs() As Variant
    Set sh = ResolveSheet(wbName, shName)
    ReDim dims(0): ReDim mbrs(0)
    SVB_OpenForm = Result(HypOpenForm(sh.Name, folderPath, formName, dims, mbrs))
    Exit Function
Fail:
    SVB_OpenForm = VbaError()
End Function

Public Function SVB_AdhocOnForm(ByVal wbName As Variant, ByVal shName As Variant, ByVal formName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_AdhocOnForm = Result(HypPerformAdhocOnForm(sh.Name, formName))
    Exit Function
Fail:
    SVB_AdhocOnForm = VbaError()
End Function

Public Function SVB_DrillThroughReports(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, ids As Variant, names As Variant, urls As Variant, tpls As Variant, types As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetDrillThroughReports(sh.Name, sh.Range(CStr(address)), ids, names, urls, tpls, types)
    SVB_DrillThroughReports = Result(rc, JObj("ids", JVal(ids), "names", JVal(names), "urls", JVal(urls), _
                                              "urlTemplates", JVal(tpls), "types", JVal(types)))
    Exit Function
Fail:
    SVB_DrillThroughReports = VbaError()
End Function

Public Function SVB_ExecuteDrillThrough(ByVal wbName As Variant, ByVal shName As Variant, ByVal address As Variant, _
                                        Optional ByVal index As Variant = 0) As String
    ' Runs the drill-through report number "index" (0-based) listed for that cell.
    On Error GoTo Fail
    Dim sh As Object, ids As Variant, names As Variant, urls As Variant, tpls As Variant, types As Variant
    Dim rc As Long, i As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetDrillThroughReports(sh.Name, sh.Range(CStr(address)), ids, names, urls, tpls, types)
    If rc <> 0 Then SVB_ExecuteDrillThrough = Result(rc): Exit Function
    If Not IsArray(ids) Then SVB_ExecuteDrillThrough = Refused("No drill-through report for that cell."): Exit Function
    i = LBound(ids) + CLng(index)
    If i > UBound(ids) Then SVB_ExecuteDrillThrough = Refused("Index out of range."): Exit Function
    rc = HypExecuteDrillThroughReport(sh.Name, sh.Range(CStr(address)), ids(i), names(i), urls(i), tpls(i), types(i))
    SVB_ExecuteDrillThrough = Result(rc, JObj("report", JVal(names(i))))
    Exit Function
Fail:
    SVB_ExecuteDrillThrough = VbaError()
End Function

'------------------------------------------------------------------------------
' Options (prefer sheet options; restore global ones after use)
'------------------------------------------------------------------------------
Public Function SVB_GetSheetOption(ByVal wbName As Variant, ByVal shName As Variant, ByVal item As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_GetSheetOption = Okay(JVal(HypGetSheetOption(sh.Name, CLng(item))))
    Exit Function
Fail:
    SVB_GetSheetOption = VbaError()
End Function

Public Function SVB_SetSheetOption(ByVal wbName As Variant, ByVal shName As Variant, ByVal item As Variant, _
                                   ByVal value As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_SetSheetOption = Result(HypSetSheetOption(sh.Name, CLng(item), CoerceOption(value)))
    Exit Function
Fail:
    SVB_SetSheetOption = VbaError()
End Function

Public Function SVB_GetGlobalOption(ByVal item As Variant) As String
    On Error GoTo Fail
    SVB_GetGlobalOption = Okay(JVal(HypGetGlobalOption(CLng(item))))
    Exit Function
Fail:
    SVB_GetGlobalOption = VbaError()
End Function

Public Function SVB_SetGlobalOption(ByVal item As Variant, ByVal value As Variant) As String
    On Error GoTo Fail
    SVB_SetGlobalOption = Result(HypSetGlobalOption(CLng(item), CoerceOption(value)))
    Exit Function
Fail:
    SVB_SetGlobalOption = VbaError()
End Function

Public Function SVB_GetOption(ByVal item As Variant, Optional ByVal wbName As Variant = "", _
                              Optional ByVal shName As Variant = "") As String
    ' HypGetOption: sheet option when a sheet is given, otherwise default (global) option.
    On Error GoTo Fail
    Dim ret As Variant, rc As Long, sheetArg As Variant
    If Len(CStr(shName)) > 0 Or Len(CStr(wbName)) > 0 Then sheetArg = ResolveSheet(wbName, shName).Name Else sheetArg = ""
    rc = HypGetOption(CLng(item), ret, sheetArg)
    SVB_GetOption = Result(rc, JVal(ret))
    Exit Function
Fail:
    SVB_GetOption = VbaError()
End Function

Public Function SVB_SetOption(ByVal item As Variant, ByVal value As Variant, Optional ByVal wbName As Variant = "", _
                              Optional ByVal shName As Variant = "") As String
    ' Unified option API (HSV_* indices). With a sheet: that sheet only. Without: the
    ' user's default options, which persist; read and restore them.
    On Error GoTo Fail
    Dim sheetArg As Variant
    If Len(CStr(shName)) > 0 Or Len(CStr(wbName)) > 0 Then sheetArg = ResolveSheet(wbName, shName).Name Else sheetArg = ""
    SVB_SetOption = Result(HypSetOption(CLng(item), CoerceOption(value), sheetArg))
    Exit Function
Fail:
    SVB_SetOption = VbaError()
End Function

Private Function CoerceOption(ByVal value As Variant) As Variant
    ' Command-line values arrive as text: map true/false and numbers.
    Select Case LCase$(CStr(value))
        Case "true": CoerceOption = True
        Case "false": CoerceOption = False
        Case Else
            If IsNumeric(value) And InStr(CStr(value), ".") = 0 And InStr(CStr(value), ",") = 0 Then
                CoerceOption = CLng(value)
            Else
                CoerceOption = CStr(value)
            End If
    End Select
End Function

'------------------------------------------------------------------------------
' Cell functions (HsGetValue and friends)
'------------------------------------------------------------------------------
Public Function SVB_Refresh(ByVal wbName As Variant, ByVal shName As Variant) As String
    ' Smart View Refresh of one sheet, like the ribbon button: refreshes its grid
    ' and runs its Hs* functions (HsGetValue cells are only computed this way).
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_Refresh = Result(HypMenuVRefresh())
    Exit Function
Fail:
    SVB_Refresh = VbaError()
End Function

Public Function SVB_ReadFormulas(Optional ByVal wbName As Variant = "", Optional ByVal shName As Variant = "", _
                                 Optional ByVal address As Variant = "", Optional ByVal maxCells As Variant = 2000) As String
    ' Formula cells as TSV (not JSON): address, formula (English syntax), shown text.
    On Error GoTo Fail
    Dim sh As Object, rng As Range, cel As Range, lines() As String, n As Long, cap As Long
    Set sh = ResolveSheet(wbName, shName)
    If Len(CStr(address)) = 0 Then Set rng = sh.UsedRange Else Set rng = sh.Range(CStr(address))
    cap = CLng(maxCells)
    ReDim lines(0 To cap)
    lines(0) = "#sheet" & vbTab & sh.Name & vbTab & "#workbook" & vbTab & sh.Parent.Name
    On Error Resume Next
    ' Intersect: on a single cell, SpecialCells would scan the whole sheet.
    Set rng = Application.Intersect(rng, rng.SpecialCells(xlCellTypeFormulas))
    If Err.Number <> 0 Or rng Is Nothing Then
        On Error GoTo Fail
        SVB_ReadFormulas = lines(0) & vbLf & "#none"
        Exit Function
    End If
    On Error GoTo Fail
    For Each cel In rng.Cells
        If n >= cap Then Exit For
        n = n + 1
        lines(n) = cel.Address(False, False) & vbTab & OneLine(cel.Formula) & vbTab & OneLine(cel.Text)
    Next cel
    ReDim Preserve lines(0 To n)
    SVB_ReadFormulas = Join(lines, vbLf)
    Exit Function
Fail:
    SVB_ReadFormulas = VbaError()
End Function

Public Function SVB_AuditFunctions(Optional ByVal wbName As Variant = "", Optional ByVal shName As Variant = "", _
                                   Optional ByVal address As Variant = "", Optional ByVal maxCells As Variant = 2000) As String
    ' Every Smart View cell function (HsGetValue, HsSetValue, HsDescription...) in the
    ' range, with its arguments EVALUATED by Excel (so POVs built with & and cell
    ' references show the real members), the text the cell shows, and whether the
    ' formula carries a hard-coded path to the functions add-in (broken links).
    On Error GoTo Fail
    Dim sh As Object, rng As Range, cel As Range, calls As Collection, c As Variant
    Dim parts() As String, n As Long, cap As Long, cells As Long, truncated As Boolean
    Dim argsJson() As String, i As Long, f As String
    Set sh = ResolveSheet(wbName, shName)
    If Len(CStr(address)) = 0 Then Set rng = sh.UsedRange Else Set rng = sh.Range(CStr(address))
    cap = CLng(maxCells)
    On Error Resume Next
    ' Intersect: on a single cell, SpecialCells would scan the whole sheet.
    Set rng = Application.Intersect(rng, rng.SpecialCells(xlCellTypeFormulas))
    If Err.Number <> 0 Or rng Is Nothing Then
        On Error GoTo Fail
        SVB_AuditFunctions = Okay(JObj("sheet", JStr(sh.Name), "cellsScanned", "0", "calls", "[]", "truncated", "false"))
        Exit Function
    End If
    On Error GoTo Fail
    ReDim parts(0 To 0)
    For Each cel In rng.Cells
        f = cel.Formula
        If InStr(1, f, "Hs", vbTextCompare) > 0 Then
            Set calls = FindHsCalls(f)
            For Each c In calls
                If n >= cap Then truncated = True: Exit For
                ReDim argsJson(0 To UBound(c(1)))
                For i = 0 To UBound(c(1))
                    argsJson(i) = JStr(EvalArg(sh, c(1)(i)))
                Next i
                ReDim Preserve parts(0 To n)
                parts(n) = JObj("cell", JStr(cel.Address(False, False)), "fn", JStr(c(0)), _
                                "args", "[" & Join(argsJson, ",") & "]", "text", JStr(cel.Text), _
                                "addinPathInFormula", JVal(InStr(1, f, ".xla", vbTextCompare) > 0))
                n = n + 1
            Next c
        End If
        cells = cells + 1
        If truncated Then Exit For
    Next cel
    If n = 0 Then
        SVB_AuditFunctions = Okay(JObj("sheet", JStr(sh.Name), "cellsScanned", CStr(cells), "calls", "[]", "truncated", "false"))
    Else
        SVB_AuditFunctions = Okay(JObj("sheet", JStr(sh.Name), "cellsScanned", CStr(cells), _
                                       "calls", "[" & Join(parts, ",") & "]", "truncated", JVal(truncated)))
    End If
    Exit Function
Fail:
    SVB_AuditFunctions = VbaError()
End Function

Private Function OneLine(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then OneLine = "#ERROR": Exit Function
    s = CStr(v)
    OneLine = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Function EvalArg(ByVal sh As Object, ByVal arg As String) As String
    ' Evaluates one argument expression as Excel would (literals, & concatenation,
    ' cell references). Falls back to the raw text if it cannot be evaluated.
    Dim v As Variant
    If Len(Trim$(arg)) = 0 Then EvalArg = "": Exit Function
    On Error GoTo Raw
    v = sh.Evaluate(arg)
    If IsError(v) Then GoTo Raw
    If IsArray(v) Then GoTo Raw
    EvalArg = CStr(v)
    Exit Function
Raw:
    EvalArg = "=" & arg
End Function

Private Function IsHsName(ByVal nm As String) As Boolean
    Select Case LCase$(nm)
        Case "hsgetvalue", "hssetvalue", "hsdescription", "hsalias", "hsgetsheetinfo", "hsgetvariable", _
             "hscurrency", "hslabel", "hsgettext", "hssettext", "hsgetsharedconnectionurl"
            IsHsName = True
    End Select
End Function

Private Function FindHsCalls(ByVal f As String) As Collection
    ' Finds Hs*( calls outside string literals and splits their top-level arguments.
    ' Each item is Array(functionName, argsArray).
    Dim res As New Collection, i As Long, j As Long, k As Long, n As Long
    Dim inQ As Boolean, ch As String, prev As String, nm As String
    Dim depth As Long, inQ2 As Boolean, startArg As Long, args() As String, na As Long
    n = Len(f)
    i = 1
    Do While i <= n
        ch = Mid$(f, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf Not inQ And (ch = "H" Or ch = "h") Then
            If i = 1 Then prev = "" Else prev = Mid$(f, i - 1, 1)
            If Not (prev Like "[A-Za-z0-9_.]") Then
                j = i
                Do While j <= n
                    If Not (Mid$(f, j, 1) Like "[A-Za-z]") Then Exit Do
                    j = j + 1
                Loop
                nm = Mid$(f, i, j - i)
                If j <= n And IsHsName(nm) Then
                    If Mid$(f, j, 1) = "(" Then
                        depth = 1: inQ2 = False: startArg = j + 1: na = 0
                        ReDim args(0 To 0)
                        For k = j + 1 To n
                            ch = Mid$(f, k, 1)
                            If ch = """" Then
                                inQ2 = Not inQ2
                            ElseIf Not inQ2 Then
                                If ch = "(" Then
                                    depth = depth + 1
                                ElseIf ch = ")" Then
                                    depth = depth - 1
                                    If depth = 0 Then
                                        ReDim Preserve args(0 To na)
                                        args(na) = Mid$(f, startArg, k - startArg)
                                        Exit For
                                    End If
                                ElseIf ch = "," And depth = 1 Then
                                    ReDim Preserve args(0 To na)
                                    args(na) = Mid$(f, startArg, k - startArg)
                                    na = na + 1
                                    startArg = k + 1
                                End If
                            End If
                        Next k
                        res.Add Array(nm, args)
                        i = j
                    End If
                End If
            End If
        End If
        i = i + 1
    Loop
    Set FindHsCalls = res
End Function

Public Function SVB_CellValue(ByVal wbName As Variant, ByVal shName As Variant, ByVal members As Variant) As String
    ' Value of one intersection through the sheet's connection, read without
    ' writing to the sheet (like a one-off HsGetValue). members: "Dimension#Member"
    ' items separated by "|", up to 16. Returns data: only when the user asked for it.
    On Error GoTo Fail
    Dim sh As Object, m As Variant, v As Variant
    Set sh = ResolveSheet(wbName, shName)
    m = SplitList(members)
    Select Case UBound(m) + 1
        Case 1: v = HypCell(sh.Name, m(0))
        Case 2: v = HypCell(sh.Name, m(0), m(1))
        Case 3: v = HypCell(sh.Name, m(0), m(1), m(2))
        Case 4: v = HypCell(sh.Name, m(0), m(1), m(2), m(3))
        Case 5: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4))
        Case 6: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5))
        Case 7: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6))
        Case 8: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7))
        Case 9: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8))
        Case 10: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9))
        Case 11: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10))
        Case 12: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11))
        Case 13: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12))
        Case 14: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12), m(13))
        Case 15: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12), m(13), m(14))
        Case 16: v = HypCell(sh.Name, m(0), m(1), m(2), m(3), m(4), m(5), m(6), m(7), m(8), m(9), m(10), m(11), m(12), m(13), m(14), m(15))
        Case Else
            SVB_CellValue = Refused("Pass between 1 and 16 items separated by |.")
            Exit Function
    End Select
    SVB_CellValue = Okay(JVal(v))
    Exit Function
Fail:
    SVB_CellValue = VbaError()
End Function

'------------------------------------------------------------------------------
' Members and outline (provider support varies; see references/vba-api.md)
'------------------------------------------------------------------------------
Public Function SVB_FindMember(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant, _
                               Optional ByVal aliasTable As Variant = "") As String
    On Error GoTo Fail
    Dim sh As Object, dimN As Variant, aliasN As Variant, genN As Variant, lvlN As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypFindMember(sh.Name, memberName, aliasTable, dimN, aliasN, genN, lvlN)
    SVB_FindMember = Result(rc, JObj("dimension", JVal(dimN), "alias", JVal(aliasN), _
                                     "generationName", JVal(genN), "levelName", JVal(lvlN)))
    Exit Function
Fail:
    SVB_FindMember = VbaError()
End Function

Public Function SVB_MemberInfo(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant) As String
    ' All properties Smart View exposes for one member.
    On Error GoTo Fail
    Dim sh As Object, names As Variant, values As Variant, strs As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetMemberInformationEx(sh.Name, memberName, names, values, strs)
    SVB_MemberInfo = Result(rc, JObj("names", JVal(names), "values", JVal(values), "valueStrings", JVal(strs)))
    Exit Function
Fail:
    SVB_MemberInfo = VbaError()
End Function

Public Function SVB_MemberProperty(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant, _
                                   ByVal propertyName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, value As Variant, strs As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetMemberInformation(sh.Name, memberName, propertyName, value, strs)
    SVB_MemberProperty = Result(rc, JObj("value", JVal(value), "valueStrings", JVal(strs)))
    Exit Function
Fail:
    SVB_MemberProperty = VbaError()
End Function

Public Function SVB_Children(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant, _
                             Optional ByVal count As Variant = 0) As String
    On Error GoTo Fail
    Dim sh As Object, kids As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetChildren(sh.Name, memberName, CInt(count), kids)
    SVB_Children = Result(rc, JVal(kids))
    Exit Function
Fail:
    SVB_Children = VbaError()
End Function

Public Function SVB_Parent(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, p As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetParent(sh.Name, memberName, p)
    SVB_Parent = Result(rc, JVal(p))
    Exit Function
Fail:
    SVB_Parent = VbaError()
End Function

Public Function SVB_Ancestor(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant, _
                             ByVal layerType As Variant, ByVal layerNumber As Variant) As String
    On Error GoTo Fail
    Dim sh As Object, a As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypGetAncestor(sh.Name, memberName, layerType, CInt(layerNumber), a)
    SVB_Ancestor = Result(rc, JVal(a))
    Exit Function
Fail:
    SVB_Ancestor = VbaError()
End Function

Public Function SVB_QueryMembers(ByVal wbName As Variant, ByVal shName As Variant, ByVal memberName As Variant, _
                                 ByVal predicate As Variant, Optional ByVal queryOption As Variant = 1, _
                                 Optional ByVal dimensionName As Variant = "", Optional ByVal input1 As Variant = "", _
                                 Optional ByVal input2 As Variant = "") As String
    ' predicate: HYP_CHILDREN=1, HYP_DESCENDANTS=2, HYP_BOTTOMLEVEL=3, HYP_SIBLINGS=4,
    ' HYP_SAMELEVEL=5, HYP_SAMEGENERATION=6, HYP_PARENT=7, HYP_DIMENSION=8,
    ' HYP_NAMEDGENERATION=9, HYP_NAMEDLEVEL=10, HYP_SEARCH=11, HYP_WILDSEARCH=12,
    ' HYP_USERATTRIBUTE=13, HYP_ANCESTORS=14 ...
    ' queryOption: HYP_MEMBERSONLY=1, HYP_ALIASESONLY=2, HYP_MEMBERSANDALIASES=4
    On Error GoTo Fail
    Dim sh As Object, arr As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypQueryMembers(sh.Name, memberName, CLng(predicate), CLng(queryOption), dimensionName, input1, input2, arr)
    SVB_QueryMembers = Result(rc, JVal(arr))
    Exit Function
Fail:
    SVB_QueryMembers = VbaError()
End Function

Public Function SVB_OutlineMemberInfo(ByVal wbName As Variant, ByVal shName As Variant, ByVal dimensionName As Variant, _
                                      ByVal memberName As Variant, ByVal predicate As Variant) As String
    ' predicate: HYP_COMMENT=1, HYP_FORMULA=2, HYP_UDA=3, HYP_ATTRIBUTE=4
    On Error GoTo Fail
    Dim sh As Object, arr As Variant, rc As Long
    Set sh = ResolveSheet(wbName, shName)
    rc = HypOtlGetMemberInfo(sh.Name, dimensionName, memberName, CLng(predicate), arr)
    SVB_OutlineMemberInfo = Result(rc, JVal(arr))
    Exit Function
Fail:
    SVB_OutlineMemberInfo = VbaError()
End Function

Public Function SVB_IsUDA(ByVal wbName As Variant, ByVal shName As Variant, ByVal dimensionName As Variant, _
                          ByVal memberName As Variant, ByVal uda As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_IsUDA = Okay(JVal(HypIsUDA(sh.Name, dimensionName, memberName, uda)))
    Exit Function
Fail:
    SVB_IsUDA = VbaError()
End Function

Public Function SVB_IsAttribute(ByVal wbName As Variant, ByVal shName As Variant, ByVal dimensionName As Variant, _
                                ByVal memberName As Variant, ByVal attributeName As Variant) As String
    On Error GoTo Fail
    Dim sh As Object
    Set sh = ResolveSheet(wbName, shName)
    SVB_IsAttribute = Okay(JVal(HypIsAttribute(sh.Name, dimensionName, memberName, attributeName)))
    Exit Function
Fail:
    SVB_IsAttribute = VbaError()
End Function

'------------------------------------------------------------------------------
' Ribbon menus (guarded)
'------------------------------------------------------------------------------
Public Function SVB_ExecuteMenu(ByVal wbName As Variant, ByVal shName As Variant, ByVal menuName As Variant) As String
    ' Runs a Smart View ribbon command by its label. Menus that open a modal dialog
    ' block until the user closes it. Data-changing commands are refused.
    On Error GoTo Fail
    Dim sh As Object, m As String, banned As Variant, b As Variant
    ' Labels follow the Office language, so the list covers English, Spanish,
    ' French, German, Italian and Portuguese stems (accents left out on purpose).
    m = LCase$(CStr(menuName))
    banned = Array( _
        "submit", "enviar", "envoyer", "soumettre", "senden", "inviare", "invia", _
        "calcul", "berechn", _
        "rule", "regla", "regle", "regel", "regra", "regol", _
        "approv", "aprob", "genehm", _
        "copy version", "copiar vers", "copier la vers", "version kopier", "copia vers", _
        "adjust", "ajust", "anpass", "rettific", _
        "spread", "distribu", "repart", "verteil", "ripart", _
        "consolid", "konsolid", "translat", "tradu", "umrechn", _
        "lock", "bloque", "verrou", "sperr", "blocc", _
        "delete", "elimin", "borrar", "suppr", "losch", "exclu", "cancell", _
        "clear", "limpi", "effac", "leeren", "cancella", _
        "supporting", "detalle", "detail", "dettagl", _
        "cell text", "texto de celda", "texte de cellule", "zelltext", "testo cella", _
        "attach", "adjunt", "piece jointe", "anhang", "anexo", "allegat", _
        "sync", "sincron", "synchron", _
        "copy data", "copiar datos", "paste data", "pegar datos", _
        "job", "trabajo", "travail", _
        "launch", "lanzar", "import", "export")
    For Each b In banned
        If InStr(m, b) > 0 Then
            SVB_ExecuteMenu = Refused("Menu '" & menuName & "' can change data or server artifacts; ask the user to run it.")
            Exit Function
        End If
    Next b
    Set sh = ResolveSheet(wbName, shName)
    SVB_ExecuteMenu = Result(HypExecuteMenu(sh.Name, menuName))
    Exit Function
Fail:
    SVB_ExecuteMenu = VbaError()
End Function

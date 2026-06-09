Option Explicit

'==============================================================================
' MODULE: ContTRF
' PURPOSE: Main controller + helpers to synchronise the "Expiry" sheet in
'          Daily Processes.xlsm with the "Position" sheet in Client Positions.xlsm
'          using FXDC Ref as the unique key.
'==============================================================================

' ── Constants ────────────────────────────────────────────────────────────────
Private Const CP_PATH  As String = "C:\Users\YOUR_USERNAME\Documents\Customer Related\Client Positions.xlsm"
Private Const DP_PATH  As String = "C:\Users\YOUR_USERNAME\Documents\Tools\Daily Processes.xlsm"
Private Const CP_SHEET As String = "Position"
Private Const DP_SHEET As String = "Expiry"
Private Const AUDIT_SHEET As String = "AuditAQDQ"

Private Const CP_HDR_ROW As Long = 6
Private Const DP_HDR_ROW As Long = 1
Private Const CP_DATA_START As Long = 7
Private Const CP_FORMULA_ROW As Long = 7   ' Row whose formulas we replicate
Private Const FXDC_MIN As Long = 10000
Private Const FXDC_MAX As Long = 99999
Private Const NUM_TOLERANCE As Double = 0.000001
Private Const PAUSE_SECS As Double = 1

' ── Entry Point ──────────────────────────────────────────────────────────────
Public Sub RunContTRF()
    Dim wbCP As Workbook, wbDP As Workbook
    Dim wsCP As Worksheet, wsDP As Worksheet
    Dim wsAudit As Worksheet

    Dim dictDP As Object   ' Scripting.Dictionary  key=FXDC Ref(Long) -> row array
    Dim dictCP As Object   ' Scripting.Dictionary  key=FXDC Ref(Long) -> row number

    Dim colMapCP As Object ' header->col index for CP
    Dim colMapDP As Object ' header->col index for DP

    Dim newEntries As Collection
    Dim dowMismatches As Collection

    On Error GoTo ErrHandler

    ' 1. Open / attach workbooks
    Set wbCP = GetOrOpenWorkbook(CP_PATH)
    Set wbDP = GetOrOpenWorkbook(DP_PATH)

    Set wsCP = GetWorksheet(wbCP, CP_SHEET)
    Set wsDP = GetWorksheet(wbDP, DP_SHEET)

    ' 2. Map headers
    Set colMapCP = MapHeaders(wsCP, CP_HDR_ROW)
    Set colMapDP = MapHeaders(wsDP, DP_HDR_ROW)

    ValidateRequiredHeaders colMapCP, Array("FXDC Ref", "Customer", "CCYpair", "Struct", "Ccy", "Live", _
                                            "Notional", "Lower eKI", "Lower Strike", "Upper Strike", _
                                            "Upper eKI", "KO/Pivot", "Expiry")
    ValidateRequiredHeaders colMapDP, Array("FXDC Ref", "Customer", "CCYpair", "Struct", "Ccy", "Live", _
                                            "Notional", "Lower eKI", "Lower Strike", "Upper Strike", _
                                            "Upper eKI", "KO/Pivot", "Expiry")

    ' 3. Load dictionaries
    Set dictDP = LoadDPDictionary(wsDP, colMapDP)
    Dim dpDOW As Integer
    dpDOW = GetDPExpiryDOWNum(wsDP, colMapDP)   ' 1=Mon..5=Fri
    Set dictCP = LoadCPDictionary(wsCP, colMapCP, dpDOW)

    ' 4. Capture formula templates from CP row 7 (before any writes)
    Dim formulaTemplates() As String
    formulaTemplates = CaptureFormulas(wsCP, CP_FORMULA_ROW, colMapCP)

    ' 5. Synchronise
    Set newEntries = New Collection
    Set dowMismatches = New Collection

    SynchroniseCPwithDP wsCP, wsDP, colMapCP, colMapDP, dictCP, dictDP, _
                         formulaTemplates, newEntries, dowMismatches

    ' 6. Call existing macros in Client Positions
    Dim expiryDOW As String
    expiryDOW = GetDPExpiryDOW(wsDP, colMapDP)
    CallExistingMacros wbCP, expiryDOW

    ' 7. Generate audit sheet
    Set wsAudit = GetOrCreateSheet(wbCP, AUDIT_SHEET)
    GenerateAuditSheet wsAudit, dictDP, dictCP, newEntries, dowMismatches

    ' 8. User feedback
    wsAudit.Activate
    MsgBox "Process Complete Successfully.", vbInformation, "ContTRF"

    Exit Sub

ErrHandler:
    MsgBox "ContTRF Error:" & vbCrLf & Err.Number & " - " & Err.Description, _
           vbCritical, "ContTRF"
End Sub

'==============================================================================
' SECTION 1: Workbook / Sheet helpers
'==============================================================================

Private Function GetOrOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    Dim wbName As String
    wbName = Mid(fullPath, InStrRev(fullPath, "\") + 1)

    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0

    If wb Is Nothing Then
        If Dir(fullPath) = "" Then
            Err.Raise vbObjectError + 1, "GetOrOpenWorkbook", _
                      "File not found: " & fullPath
        End If
        Set wb = Workbooks.Open(fullPath)
    End If

    Set GetOrOpenWorkbook = wb
End Function

Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2, "GetWorksheet", _
                  "Sheet '" & sheetName & "' not found in " & wb.Name
    End If
    Set GetWorksheet = ws
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

'==============================================================================
' SECTION 2: Header mapping
'==============================================================================

Private Function MapHeaders(ByVal ws As Worksheet, ByVal hdrRow As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastCol As Long
    lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        Dim hdr As String
        hdr = Trim(CStr(ws.Cells(hdrRow, c).Value))
        If hdr <> "" Then
            If Not dict.Exists(hdr) Then
                dict.Add hdr, c
            End If
        End If
    Next c

    Set MapHeaders = dict
End Function

Private Sub ValidateRequiredHeaders(ByVal colMap As Object, ByVal required As Variant)
    Dim h As Variant
    For Each h In required
        If Not colMap.Exists(h) Then
            Err.Raise vbObjectError + 3, "ValidateRequiredHeaders", _
                      "Required header '" & h & "' not found."
        End If
    Next h
End Sub

'==============================================================================
' SECTION 3: Load dictionaries
'==============================================================================

Private Function LoadDPDictionary(ByVal wsDP As Worksheet, _
                                   ByVal colMap As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fxdcCol As Long, liveCol As Long
    fxdcCol = colMap("FXDC Ref")
    liveCol = colMap("Live")

    Dim lastRow As Long
    lastRow = wsDP.Cells(wsDP.Rows.Count, fxdcCol).End(xlUp).Row

    Dim r As Long
    For r = DP_HDR_ROW + 1 To lastRow
        Dim rawFXDC As Variant
        rawFXDC = wsDP.Cells(r, fxdcCol).Value

        ' Validate 5-digit
        If Not IsNumeric(rawFXDC) Then GoTo NextDPRow
        Dim fxdcLong As Long
        fxdcLong = CLng(rawFXDC)
        If fxdcLong < FXDC_MIN Or fxdcLong > FXDC_MAX Then GoTo NextDPRow

        ' Validate Live = "Live"
        Dim liveVal As String
        liveVal = Trim(CStr(wsDP.Cells(r, liveCol).Value))
        If liveVal <> "Live" Then GoTo NextDPRow

        ' Duplicate check
        If dict.Exists(fxdcLong) Then
            Err.Raise vbObjectError + 4, "LoadDPDictionary", _
                      "Duplicate FXDC Ref " & fxdcLong & " found in DP sheet. Process aborted."
        End If

        dict.Add fxdcLong, r
NextDPRow:
    Next r

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 5, "LoadDPDictionary", _
                  "No valid Live rows found in DP sheet. Process aborted."
    End If

    Set LoadDPDictionary = dict
End Function

Private Function LoadCPDictionary(ByVal wsCP As Worksheet, _
                                   ByVal colMap As Object, _
                                   ByVal dpDOW As Integer) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fxdcCol As Long, expiryCol As Long
    fxdcCol  = colMap("FXDC Ref")
    expiryCol = colMap("Expiry")

    Dim lastRow As Long
    lastRow = wsCP.Cells(wsCP.Rows.Count, fxdcCol).End(xlUp).Row

    Dim r As Long
    For r = CP_DATA_START To lastRow
        Dim rawFXDC As Variant
        rawFXDC = wsCP.Cells(r, fxdcCol).Value
        If Not IsNumeric(rawFXDC) Then GoTo NextCPRow
        Dim fxdcLong As Long
        fxdcLong = CLng(rawFXDC)

        ' Skip blank or MANUAL expiry
        Dim expiryVal As Variant
        expiryVal = wsCP.Cells(r, expiryCol).Value
        If IsEmpty(expiryVal) Then GoTo NextCPRow
        If Trim(CStr(expiryVal)) = "" Then GoTo NextCPRow
        If UCase(Trim(CStr(expiryVal))) = "MANUAL" Then GoTo NextCPRow

        ' Skip if Expiry DOW does not match DP DOW
        If Not IsDate(expiryVal) Then GoTo NextCPRow
        If Weekday(CDate(expiryVal), vbMonday) <> dpDOW Then GoTo NextCPRow

        If dict.Exists(fxdcLong) Then
            Err.Raise vbObjectError + 6, "LoadCPDictionary", _
                      "Duplicate FXDC Ref " & fxdcLong & " found in CP sheet. Process aborted."
        End If

        dict.Add fxdcLong, r
NextCPRow:
    Next r

    Set LoadCPDictionary = dict
End Function

'==============================================================================
' SECTION 4: Capture formula templates from CP row 7
'==============================================================================

Private Function CaptureFormulas(ByVal wsCP As Worksheet, _
                                  ByVal templateRow As Long, _
                                  ByVal colMap As Object) As String()
    Dim lastCol As Long
    lastCol = wsCP.Cells(CP_HDR_ROW, wsCP.Columns.Count).End(xlToLeft).Column

    Dim formulas() As String
    ReDim formulas(1 To lastCol)

    Dim c As Long
    For c = 1 To lastCol
        ' Skip column 15 as per spec
        If c = 15 Then
            formulas(c) = ""
        Else
            Dim cell As Range
            Set cell = wsCP.Cells(templateRow, c)
            If cell.HasFormula Then
                formulas(c) = cell.Formula
            Else
                formulas(c) = ""
            End If
        End If
    Next c

    CaptureFormulas = formulas
End Function

' Adjust a formula from templateRow to targetRow, keeping absolute $row refs intact
Private Function AdjustFormula(ByVal formula As String, _
                                ByVal fromRow As Long, _
                                ByVal toRow As Long) As String
    If formula = "" Then
        AdjustFormula = ""
        Exit Function
    End If
    ' Use Excel's built-in offset: write to template cell temporarily? 
    ' Instead, do a simple relative-row token replacement.
    ' Replace non-$ row references only.
    Dim adjusted As String
    adjusted = formula

    ' Replace relative row references (not preceded by $)
    ' Strategy: shift by (toRow - fromRow)
    Dim offset As Long
    offset = toRow - fromRow

    ' Use regex-like replacement via a helper
    adjusted = ShiftRelativeRows(adjusted, fromRow, offset)
    AdjustFormula = adjusted
End Function

Private Function ShiftRelativeRows(ByVal formula As String, _
                                    ByVal fromRow As Long, _
                                    ByVal offset As Long) As String
    ' Walk char by char; when we see a letter (col ref) not preceded by $
    ' and followed by digits, shift those digits if not preceded by $
    Dim result As String
    result = ""
    Dim i As Long
    Dim n As Long
    n = Len(formula)

    i = 1
    Do While i <= n
        Dim ch As String
        ch = Mid(formula, i, 1)

        ' Detect a row number token: digit sequence possibly after column letters
        ' We look for: [A-Za-z] not preceded by $ + digits
        If ch Like "[A-Za-z]" Then
            ' Collect full column letters
            Dim colPart As String
            colPart = ""
            Do While i <= n And Mid(formula, i, 1) Like "[A-Za-z]"
                colPart = colPart & Mid(formula, i, 1)
                i = i + 1
            Loop
            ' Now collect digits (row number)
            If i <= n And Mid(formula, i, 1) Like "[0-9]" Then
                Dim rowPart As String
                rowPart = ""
                Do While i <= n And Mid(formula, i, 1) Like "[0-9]"
                    rowPart = rowPart & Mid(formula, i, 1)
                    i = i + 1
                Loop
                ' Check whether the preceding char was $
                Dim prevChar As String
                If Len(result) > 0 Then
                    prevChar = Right(result, 1)
                Else
                    prevChar = ""
                End If
                If prevChar = "$" Then
                    ' Absolute row - do not shift
                    result = result & colPart & rowPart
                Else
                    ' Relative row - shift
                    Dim origRow As Long
                    origRow = CLng(rowPart)
                    result = result & colPart & CStr(origRow + offset)
                End If
            Else
                result = result & colPart
            End If
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    ShiftRelativeRows = result
End Function

'==============================================================================
' SECTION 5: Synchronise CP with DP
'==============================================================================

Private Sub SynchroniseCPwithDP(ByVal wsCP As Worksheet, _
                                  ByVal wsDP As Worksheet, _
                                  ByVal colMapCP As Object, _
                                  ByVal colMapDP As Object, _
                                  ByVal dictCP As Object, _
                                  ByVal dictDP As Object, _
                                  ByRef  formulaTemplates() As String, _
                                  ByVal newEntries As Collection, _
                                  ByVal dowMismatches As Collection)

    Dim lastCPCol As Long
    lastCPCol = wsCP.Cells(CP_HDR_ROW, wsCP.Columns.Count).End(xlToLeft).Column

    ' Pre-size a 2D array for new rows: rows up to dictDP.Count, cols = lastCPCol
    ' We'll track actual count separately and slice before writing.
    Dim maxNewRows As Long
    maxNewRows = dictDP.Count   ' upper bound - actual new rows <= this
    Dim newRowsData() As Variant
    ReDim newRowsData(1 To maxNewRows, 1 To lastCPCol)
    Dim newRowCount As Long
    newRowCount = 0

    ' Determine next empty row in CP
    Dim fxdcColCP As Long
    fxdcColCP = colMapCP("FXDC Ref")
    Dim nextCPRow As Long
    nextCPRow = wsCP.Cells(wsCP.Rows.Count, fxdcColCP).End(xlUp).Row + 1

    Dim dpKey As Variant
    For Each dpKey In dictDP.Keys
        Dim dpRow As Long
        dpRow = dictDP(dpKey)

        If dictCP.Exists(dpKey) Then
            ' ── Existing CP row: update fields ──────────────────────────────
            Dim cpRow As Long
            cpRow = dictCP(dpKey)
            UpdateCPRow wsCP, wsDP, cpRow, dpRow, colMapCP, colMapDP, dpKey, dowMismatches
        Else
            ' ── New trade: stage into 2D array ──────────────────────────────
            newEntries.Add dpKey
            newRowCount = newRowCount + 1

            ' Copy all matching fields from DP into the staged row
            Dim dpHdr As Variant
            For Each dpHdr In colMapDP.Keys
                If colMapCP.Exists(dpHdr) Then
                    Dim dpVal As Variant
                    dpVal = wsDP.Cells(dpRow, colMapDP(dpHdr)).Value

                    ' Notional: store as text with leading apostrophe
                    If CStr(dpHdr) = "Notional" Then
                        dpVal = "'" & CStr(dpVal)
                    End If

                    newRowsData(newRowCount, colMapCP(dpHdr)) = dpVal
                End If
            Next dpHdr

            ' Track CP dictionary so audit sees the new row
            dictCP.Add dpKey, nextCPRow
            nextCPRow = nextCPRow + 1
        End If
    Next dpKey

    ' ── Batch-write new rows ────────────────────────────────────────────────
    If newRowCount > 0 Then
        WriteNewRows wsCP, colMapCP, formulaTemplates, newRowsData, newRowCount, lastCPCol
    End If
End Sub

Private Sub UpdateCPRow(ByVal wsCP As Worksheet, _
                         ByVal wsDP As Worksheet, _
                         ByVal cpRow As Long, _
                         ByVal dpRow As Long, _
                         ByVal colMapCP As Object, _
                         ByVal colMapDP As Object, _
                         ByVal fxdcRef As Variant, _
                         ByVal dowMismatches As Collection)

    ' Text fields: overwrite if different
    Dim textFields As Variant
    textFields = Array("Customer", "CCYpair", "Struct", "Ccy", "Live")

    Dim f As Variant
    For Each f In textFields
        If colMapCP.Exists(f) And colMapDP.Exists(f) Then
            Dim dpTxt As String, cpTxt As String
            dpTxt = CStr(wsDP.Cells(dpRow, colMapDP(f)).Value)
            cpTxt = CStr(wsCP.Cells(cpRow, colMapCP(f)).Value)
            If dpTxt <> cpTxt Then
                wsCP.Cells(cpRow, colMapCP(f)).Value = dpTxt
            End If
        End If
    Next f

    ' Notional: compare as text, write prefixed with `
    If colMapCP.Exists("Notional") And colMapDP.Exists("Notional") Then
        Dim dpNot As String, cpNot As String
        dpNot = CStr(wsDP.Cells(dpRow, colMapDP("Notional")).Value)
        cpNot = CStr(wsCP.Cells(cpRow, colMapCP("Notional")).Value)
        If dpNot <> cpNot Then
            wsCP.Cells(cpRow, colMapCP("Notional")).Value = "'" & dpNot
        End If
    End If

    ' Numeric fields: overwrite if difference > tolerance
    Dim numFields As Variant
    numFields = Array("Lower eKI", "Lower Strike", "Upper Strike", "Upper eKI")

    For Each f In numFields
        If colMapCP.Exists(f) And colMapDP.Exists(f) Then
            Dim dpNum As Double, cpNum As Double
            dpNum = ToDouble(wsDP.Cells(dpRow, colMapDP(f)).Value)
            cpNum = ToDouble(wsCP.Cells(cpRow, colMapCP(f)).Value)
            If Abs(dpNum - cpNum) > NUM_TOLERANCE Then
                wsCP.Cells(cpRow, colMapCP(f)).Value = dpNum
            End If
        End If
    Next f

    ' KO/Pivot: overwrite only if DP is non-zero and non-blank and differs
    If colMapCP.Exists("KO/Pivot") And colMapDP.Exists("KO/Pivot") Then
        Dim dpKO As Variant, cpKO As Variant
        dpKO = wsDP.Cells(dpRow, colMapDP("KO/Pivot")).Value
        cpKO = wsCP.Cells(cpRow, colMapCP("KO/Pivot")).Value
        If IsNumeric(dpKO) Then
            If CDbl(dpKO) <> 0 And Abs(ToDouble(dpKO) - ToDouble(cpKO)) > NUM_TOLERANCE Then
                wsCP.Cells(cpRow, colMapCP("KO/Pivot")).Value = dpKO
            End If
        ElseIf CStr(dpKO) <> "" Then
            If CStr(dpKO) <> CStr(cpKO) Then
                wsCP.Cells(cpRow, colMapCP("KO/Pivot")).Value = dpKO
            End If
        End If
    End If

    ' Expiry: check DOW mismatch
    If colMapCP.Exists("Expiry") And colMapDP.Exists("Expiry") Then
        Dim dpExpiry As Variant, cpExpiry As Variant
        dpExpiry = wsDP.Cells(dpRow, colMapDP("Expiry")).Value
        cpExpiry = wsCP.Cells(cpRow, colMapCP("Expiry")).Value

        If IsDate(dpExpiry) And IsDate(cpExpiry) Then
            Dim dpDOW As Integer, cpDOW As Integer
            dpDOW = Weekday(CDate(dpExpiry), vbMonday)
            cpDOW = Weekday(CDate(cpExpiry), vbMonday)
            If dpDOW <> cpDOW Then
                ' Color CP Expiry cell coral
                wsCP.Cells(cpRow, colMapCP("Expiry")).Interior.Color = RGB(255, 127, 80)
                dowMismatches.Add fxdcRef
            End If
        End If
    End If
End Sub

Private Sub WriteNewRows(ByVal wsCP As Worksheet, _
                          ByVal colMapCP As Object, _
                          ByRef  formulaTemplates() As String, _
                          ByRef  newRowsData() As Variant, _
                          ByVal newRowCount As Long, _
                          ByVal lastCPCol As Long)

    Dim fxdcColCP As Long
    fxdcColCP = colMapCP("FXDC Ref")
    Dim startRow As Long
    startRow = wsCP.Cells(wsCP.Rows.Count, fxdcColCP).End(xlUp).Row + 1

    Dim i As Long
    For i = 1 To newRowCount
        Dim targetRow As Long
        targetRow = startRow + (i - 1)

        ' Write value columns first
        Dim c As Long
        For c = 1 To lastCPCol
            If c <> 15 Then   ' skip column 15 per spec
                Dim cellVal As Variant
                cellVal = newRowsData(i, c)
                If formulaTemplates(c) <> "" Then
                    ' Formula column - handled below
                ElseIf Not IsEmpty(cellVal) Then
                    wsCP.Cells(targetRow, c).Value = cellVal
                End If
            End If
        Next c

        ' Apply adjusted formulas
        For c = 1 To lastCPCol
            If formulaTemplates(c) <> "" And c <> 15 Then
                Dim adj As String
                adj = AdjustFormula(formulaTemplates(c), CP_FORMULA_ROW, targetRow)
                If adj <> "" Then
                    wsCP.Cells(targetRow, c).Formula = adj
                End If
            End If
        Next c
    Next i
End Sub

'==============================================================================
' SECTION 6: DOW helper
'==============================================================================

Private Function GetDPExpiryDOW(ByVal wsDP As Worksheet, _
                                  ByVal colMapDP As Object) As String
    Dim expiryCol As Long
    expiryCol = colMapDP("Expiry")

    Dim firstDataRow As Long
    firstDataRow = DP_HDR_ROW + 1

    Dim expiryVal As Variant
    expiryVal = wsDP.Cells(firstDataRow, expiryCol).Value

    If IsDate(expiryVal) Then
        Select Case Weekday(CDate(expiryVal), vbMonday)
            Case 1: GetDPExpiryDOW = "Mon"
            Case 2: GetDPExpiryDOW = "Tue"
            Case 3: GetDPExpiryDOW = "Wed"
            Case 4: GetDPExpiryDOW = "Thu"
            Case 5: GetDPExpiryDOW = "Fri"
            Case Else: GetDPExpiryDOW = "Mon"
        End Select
    Else
        GetDPExpiryDOW = "Mon"
    End If
End Function

' Returns DP expiry DOW as Integer (1=Mon .. 5=Fri, vbMonday base)
Private Function GetDPExpiryDOWNum(ByVal wsDP As Worksheet, _
                                    ByVal colMapDP As Object) As Integer
    Dim expiryCol As Long
    expiryCol = colMapDP("Expiry")

    Dim r As Long
    Dim lastRow As Long
    lastRow = wsDP.Cells(wsDP.Rows.Count, expiryCol).End(xlUp).Row

    ' Find the first valid date in the Expiry column
    For r = DP_HDR_ROW + 1 To lastRow
        Dim v As Variant
        v = wsDP.Cells(r, expiryCol).Value
        If IsDate(v) Then
            GetDPExpiryDOWNum = Weekday(CDate(v), vbMonday)
            Exit Function
        End If
    Next r

    GetDPExpiryDOWNum = 1   ' default Monday if nothing found
End Function

'==============================================================================
' SECTION 7: Call existing macros in Client Positions
'==============================================================================

Private Sub CallExistingMacros(ByVal wbCP As Workbook, ByVal dow As String)
    Application.Wait Now + TimeValue("00:00:01")

    ' 1. FixDayOfWeek variant based on DOW
    Dim fixDOWMacro As String
    Select Case dow
        Case "Mon": fixDOWMacro = "FixMon"
        Case "Tue": fixDOWMacro = "FixTue"
        Case "Wed": fixDOWMacro = "FixWed"
        Case "Thu": fixDOWMacro = "FixThu"
        Case "Fri": fixDOWMacro = "FixFri"
        Case Else:  fixDOWMacro = "FixMon"
    End Select

    ' --- STUB: replace the Run calls below with actual macro invocations ---
    ' e.g. Application.Run "'" & wbCP.Name & "'!ModuleName." & fixDOWMacro
    RunMacroStub wbCP, fixDOWMacro
    Application.Wait Now + TimeValue("00:00:01")

    RunMacroStub wbCP, "AQQQs"
    Application.Wait Now + TimeValue("00:00:01")

    RunMacroStub wbCP, "SortCustomer"
    Application.Wait Now + TimeValue("00:00:01")

    RunMacroStub wbCP, "FixFormatting"
    Application.Wait Now + TimeValue("00:00:01")
End Sub

' STUB – replace body with the real call once you know the module name
Private Sub RunMacroStub(ByVal wb As Workbook, ByVal macroName As String)
    On Error Resume Next
    Application.Run "'" & wb.Name & "'!" & macroName
    If Err.Number <> 0 Then
        Debug.Print "Stub/macro not found: " & macroName & " (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' SECTION 8: Audit sheet (AuditAQDQ)
'==============================================================================

Private Sub GenerateAuditSheet(ByVal wsAudit As Worksheet, _
                                 ByVal dictDP As Object, _
                                 ByVal dictCP As Object, _
                                 ByVal newEntries As Collection, _
                                 ByVal dowMismatches As Collection)

    ' Clear existing data from row 11 down (preserve rows 1-10 headers / tables)
    Dim lastAuditRow As Long
    With wsAudit
        lastAuditRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lastAuditRow >= 11 Then
            .Rows("11:" & lastAuditRow).ClearContents
            .Rows("11:" & lastAuditRow).Interior.ColorIndex = xlNone
        End If

        ' Add tables row 10 (simple labels)
        .Cells(10, 1).Value = "Col A: All DP FXDC Refs"
        .Cells(10, 2).Value = "Col B: CP Refs matching DP DOW"
        .Cells(10, 5).Value = "Col E: Newly Added to CP"
        .Cells(10, 6).Value = "Col F: DOW Mismatches"
    End With

    ' ── Column A: all DP FXDC Refs (sorted ascending) ─────────────────────
    Dim dpKeys() As Variant
    dpKeys = SortedKeys(dictDP)

    Dim r As Long
    r = 11
    Dim k As Long
    For k = 0 To UBound(dpKeys)
        wsAudit.Cells(r, 1).Value = dpKeys(k)
        r = r + 1
    Next k

    ' ── Column B: CP FXDC Refs whose DOW matches DP DOW ───────────────────
    ' (We've already flagged mismatches in dowMismatches; B = dictCP keys NOT in dowMismatches)
    Dim mismatchDict As Object
    Set mismatchDict = CreateObject("Scripting.Dictionary")
    Dim mm As Variant
    For Each mm In dowMismatches
        mismatchDict(mm) = True
    Next mm

    Dim cpKeys() As Variant
    cpKeys = SortedKeys(dictCP)

    r = 11
    For k = 0 To UBound(cpKeys)
        If Not mismatchDict.Exists(cpKeys(k)) Then
            wsAudit.Cells(r, 2).Value = cpKeys(k)
            r = r + 1
        End If
    Next k

    ' ── Column E: newly added FXDC Refs ───────────────────────────────────
    Dim newArr() As Long
    ReDim newArr(0 To newEntries.Count - 1)
    Dim idx As Long
    idx = 0
    Dim ne As Variant
    For Each ne In newEntries
        newArr(idx) = ne
        idx = idx + 1
    Next ne
    BubbleSortLong newArr

    r = 11
    For k = 0 To UBound(newArr)
        wsAudit.Cells(r, 5).Value = newArr(k)
        r = r + 1
    Next k

    ' ── Column F: DOW mismatch FXDC Refs ──────────────────────────────────
    Dim mmArr() As Long
    ReDim mmArr(0 To dowMismatches.Count - 1)
    idx = 0
    For Each mm In dowMismatches
        mmArr(idx) = mm
        idx = idx + 1
    Next mm
    BubbleSortLong mmArr

    r = 11
    For k = 0 To UBound(mmArr)
        wsAudit.Cells(r, 6).Value = mmArr(k)
        r = r + 1
    Next k

    ' ── Conditional formatting A11:B<last> ────────────────────────────────
    ApplyUniqueHighlight wsAudit, dpKeys, cpKeys

    wsAudit.Columns("A:F").AutoFit
End Sub

' Highlight values in col A or B that appear in only ONE of the two columns (coral fill)
Private Sub ApplyUniqueHighlight(ByVal wsAudit As Worksheet, _
                                   ByVal dpKeys() As Variant, _
                                   ByVal cpKeys() As Variant)
    Dim dictA As Object, dictB As Object
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In dpKeys
        dictA(k) = True
    Next k
    For Each k In cpKeys
        dictB(k) = True
    Next k

    Dim r As Long
    r = 11
    Do While wsAudit.Cells(r, 1).Value <> "" Or wsAudit.Cells(r, 2).Value <> ""
        ' Col A
        Dim valA As Variant
        valA = wsAudit.Cells(r, 1).Value
        If valA <> "" Then
            If Not dictB.Exists(valA) Then
                wsAudit.Cells(r, 1).Interior.Color = RGB(255, 127, 80) ' coral
            End If
        End If
        ' Col B
        Dim valB As Variant
        valB = wsAudit.Cells(r, 2).Value
        If valB <> "" Then
            If Not dictA.Exists(valB) Then
                wsAudit.Cells(r, 2).Interior.Color = RGB(255, 127, 80)
            End If
        End If
        r = r + 1
        If r > wsAudit.Rows.Count Then Exit Do
    Loop
End Sub

'==============================================================================
' SECTION 9: Utilities
'==============================================================================

Private Function ToDouble(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToDouble = CDbl(v)
    Else
        ToDouble = 0
    End If
End Function

Private Function SortedKeys(ByVal dict As Object) As Variant()
    Dim arr() As Variant
    arr = dict.Keys

    Dim i As Long, j As Long
    Dim tmp As Variant
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    SortedKeys = arr
End Function

Private Sub BubbleSortLong(ByRef arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

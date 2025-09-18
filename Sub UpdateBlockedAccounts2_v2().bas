Sub UpdateBlockedAccounts3()
    Application.ScreenUpdating = False
    Dim wb As Workbook, ws As Worksheet, rng As Range, cell As Range
    Dim arrData() As Variant, i As Long, lastRow As Long
    Dim SrcFullPath As String, DstWkbk As Workbook, strPassword As String
    Dim UV As UniqueValues, MyRange As Range
    Dim vCrit As Variant
    Dim wsCriteria As Worksheet
    Dim rngCriteria As Range
    Dim lastRowCriteria As Long

    Set DstWkbk = ThisWorkbook
    SrcFullPath = DstWkbk.Worksheets("Control").Range("B9").Value

    If Dir(SrcFullPath) = "" Then
        MsgBox SrcFullPath & " Not found", vbExclamation
        GoTo EndProc
    End If

    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    On Error Resume Next

    If Left(Right(SrcFullPath, Len(SrcFullPath) - InStrRev(SrcFullPath, "\")), 7) = "Report_" Then
        strPassword = "abc"
    Else
        strPassword = ""
    End If

    Set wb = Workbooks.Open(SrcFullPath, False, True, , strPassword)
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    On Error GoTo 0

    If wb Is Nothing Then
        MsgBox SrcFullPath & " could not be opened.", vbExclamation
        GoTo EndProc
    End If

    If Left(wb.Name, 7) = "Report_" Then
        Set ws = wb.Sheets("ATTENTION REQUIRED")
        Set rng = ws.Range("A1").CurrentRegion
        
        ' Create a temporary criteria range for Advanced Filter
        Set wsCriteria = wb.Worksheets.Add(After:=ws)
        wsCriteria.Name = "Criteria_Temp"

        ' Set the headers for the criteria range
        wsCriteria.Range("A1").Value = ws.Range("B1").Value ' Column B header
        wsCriteria.Range("B1").Value = ws.Range("M1").Value ' Column M header
        wsCriteria.Range("C1").Value = ws.Range("P1").Value ' Column P header
        wsCriteria.Range("D1").Value = ws.Range("V1").Value ' Column V header
        wsCriteria.Range("E1").Value = ws.Range("Y1").Value ' Column Y header
        wsCriteria.Range("F1").Value = ws.Range("AB1").Value ' Column AB header
        wsCriteria.Range("G1").Value = ws.Range("AE1").Value ' Column AE header

        ' Fill in the criteria for "OR" logic
        wsCriteria.Range("A2").Value = "Y"
        wsCriteria.Range("B2").Value = "FX"
        wsCriteria.Range("A3").Value = "Y"
        wsCriteria.Range("C3").Value = "FX"
        wsCriteria.Range("A4").Value = "Y"
        wsCriteria.Range("D4").Value = "FX"
        wsCriteria.Range("A5").Value = "Y"
        wsCriteria.Range("E5").Value = "FX"
        wsCriteria.Range("A6").Value = "Y"
        wsCriteria.Range("F6").Value = "FX"
        wsCriteria.Range("A7").Value = "Y"
        wsCriteria.Range("G7").Value = "FX"
        
        lastRowCriteria = wsCriteria.Cells(wsCriteria.Rows.Count, "A").End(xlUp).Row
        Set rngCriteria = wsCriteria.Range("A1:G" & lastRowCriteria)
        
        ' Apply the Advanced Filter
        rng.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rngCriteria, Unique:=False
        
        ' Data extraction logic
        If rng.Columns(1).SpecialCells(xlCellTypeVisible).Count > 0 Then
            ReDim arrData(1 To rng.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1)
            i = 1
            For Each cell In rng.Columns(1).SpecialCells(xlCellTypeVisible)
                If i <> 1 Then
                    arrData(i - 1) = cell.Value
                End If
                i = i + 1
            Next cell
            Call SortArray(arrData)
        Else
            ReDim arrData(0)
        End If
        
        ws.AutoFilterMode = False ' Clear the filter
        Application.DisplayAlerts = False
        wsCriteria.Delete ' Delete the temporary criteria sheet
        Application.DisplayAlerts = True

    ElseIf Right(wb.Name, 35) = "_Client_Service_Registry_Report.csv" Then
        ' ... (No changes here)
        For Each ws In wb.Sheets
            If Right(ws.Name, 22) = "Client_Service_Registr" Then
                Set rng = ws.Range("A1").CurrentRegion
                rng.AutoFilter Field:=3, Criteria1:="Y"
                rng.AutoFilter Field:=13, Criteria1:="Yes"
                rng.AutoFilter Field:=14, Criteria1:="FX"
                
                If rng.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
                    ReDim arrData(1 To rng.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1)
                    i = 1
                    For Each cell In rng.Columns(1).SpecialCells(xlCellTypeVisible)
                        If cell.Row > 1 Then
                            arrData(i) = cell.Value
                            i = i + 1
                        End If
                    Next cell
                    Call SortArray(arrData)
                Else
                    ReDim arrData(0)
                End If
                Exit For
            End If
        Next ws
    Else
        MsgBox "File does not look right.", vbExclamation
        GoTo EndProc:
    End If

    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True

    DstWkbk.Sheets("BlockAC").Activate
    Columns("A:A").ClearContents
    ActiveSheet.Range("A2:A" & UBound(arrData) + 1).Value = Application.Transpose(arrData)
    Range("A1").Select
    Selection.Value = Format(Date, "dd-mmm-yy")
  
    lastRow = max(ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row, ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row)
    Set MyRange = Range("$A$2:$B$" & lastRow)
    ActiveSheet.Cells.FormatConditions.Delete
    Set UV = MyRange.FormatConditions.AddUniqueValues
    
    With UV
        .DupeUnique = xlUnique
        .Interior.Color = RGB(247, 101, 115)
        .Font.Color = RGB(0, 0, 0)
    End With
    
EndProc:
    Application.ScreenUpdating = True
End Sub
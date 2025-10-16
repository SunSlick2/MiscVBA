Sub EmailExcelRange()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim rngToEmail As Range
    Dim EmailBody As String
    Dim SignatureName As String
    
    ' Set your hard-coded values here
    Const EmailTo As String = "recipient@example.com"
    Const EmailCC As String = "cc@example.com"
    Const EmailTitle As String = "Your Email Subject"
    Const EmailSalutation As String = "Dear Team,"
    SignatureName = "Your Signature Name" ' Change to your actual signature name
    
    ' Set the range to email (A2:D28)
    Set rngToEmail = ThisWorkbook.Worksheets("Sheet1").Range("A2:D28")
    
    ' Create Outlook application
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    If OutlookApp Is Nothing Then
        MsgBox "Cannot create Outlook object. Please make sure Outlook is installed."
        Exit Sub
    End If
    
    ' Create new email
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Build email body
    EmailBody = EmailSalutation & "<br><br>" & _
                "Please find the data below:<br><br>"
    
    With OutlookMail
        .To = EmailTo
        .CC = EmailCC
        .Subject = EmailTitle
        .HTMLBody = EmailBody & RangeToHTML(rngToEmail)
        
        ' Add signature
        On Error Resume Next
        .Display ' Display email to apply signature
        .HTMLBody = .HTMLBody & GetSignature(SignatureName)
        On Error GoTo 0
        
        '.Send ' Uncomment this line to send automatically
    End With
    
    ' Clean up
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Set rngToEmail = Nothing
    
    MsgBox "Email created successfully!", vbInformation
End Sub

Function RangeToHTML(rng As Range) As String
    ' Convert Excel range to HTML table
    Dim TempWorksheet As Worksheet
    Dim TempWorkbook As Workbook
    Dim FilePath As String
    
    Application.ScreenUpdating = False
    
    ' Create temporary workbook
    Set TempWorkbook = Workbooks.Add(1)
    Set TempWorksheet = TempWorkbook.Sheets(1)
    
    ' Copy range to temporary worksheet
    rng.Copy
    With TempWorksheet.Cells(1, 1)
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
    End With
    
    ' Save as HTML
    FilePath = Environ$("temp") & "\temp_range.html"
    TempWorkbook.SaveAs FilePath, FileFormat:=xlHtml
    TempWorkbook.Close SaveChanges:=False
    
    ' Read HTML content
    Dim FSO As Object
    Dim TS As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TS = FSO.OpenTextFile(FilePath)
    RangeToHTML = TS.ReadAll
    TS.Close
    
    ' Clean up temporary file
    Kill FilePath
    
    Application.ScreenUpdating = True
    
    Set TS = Nothing
    Set FSO = Nothing
    Set TempWorksheet = Nothing
    Set TempWorkbook = Nothing
End Function

Function GetSignature(SignatureName As String) As String
    Dim OutlookApp As Object
    Dim Signature As String
    Dim SignaturePath As String
    
    On Error GoTo ErrorHandler
    
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Try to get signature by creating a temporary email
    Dim TempMail As Object
    Set TempMail = OutlookApp.CreateItem(0)
    TempMail.Display
    
    ' Outlook signatures are automatically added when displaying
    Signature = TempMail.HTMLBody
    
    ' Clean up
    TempMail.Close 0 ' olDiscard
    Set TempMail = Nothing
    Set OutlookApp = Nothing
    
    GetSignature = Signature
    Exit Function
    
ErrorHandler:
    GetSignature = "<br><br>Best regards,<br>Your Name" ' Fallback signature
    Set TempMail = Nothing
    Set OutlookApp = Nothing
End Function
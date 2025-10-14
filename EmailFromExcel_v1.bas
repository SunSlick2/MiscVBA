Sub EmailRangeWithSalutation()

    ' Requires the Microsoft Outlook Object Library reference
    ' Go to Tools -> References in the VBE and check 'Microsoft Outlook [version] Object Library'

    Dim olApp As Object 'Outlook.Application
    Dim olMail As Object 'Outlook.MailItem
    Dim olInsp As Object 'Outlook.Inspector
    Dim olDoc As Object 'Word.Document
    
    Dim rngEmail As Range
    Dim ws As Worksheet
    
    ' --- Email Configuration ---
    Const EMAIL_TO As String = "recipient@example.com" ' Change this to the primary recipient's email address
    Const EMAIL_CC As String = "copy_recipient@example.com" ' Change this to the CC recipient's email address (Use "" for no CC)
    Const EMAIL_TITLE As String = "Weekly Sales Report" ' Change this to your desired subject line
    Const EMAIL_SALUTATION As String = "Hi Team," & "<br><br>Please find the latest sales data for the week below:" ' The new salutation/greeting
    Const EMAIL_SIGNATURE As String = "<br><br>Kind regards,<br>Your Name<br>Your Title" ' Simple HTML signature
    ' ---------------------------

    ' Set the worksheet and the range to be emailed
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet
    Set rngEmail = ws.Range("A2:D28")

    On Error Resume Next ' Handles errors if Outlook is not open or installed
    Set olApp = GetObject(Class:="Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject(Class:="Outlook.Application")
    End If
    On Error GoTo 0 ' Stop error handling

    If olApp Is Nothing Then
        MsgBox "Outlook is not installed or could not be started.", vbCritical
        Exit Sub
    End If

    ' Copy the range as a picture (to ensure formatting is preserved)
    rngEmail.Copy

    ' Create a new email item
    Set olMail = olApp.CreateItem(0) ' 0 = olMailItem

    With olMail
        .Display ' Show the email to the user before sending
        
        ' Set the To, CC, and Subject
        .To = EMAIL_TO
        .CC = EMAIL_CC
        .Subject = EMAIL_TITLE & " - " & Format(Date, "yyyy-mm-dd")
        
        ' Set the body format to HTML to allow for the salutation and signature
        .BodyFormat = 2 ' 2 = olFormatHTML
        
        ' Insert the initial salutation/greeting
        .HTMLBody = EMAIL_SALUTATION
        
        ' Get the Word editor from the mail item to paste the range
        Set olInsp = .GetInspector
        Set olDoc = olInsp.WordEditor

        ' Set the cursor to the end of the current body (after the salutation)
        olDoc.Application.Selection.EndKey Unit:=6 ' 6 = wdStory, moves to end of document
        
        ' Paste the copied range (the range will be pasted as an image/object)
        olDoc.Application.Selection.Paste

        ' Add a space between the pasted image and the signature
        olDoc.Application.Selection.TypeText Text:=" "
        
        ' Append the signature
        .HTMLBody = .HTMLBody & EMAIL_SIGNATURE
        
        ' If you want to send automatically without displaying, uncomment the line below:
        ' .Send 
    End With

    ' Clean up
    Set olDoc = Nothing
    Set olInsp = Nothing
    Set olMail = Nothing
    Set olApp = Nothing

    MsgBox "Email prepared successfully in Outlook!", vbInformation

End Sub
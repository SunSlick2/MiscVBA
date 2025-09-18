Sub ForwardTradeEMailsHTML6()

    Dim myItem As Outlook.mailItem
    Dim Reg1 As RegExp
    Dim Matches As Variant
    Dim Match As Variant

    Dim HTMLDoc As HTMLDocument ' This will be used to manipulate the HTML body
    Dim OftFilePath As String   ' Path for the temporary OFT file

    Dim eMailType As String
    Dim salutation As String
    Dim eMailTo As String
    Dim eMailCC As String

    Dim clientID As String
    
    Dim EAMFlag As Boolean
    Dim ToAddressOnePersonOnly As Boolean
    Dim oAttachement As Outlook.attachment ' Declared once for scope
    Const PR_ATTACH_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x3712000B" ' MAPI_PROPTAG_ATTH_HIDDEN

    Dim olApp As Object
    Dim bXStarted As Boolean
    
    ' Initialize Outlook Application once
    Set olApp = GetObject(, "Outlook.Application")
    If Err <> 0 Then
        Set olApp = CreateObject("Outlook.Application")
        bXStarted = True
    End If
    
    ' Check if there are any items selected
    If ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No emails were selected. Please select one or more emails to forward.", vbInformation
        If bXStarted Then olApp.Quit
        Exit Sub
    End If
    
    ' Loop through each selected email
    For Each myItem In ActiveExplorer.Selection
        Dim SendMailItem As Outlook.mailItem ' Re-declare inside the loop for each item
        
        Set Reg1 = New RegExp ' Initialize RegExp object
        
        ' Initialize HTMLDoc with the original item's body for initial identification
        Set HTMLDoc = CreateObject("htmlfile")
        HTMLDoc.Body.innerHTML = myItem.htmlBody
        
        ' --- Start of Product Type Identification ---
        With Reg1
            .Pattern = "Below shows transaction details for client [0-9]+"
            .Global = True
            Set Matches = .Execute(myItem.Body)
        End With

        If (((myItem.Sender = SENDER_WFX_EMAIL_DOMAIN) Or (myItem.Sender = SENDER_WFX_EMAIL_SHORT)) And (Matches.Count > 0)) Then
            eMailType = "WFX"
            clientID = Right(Matches(0).Value, 6)
        ElseIf (myItem.Sender = SENDER_FXDCONNECT) Then
            eMailType = "FXDC"
            
            Dim allTdsForId As Object
            Dim tdFoundForId As Boolean
            tdFoundForId = False
            clientID = ""
            
            Set allTdsForId = HTMLDoc.getElementsByTagName("td")
            Dim currentTdForId As Object
            
            For Each currentTdForId In allTdsForId
                If InStr(currentTdForId.innerText, "Murex Counterparty ID") > 0 Then
                    Dim parentRowForId As Object
                    Set parentRowForId = currentTdForId.parentElement
                    If Not parentRowForId Is Nothing Then
                        If parentRowForId.cells.Length >= 3 Then
                            clientID = RemoveNonPrintChar(parentRowForId.cells(3 - 1).innerText)
                            tdFoundForId = True
                            Exit For
                        End If
                    End If
                End If
            Next currentTdForId
            
            If Not tdFoundForId Or clientID = "" Then
                MsgBox "FXDC product identified, but 'Murex Counterparty ID' or its corresponding Client ID cell not found in email with subject: " & myItem.Subject & ". Cannot proceed."
                GoTo NextItem ' Skip to the next email
            End If
            
        Else
            MsgBox "Email sender does not match WFX or FXDC for email with subject: " & myItem.Subject
            GoTo NextItem ' Skip to the next email
        End If
        ' --- End of Email Type Identification ---
        
        EAMFlag = False
        Dim eMailDeets As Object
        Set eMailDeets = GetReplyToCCSalutation(clientID)
        
        If eMailDeets Is Nothing Then
            MsgBox "Client details (salutation, To, CC, EAMFlag) could not be retrieved for Client ID: " & clientID & " in email with subject: " & myItem.Subject & ". Cannot proceed."
            GoTo NextItem ' Skip to the next email
        End If
        
        salutation = eMailDeets("Salutation")
        eMailTo = eMailDeets("To")
        eMailCC = eMailDeets("CC")
        EAMFlag = eMailDeets("EAMFlag")
        
        Dim TextonTop As String
        Dim Sig As String
        Sig = ReadSignature("AutoSig.htm")
        Sig = "</p>" & Sig & "<p>"
        
        Dim emailPhrase As String
        If InStr(myItem.Subject, "Indicative") Then
            emailPhrase = " Following are the indicative levels."
        Else
            emailPhrase = " Following are the details of the trade done."
        End If
        
        If (eMailType = "FXDC") Then
            OftFilePath = GLOBAL_ATTACHMENT_FOLDER & "TempEmailTemplate.oft"
            On Error Resume Next
            myItem.SaveAs OftFilePath, olTemplate
            If Err.Number <> 0 Then
                MsgBox "Error saving email as template for subject: " & myItem.Subject & ". Error: " & Err.Description
                On Error GoTo 0
                GoTo NextItem ' Skip to the next email
            End If
            On Error GoTo 0
            
            Set SendMailItem = olApp.CreateItemFromTemplate(OftFilePath)
            
            On Error Resume Next
            Kill OftFilePath
            On Error GoTo 0
            
            Set HTMLDoc = CreateObject("htmlfile")
            HTMLDoc.Body.innerHTML = SendMailItem.htmlBody
            
            TextonTop = salutation & "<BR><BR>" & emailPhrase & " Please let me know if you note any discrepancy." & "<BR><BR>"
            
            Dim divElement As Object
            Set divElement = HTMLDoc.createElement("div")
            divElement.innerHTML = TextonTop
            
            If HTMLDoc.Body.firstChild Is Nothing Then
                HTMLDoc.Body.appendChild divElement
            Else
                HTMLDoc.Body.insertBefore divElement, HTMLDoc.Body.firstChild
            End If
            
            Dim allTds As Object
            Dim currentTd As Object
            Set allTds = HTMLDoc.getElementsByTagName("td")
            
            For Each currentTd In allTds
                If InStr(currentTd.innerText, "(FOR INTERNAL USE ONLY : PLEASE REMOVE 'IB PREMIUM', 'BOOKING UPFRONT' AND 'COUNTERPARTY DEALT' FROM TRADE SUMMARY TABLE BEFORE SENDING TO CLIENTS)") > 0 Then
                    currentTd.innerText = Replace(currentTd.innerText, "(FOR INTERNAL USE ONLY : PLEASE REMOVE 'IB PREMIUM', 'BOOKING UPFRONT' AND 'COUNTERPARTY DEALT' FROM TRADE SUMMARY TABLE BEFORE SENDING TO CLIENTS)", "")
                ElseIf InStr(currentTd.innerText, "Murex Counterparty ID") > 0 Then
                    currentTd.innerText = Replace(currentTd.innerText, "Murex Counterparty ID", "")
                ElseIf InStr(currentTd.innerText, "Counterparty Dealt") > 0 Then
                    ClearParentRowCells currentTd
                ElseIf InStr(currentTd.innerText, "Trade Rationale") > 0 Then
                    ClearParentRowCells currentTd
                ElseIf Not EAMFlag Then
                    If InStr(currentTd.innerText, "IB Premium (Receives // Pays)") > 0 Then
                        ClearParentRowCells currentTd
                    ElseIf InStr(currentTd.innerText, "Booking Upfront") > 0 Then
                        ClearParentRowCells currentTd
                    End If
                End If
            Next currentTd
            
            Dim allImgs As Object
            Dim currentImg As Object
            Set allImgs = HTMLDoc.getElementsByTagName("img")
            For Each currentImg In allImgs
                currentImg.Style.Width = "450px"
                currentImg.Style.Height = "200px"
                currentImg.Width = 450
                currentImg.Height = 200
            Next currentImg
            
            With SendMailItem
                .BodyFormat = olFormatHTML
                .htmlBody = HTMLDoc.Body.innerHTML
                .Subject = myItem.Subject
                .To = eMailTo
                .CC = eMailCC
                
                Dim htmlBodyContent As String
                htmlBodyContent = .htmlBody
                Dim startPos As Long
                Dim endPos As Long
                Dim textToReplace As String
                Dim newHtmlBody As String
                
                startPos = InStr(htmlBodyContent, "Please contact")
                endPos = InStr(htmlBodyContent, "IMPORTANT NOTICE")
                
                If startPos > 0 And endPos > 0 And endPos > startPos Then
                    textToReplace = Mid(htmlBodyContent, startPos, endPos - startPos)
                    newHtmlBody = Replace(htmlBodyContent, textToReplace, Sig)
                    .htmlBody = newHtmlBody
                Else
                    Debug.Print "Signature replacement (Please contact...IMPORTANT NOTICE) markers not found or in incorrect order for subject: " & myItem.Subject
                End If
                
                .Recipients.ResolveAll
            End With
            
        ElseIf eMailType = "WFX" Then
            Set SendMailItem = olApp.CreateItem(olMailItem)
            
            TextonTop = salutation & "<BR><BR>" & " Following are the details of the trade done. Please let me know if you note any discrepancy." & "<BR><BR>"
            
            HTMLDoc.Body.innerHTML = TextonTop & HTMLDoc.Body.innerHTML
            
            Dim TextonBottom As String
            TextonBottom = "</p>" & Sig & "<p class=""disclaimer-font"">Disclaimer</p>"
            
            With SendMailItem
                .BodyFormat = myItem.BodyFormat
                .htmlBody = HTMLDoc.Body.innerHTML
                .Subject = myItem.Subject
                .To = eMailTo
                .CC = eMailCC
                
                .htmlBody = Replace(.htmlBody, "Disclaimer", TextonBottom)
                .Recipients.ResolveAll
            End With
            
            SaveAttachementstoDisk myItem
            For Each oAttachement In myItem.Attachments
                Dim isHidden As Boolean
                On Error Resume Next
                isHidden = oAttachement.propertyAccessor.GetProperty(PR_ATTACH_HIDDEN)
                If Err.Number <> 0 Then
                    isHidden = False
                    Err.Clear
                End If
                On Error GoTo 0
                
                If isHidden = False Then
                    SendMailItem.Attachments.Add GLOBAL_ATTACHMENT_FOLDER & oAttachement.DisplayName
                End If
            Next
            
        Else
            ' This block is now unreachable due to the initial checks
        End If
        
        ApplyMIPLabelConfidentail SendMailItem
        SendMailItem.Display
        
NextItem:
        ' Clean up for the current iteration
        Set SendMailItem = Nothing
        Set Reg1 = Nothing
        Set HTMLDoc = Nothing
        
    Next myItem
    
    ' Ensure Outlook.Application is properly closed if opened by this macro
    If bXStarted And Not olApp Is Nothing Then
        olApp.Quit
    End If
    Set olApp = Nothing

End Sub
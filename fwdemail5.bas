Sub ForwardTradeEMailsHTML5()

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

    Set myItem = ActiveExplorer.Selection.Item(1)
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
        clientID = Right(Matches(0).Value, 6) ' RegExp Matches collection is 0-indexed
    ElseIf (myItem.Sender = SENDER_FXDCONNECT) Then
        eMailType = "FXDC" ' Categorize as general FXDC for now

        ' Dynamically find Client ID from the row where "Murex Counterparty ID" appears
        Dim allTdsForId As Object
        Dim tdFoundForId As Boolean
        tdFoundForId = False
        clientID = "" ' Initialize clientID

        Set allTdsForId = HTMLDoc.getElementsByTagName("td")
        Dim currentTdForId As Object
        For Each currentTdForId In allTdsForId
            If InStr(currentTdForId.innerText, "Murex Counterparty ID") > 0 Then
                Dim parentRowForId As Object
                Set parentRowForId = currentTdForId.parentElement ' Get the parent <tr>

                If Not parentRowForId Is Nothing Then
                    ' Assuming client ID is in the 3rd cell (1-based index) of that row
                    If parentRowForId.cells.Length >= 3 Then
                        ' The RemoveNonPrintChar function is assumed to be available externally
                        clientID = RemoveNonPrintChar(parentRowForId.cells(3 - 1).innerText) ' Convert 1-based to 0-based for cells collection
                        tdFoundForId = True
                        Exit For ' Exit the loop once clientID is found
                    End If
                End If
            End If
        Next currentTdForId

        If Not tdFoundForId Or clientID = "" Then
            MsgBox "FXDC product identified, but 'Murex Counterparty ID' or its corresponding Client ID cell not found. Cannot proceed."
            Exit Sub
        End If

    Else
        MsgBox "Email sender does not match WFX or FXDC."
        Exit Sub
    End If
    ' --- End of Email Type Identification ---

    EAMFlag = False ' Initialize EAMFlag before calling GetReplyToCCSalutation
    ' GetReplyToCCSalutation is assumed to be available externally
    Dim eMailDeets As Object
    Set eMailDeets = GetReplyToCCSalutation(clientID)

    ' Check if eMailDeets is set before accessing its properties
    If eMailDeets Is Nothing Then
        MsgBox "Client details (salutation, To, CC, EAMFlag) could not be retrieved for Client ID: " & clientID & ". Cannot proceed."
        Exit Sub
    End If

    salutation = eMailDeets("Salutation")
    eMailTo = eMailDeets("To")
    eMailCC = eMailDeets("CC")
    EAMFlag = eMailDeets("EAMFlag") ' EAMFlag is determined here

    Dim olApp As Object
    Dim bXStarted As Boolean
    Dim SendMailItem As Outlook.mailItem
    Dim olInsp As Outlook.Inspector
    Dim wdDoc As Object
    Dim oRng As Object

    Set olApp = GetObject(, "Outlook.Application")
        If Err <> 0 Then
            Set olApp = CreateObject("Outlook.Application")
            bXStarted = True
        End If

    Set SendMailItem = olApp.CreateItem(olMailItem) ' Initialize SendMailItem here for attachments later

    Dim TextonTop As String

    ' ReadSignature is assumed to be available externally
    Dim Sig As String
    Sig = ReadSignature("AutoSig.htm")
    Sig = "</p>" & Sig & "<p>"

    Dim emailPhrase As String

    If InStr(myItem.Subject, "Indicative") Then
        emailPhrase = " Following are the indicative levels."
    Else
        emailPhrase = " Following are the details of the trade done."
    End If

    If (eMailType = "FXDC") Then ' Simplified condition to only check for FXDC

        ' --- OFT-based approach for handling HTML and images ---
        OftFilePath = GLOBAL_ATTACHMENT_FOLDER & "TempEmailTemplate.oft"
        On Error Resume Next ' In case file is in use or path issue
        myItem.SaveAs OftFilePath, olTemplate
        If Err.Number <> 0 Then
            MsgBox "Error saving email as template: " & Err.Description
            If bXStarted Then olApp.Quit ' Quit if started by this macro
            Exit Sub
        End If
        On Error GoTo 0

        Set SendMailItem = olApp.CreateItemFromTemplate(OftFilePath)

        ' Delete the temporary OFT file
        On Error Resume Next
        Kill OftFilePath
        On Error GoTo 0
        ' --- End OFT-based approach ---

        ' Re-initialize HTMLDoc with the HTMLBody of the newly created SendMailItem from OFT
        ' All HTML manipulations will happen on this HTMLDoc
        Set HTMLDoc = CreateObject("htmlfile")
        HTMLDoc.Body.innerHTML = SendMailItem.htmlBody


        TextonTop = salutation _
                    & "<BR><BR>" _
                    & emailPhrase & "  Please let me know if you note any discrepancy." _
                    & "<BR><BR>"

        ' Create a new DIV element to hold the TextonTop content
        Dim divElement As Object ' HTMLDivElement
        Set divElement = HTMLDoc.createElement("div")
        divElement.innerHTML = TextonTop

        ' Insert the new DIV element as the first child of the body
        If HTMLDoc.Body.firstChild Is Nothing Then
            HTMLDoc.Body.appendChild divElement
        Else
            HTMLDoc.Body.insertBefore divElement, HTMLDoc.Body.firstChild
        End If

        ' --- Phrase-based Cell Content Clearing Logic ---
        Dim allTds As Object ' HTMLCollection of all <td> elements
        Dim currentTd As Object ' HTMLTableCell

        Set allTds = HTMLDoc.getElementsByTagName("td")

        For Each currentTd In allTds
            If InStr(currentTd.innerText, "(FOR INTERNAL USE ONLY : PLEASE REMOVE 'IB PREMIUM', 'BOOKING UPFRONT' AND 'COUNTERPARTY DEALT' FROM TRADE SUMMARY TABLE BEFORE SENDING TO CLIENTS)") > 0 Then
                currentTd.innerText = Replace(currentTd.innerText, "(FOR INTERNAL USE ONLY : PLEASE REMOVE 'IB PREMIUM', 'BOOKING UPFRONT' AND 'COUNTERPARTY DEALT' FROM TRADE SUMMARY TABLE BEFORE SENDING TO CLIENTS)", "")
            ElseIf InStr(currentTd.innerText, "Murex Counterparty ID") > 0 Then
                currentTd.innerText = Replace(currentTd.innerText, "Murex Counterparty ID", "")
            ElseIf InStr(currentTd.innerText, "Counterparty Dealt") > 0 Then
                ClearParentRowCells currentTd ' Clear entire row
            ElseIf InStr(currentTd.innerText, "Trade Rationale") > 0 Then
                ClearParentRowCells currentTd ' Clear entire row
            ElseIf Not EAMFlag Then ' Only apply if EAMFlag is False
                If InStr(currentTd.innerText, "IB Premium (Receives // Pays)") > 0 Then
                    ClearParentRowCells currentTd ' Clear entire row
                ElseIf InStr(currentTd.innerText, "Booking Upfront") > 0 Then
                    ClearParentRowCells currentTd ' Clear entire row
                End If
            End If
        Next currentTd
        ' --- End Phrase-based Cell Content Clearing Logic ---

        ' --- FIX: Explicitly set image dimensions after OFT conversion ---
        ' Removed currentImg.Style.RemoveProperty calls as they are not supported.
        ' Direct assignment to Style.Width/Height should override existing inline styles.
        Dim allImgs As Object ' HTMLCollection of all <img> elements
        Dim currentImg As Object ' HTMLImgElement

        Set allImgs = HTMLDoc.getElementsByTagName("img")
        For Each currentImg In allImgs
            currentImg.Style.Width = "450px"
            currentImg.Style.Height = "200px"
            currentImg.Width = 450 ' Also set attributes for compatibility
            currentImg.Height = 200
        Next currentImg
        ' --- End FIX ---


        With SendMailItem
            .BodyFormat = olFormatHTML ' Ensure BodyFormat is HTML
            .htmlBody = HTMLDoc.Body.innerHTML ' Update SendMailItem's HTML body with modified HTMLDoc content
            .Subject = myItem.Subject
            .To = eMailTo
            .CC = eMailCC
            
            ' Alternative Signature Replacement Logic
            Dim htmlBodyContent As String
            htmlBodyContent = .htmlBody

            Dim startPos As Long
            Dim endPos As Long
            Dim textToReplace As String
            Dim newHtmlBody As String

            ' Find the starting position of "Please contact"
            startPos = InStr(htmlBodyContent, "Please contact")
            ' Find the starting position of "IMPORTANT NOTICE"
            endPos = InStr(htmlBodyContent, "IMPORTANT NOTICE")

            ' Ensure both markers are found and in the correct order
            If startPos > 0 And endPos > 0 And endPos > startPos Then
                ' The text to replace starts at "Please contact" and ends just before "IMPORTANT NOTICE"
                textToReplace = Mid(htmlBodyContent, startPos, endPos - startPos)
                newHtmlBody = Replace(htmlBodyContent, textToReplace, Sig)
                .htmlBody = newHtmlBody
            Else
                Debug.Print "Signature replacement (Please contact...IMPORTANT NOTICE) markers not found or in incorrect order."
            End If

            .Recipients.ResolveAll
        End With

        ' Attachment manipulation has been removed as OFT creation handles it automatically.

    ElseIf eMailType = "WFX" Then
        ' WFX logic remains largely unchanged as it operates directly on myItem's HTMLDoc
        Set SendMailItem = olApp.CreateItem(olMailItem) ' Create a new mail item for WFX

        TextonTop = salutation _
                    & "<BR><BR>" _
                    & " Following are the details of the trade done. Please let me know if you note any discrepancy." _
                    & "<BR><BR>"

        HTMLDoc.Body.innerHTML = TextonTop & HTMLDoc.Body.innerHTML ' This is operating on the original HTMLDoc derived from myItem

        Dim TextonBottom As String
        TextonBottom = "</p>" _
                    & Sig _
                    & "<p class=""disclaimer-font"">Disclaimer</p>"
        With SendMailItem
            .BodyFormat = myItem.BodyFormat
            .htmlBody = HTMLDoc.Body.innerHTML
            .Subject = myItem.Subject
            .To = eMailTo
            .CC = eMailCC
            .htmlBody = Replace(.htmlBody, "Disclaimer", TextonBottom)
            .Recipients.ResolveAll
        End With

        ' SaveAttachementstoDisk is assumed to be available externally
        SaveAttachementstoDisk myItem
        For Each oAttachement In myItem.Attachments
            Dim isHidden As Boolean
            On Error Resume Next ' Temporarily disable error handling for property access
            isHidden = oAttachement.propertyAccessor.GetProperty(PR_ATTACH_HIDDEN)
            If Err.Number <> 0 Then
                isHidden = False ' Assume not hidden if property not found or errors
                Err.Clear ' Clear the error
            End If
            On Error GoTo 0 ' Re-enable error handling

            If isHidden = False Then ' Add if not hidden
                SendMailItem.Attachments.Add GLOBAL_ATTACHMENT_FOLDER & oAttachement.DisplayName
            End If
        Next

    Else
        MsgBox "eMailType neither WFX nor a configured FXDC product type."
    End If

    ' ApplyMIPLabelConfidentail is assumed to be available externally
    ApplyMIPLabelConfidentail SendMailItem

    SendMailItem.Display

    ' Ensure Outlook.Application is properly closed if opened by this macro
    If bXStarted And Not olApp Is Nothing Then
        olApp.Quit
    End If
    Set olApp = Nothing
    Set SendMailItem = Nothing
    Set myItem = Nothing
    Set Reg1 = Nothing
    Set HTMLDoc = Nothing

End Sub

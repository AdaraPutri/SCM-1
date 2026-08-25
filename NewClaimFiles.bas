Attribute VB_Name = "NewClaimFiles"
Sub ExtractEmailTableDataToDesktop()
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim targetFolder As Object
    Dim mailItem As Object
    Dim oMail As Object
    
    Dim htmlDoc As Object
    Dim htmlCells As Object
    Dim htmlCell As Object
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim rowCount As Long
    
    Dim claimNumber As String
    Dim insuredName As String
    Dim targetFolderName As String

    Dim desktopPath As String
    
    Dim getNextAsClaim As Boolean
    Dim getNextAsInsured As Boolean
    Dim cleanText As String
    
    ' =========================================================================
    ' CONFIGURATION AREA: Enter your exact folder name here
    targetFolderName = "New Claims"
    ' =========================================================================
    
    ' Initialize Excel using Late Binding
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    
    If xlApp Is Nothing Then
        MsgBox "Excel could not be started. Please ensure Excel is installed.", vbCritical
        Exit Sub

    End If
    
    ' Hide Excel during execution for better performance
    xlApp.Visible = False
    filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\New_Claims_List.xlsx"

    If Dir(filePath) <> "" Then
        Set xlWB = xlApp.Workbooks.Open(filePath)
        Set xlWS = xlWB.Sheets(1)
    Else
        Set xlWB = xlApp.Workbooks.Add
        Set xlWS = xlWB.Sheets(1)
        xlWS.cells(1, 1).Value = "Claim Number"
        xlWS.cells(1, 2).Value = "Date Received"
        xlWS.cells(1, 3).Value = "Insured"
        xlWS.cells(1, 4).Value = "Adjuster"
        xlWB.SaveAs filePath
    End If
    
    lastRow = xlWS.cells(xlWS.Rows.Count, 1).End(-4162).Row ' -4162 = xlUp
    If lastRow < 1 Then lastRow = 1
    rowCount = lastRow + 1
    
    ' Connect to the default Outlook MAPI namespace
    Set outlookNamespace = Application.GetNamespace("MAPI")
    Set inboxFolder = outlookNamespace.GetDefaultFolder(6) ' 6 represents the Inbox folder

    
    ' Step 1: Scan for target subfolder nested inside the Inbox
    On Error Resume Next
    Set targetFolder = inboxFolder.Folders(targetFolderName)
    On Error GoTo 0
    
    ' Step 2: Fallback check at the top-level mailbox hierarchy
    If targetFolder Is Nothing Then
        On Error Resume Next
        Set targetFolder = inboxFolder.Parent.Folders(targetFolderName)
        On Error GoTo 0
    End If
    
    ' Safe exit if folder cannot be found
    If targetFolder Is Nothing Then
        xlWB.Close SaveChanges:=False
        xlApp.Quit
        Set xlWS = Nothing
        Set xlWB = Nothing

        Set xlApp = Nothing
        MsgBox "The folder '" & targetFolderName & "' could not be found.", vbCritical
        Exit Sub
    End If
    
    ' Loop through every item inside the target folder
    For Each mailItem In targetFolder.Items
        ' Ensure the item is a standard email message
        If TypeName(mailItem) = "MailItem" Then
            Set oMail = mailItem
            
            ' Only process emails formatted in HTML since they contain tables
            If oMail.BodyFormat = 2 Then ' 2 represents olFormatHTML
                
                ' Reset value variables for the current email
                claimNumber = ""
                insuredName = ""
                getNextAsClaim = False
                getNextAsInsured = False

                
                ' Initialize an HTML document parser object
                Set htmlDoc = CreateObject("htmlfile")
                htmlDoc.Body.innerHTML = oMail.HTMLBody
                
                ' Fetch all table cells (<td> tags) inside the email body
                Set htmlCells = htmlDoc.getElementsByTagName("td")
                
                ' Parse cells sequentially to look for keys
                For Each htmlCell In htmlCells
                    ' If previous cell was a target header, grab this cell's text value
                    If getNextAsClaim Then
                        claimNumber = Trim(htmlCell.innerText)
                        getNextAsClaim = False
                    ElseIf getNextAsInsured Then
                        insuredName = Trim(htmlCell.innerText)
                        getNextAsInsured = False
                    End If
                    

                    ' Clean and normalize the cell text for strict matching rules
                    cleanText = Trim(htmlCell.innerText)
                    cleanText = Replace(cleanText, ":", "")
                    cleanText = Replace(cleanText, Chr(10), "")
                    cleanText = Replace(cleanText, Chr(13), "")
                    cleanText = LCase(Trim(cleanText))
                    
                    ' Verify exact matching parameters
                    If cleanText = "claim number" Then
                        getNextAsClaim = True
                    ElseIf cleanText = "insured" Or cleanText = "insured name" Then
                        getNextAsInsured = True
                    End If
                Next htmlCell
                
                ' Extracts adjuster initials
                adjusterMarker = "Adjuster "
                markerPos = InStr(1, oMail.Body, adjusterMarker, vbTextCompare)
                If markerPos > 0 Then
                    markerPos = markerPos + Len(adjusterMarker)
                    commaPos = InStr(markerPos, oMail.Body, ",")
                    If commaPos > 0 Then
                        adjusterInitials = Trim(Mid(oMail.Body, markerPos, commaPos - markerPos))
                    End If
                End If
                
                
                ' Log results to Excel and check for duplicates
                If claimNumber <> "" Then
                    isDuplicate = False
                    Set duplicateCheck = Nothing
                    Set duplicateCheck = xlWS.Columns(1).Find(What:=claimNumber, LookAt:=1)
                    If Not duplicateCheck Is Nothing Then isDuplicate = True

                    If Not isDuplicate Then
                        xlWS.cells(rowCount, 1).Value = claimNumber
                        xlWS.cells(rowCount, 2).Value = oMail.ReceivedTime
                        xlWS.cells(rowCount, 3).Value = insuredName
                        xlWS.cells(rowCount, 4).Value = adjusterInitials
                        rowCount = rowCount + 1
                    End If
                End If

                Set htmlDoc = Nothing
            End If
        End If
    Next mailItem
                
    
    ' Adjust cell width automatically to fit information nicely
    xlWS.Columns("A:D").AutoFit
    
    ' Save back to the same file it was opened from (or just created)
    xlWB.Save
    xlWB.Close SaveChanges:=True
    xlApp.Quit

    ' Unload objects from memory
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Set targetFolder = Nothing
    Set inboxFolder = Nothing
    Set outlookNamespace = Nothing
    
    MsgBox "Data extraction complete! 'New_Claims_List.xlsx' has been saved to your Desktop.", vbInformation
End Sub

Sub ResetClaimsSheet()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim filePath As String
    Dim lastRow As Long
    Dim confirmReset As Integer

    filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\New_Claims_List.xlsx"

    If Dir(filePath) = "" Then
        MsgBox "No tracking spreadsheet exists yet on the Desktop, so there's nothing to reset.", vbInformation
        Exit Sub
    End If

    confirmReset = MsgBox("This will permanently delete every claim row currently in the tracking sheet " & _
                           "(including any highlight colors you've applied). This cannot be undone." & vbCrLf & vbCrLf & _
                           "Continue?", vbYesNo + vbExclamation, "Confirm Reset")
    If confirmReset <> vbYes Then Exit Sub

    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0
    If xlApp Is Nothing Then
        MsgBox "Excel could not be started. Please ensure Excel is installed.", vbCritical
        Exit Sub
    End If

    xlApp.Visible = False
    Set xlWB = xlApp.Workbooks.Open(filePath)
    Set xlWS = xlWB.Sheets(1)

    lastRow = xlWS.cells(xlWS.Rows.Count, 1).End(-4162).Row

    If lastRow > 1 Then
        xlWS.Rows("2:" & lastRow).Delete
    End If

    xlWB.Save
    xlWB.Close SaveChanges:=True
    xlApp.Quit

    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing

    MsgBox "The tracking sheet has been reset - only the header row remains. Run the extraction macro to repopulate it.", vbInformation
End Sub

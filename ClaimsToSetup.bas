Sub ExtractUniqueSubjectsToDesktop()
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim targetFolder As Object
    Dim mailItem As Object
    Dim oMail As Object
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim rowCount As Long
    Dim lastRow As Long
    Dim filePath As String
    
    Dim targetFolderName As String
    Dim currentSubject As String
    Dim tempSubject As String
    Dim initialLength As Long
    
    Dim duplicateCheck As Object
    Dim isDuplicate As Boolean
    
    ' =========================================================================
    ' CONFIGURATION AREA: Enter your exact folder name here
    targetFolderName = "Claims to Set Up"
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
    
    filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Claims_To_Setup_Data.xlsx"
    
    If Dir(filePath) <> "" Then
        Set xlWB = xlApp.Workbooks.Open(filePath)
        Set xlWS = xlWB.Sheets(1)
    Else
        Set xlWB = xlApp.Workbooks.Add
        Set xlWS = xlWB.Sheets(1)
        xlWS.cells(1, 1).Value = "Email Subject"
        xlWS.cells(1, 2).Value = "Date Received"
        xlWS.cells(1, 3).Value = "File Number"
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
            
            ' Grab and clean the subject line text
            currentSubject = Trim(oMail.Subject)
            
            ' Loop to recursively strip "RE:", "FW:", and "FWD:" prefixes from the beginning
            Do
                tempSubject = LCase(currentSubject)
                initialLength = Len(currentSubject)
                
                If Left(tempSubject, 3) = "re:" Then
                    currentSubject = Trim(Mid(currentSubject, 4))
                ElseIf Left(tempSubject, 3) = "fw:" Then
                    currentSubject = Trim(Mid(currentSubject, 4))
                ElseIf Left(tempSubject, 4) = "fwd:" Then
                    currentSubject = Trim(Mid(currentSubject, 5))
                End If
            Loop While Len(currentSubject) < initialLength And Len(currentSubject) > 0
            
            ' Check for duplicates against everything already in the sheet (past runs included)
            isDuplicate = False
            Set duplicateCheck = Nothing
            Set duplicateCheck = xlWS.Columns(1).Find(What:=currentSubject, LookAt:=1)
            If Not duplicateCheck Is Nothing Then
                isDuplicate = True
            End If
            
            ' If it's a completely unique subject, append it
            If Not isDuplicate Then
                xlWS.cells(rowCount, 1).Value = currentSubject
                xlWS.cells(rowCount, 2).Value = oMail.ReceivedTime
                ' Column 3 (File Number) left blank for manual entry
                rowCount = rowCount + 1
            End If
        End If
    Next mailItem
    
    ' Adjust cell width automatically to fit information nicely
    xlWS.Columns("A:C").AutoFit
    
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
    
    MsgBox "Data extraction complete! 'Claims_To_Setup_Data.xlsx' has been saved to your Desktop.", vbInformation
End Sub

Sub ResetClaimsToSetupSheet()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlWS As Object
    Dim filePath As String
    Dim lastRow As Long
    Dim confirmReset As Integer
    
    filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Claims_To_Setup_Data.xlsx"
    
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
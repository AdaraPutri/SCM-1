Attribute VB_Name = "ClaimsToSetup"
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
    
    Dim targetFolderName As String
    Dim desktopPath As String
    Dim currentSubject As String
    Dim tempSubject As String
    Dim initialLength As Long
    
    Dim duplicateCheck As Object

    Dim isDuplicate As Boolean
    
    Dim debugMsg As String
    
    ' Diagnostic Counters
    Dim totalEmailsScanned As Long
    Dim newRowsAdded As Long
    Dim duplicatesFound As Long
    Dim nonMailItemsFound As Long
    
    ' =========================================================================
    ' CONFIGURATION AREA: Enter your exact folder name here
    targetFolderName = "Claims to Setup CAT"
    ' =========================================================================
    
    ' Initialize counters
    totalEmailsScanned = 0
    newRowsAdded = 0
    duplicatesFound = 0

    nonMailItemsFound = 0
    
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
    
    ' Debug Alert 1: If the folder cannot be found at all

    If targetFolder Is Nothing Then
        debugMsg = "DEBUG ALERT: The folder '" & targetFolderName & "' could not be found anywhere in your Outlook hierarchy. "
        debugMsg = debugMsg & "Please verify the case-sensitive spelling or check if it is nested inside another folder."
        MsgBox debugMsg, vbCritical, "Folder Not Found"
        Exit Sub
    End If
    
    ' Debug Alert 2: If the folder is found but is completely empty
    If targetFolder.Items.Count = 0 Then
        debugMsg = "DEBUG ALERT: The folder '" & targetFolderName & "' was found, but Outlook says there are 0 items inside it. "
        debugMsg = debugMsg & "Please check if the emails are actually sitting inside this folder."
        MsgBox debugMsg, vbExclamation, "Folder Is Empty"
        Exit Sub
    End If
    
    ' Initialize Excel using Late Binding
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0

    
    If xlApp Is Nothing Then
        MsgBox "Excel could not be started. Please ensure Excel is installed.", vbCritical
        Exit Sub
    End If
    
    ' Force Excel to be VISIBLE so you can see it working in real-time
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Add
    Set xlWS = xlWB.Sheets(1)
    
    ' Set up spreadsheet columns
    xlWS.cells(1, 1).Value = "Email Subject"
    xlWS.cells(1, 2).Value = "Date Received"
    
    ' Start writing data on row 2
    rowCount = 2
    
    ' Loop through every item inside the target folder

    For Each mailItem In targetFolder.Items
        ' Ensure the item is a standard email message
        If TypeName(mailItem) = "MailItem" Then
            Set oMail = mailItem
            totalEmailsScanned = totalEmailsScanned + 1
            
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
            
            isDuplicate = False
            
            ' Check for duplicates only if we have already logged entries
            If rowCount > 2 Then
                Set duplicateCheck = Nothing
                
                ' Scan Column A (Email Subject) for a precise exact match
                Set duplicateCheck = xlWS.Columns(1).Find(What:=currentSubject, LookAt:=1)
                
                If Not duplicateCheck Is Nothing Then
                    isDuplicate = True
                End If
            End If
            
            ' If it's a completely unique subject, add it to our list

            If Not isDuplicate Then
                xlWS.cells(rowCount, 1).Value = currentSubject
                xlWS.cells(rowCount, 2).Value = oMail.ReceivedTime
                rowCount = rowCount + 1
                newRowsAdded = newRowsAdded + 1
            Else
                duplicatesFound = duplicatesFound + 1
            End If
        Else
            ' Track if items are calendar invites, read receipts, or delivery reports instead of raw emails
            nonMailItemsFound = nonMailItemsFound + 1
        End If
    Next mailItem
    
    ' Adjust cell width automatically to fit information nicely
    xlWS.Columns("A:B").AutoFit
    
    ' Save the finalized spreadsheet to the user's Desktop with the updated filename
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    xlWB.SaveAs desktopPath & "\Claims_To_Setup_Data.xlsx"
    xlWB.Close SaveChanges:=True
    xlApp.Quit
    
    ' Unload objects from memory
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Set targetFolder = Nothing
    Set inboxFolder = Nothing
    Set outlookNamespace = Nothing
    
    ' Display the final summary report box
    debugMsg = "Data extraction complete!" & vbCrLf & vbCrLf
    debugMsg = debugMsg & "• Total raw items in folder: " & (totalEmailsScanned + nonMailItemsFound) & vbCrLf
    debugMsg = debugMsg & "• Valid emails processed: " & totalEmailsScanned & vbCrLf
    debugMsg = debugMsg & "• Unique subjects saved: " & newRowsAdded & vbCrLf
    debugMsg = debugMsg & "• Duplicates discarded: " & duplicatesFound & vbCrLf
    debugMsg = debugMsg & "• Non-email items skipped: " & nonMailItemsFound & vbCrLf & vbCrLf

    debugMsg = debugMsg & "Saved to your Desktop as 'Claims_To_Setup_Data.xlsx'"
    
    MsgBox debugMsg, vbInformation, "Execution Summary"
End Sub

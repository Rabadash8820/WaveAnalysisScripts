Attribute VB_Name = "AnalyzeInterface"
Option Explicit

Private Const TIME_TAKEN_STR = "Time taken (hh:mm:ss)> "
Private Const COMPLETED_MESSAGE = "I hope you uncommented any functions from when you were debugging!"

Private wabsOnly As Boolean
Private folderContents As New Dictionary

Public Sub ProcessExistingWorkbook()
    Dim result As VbMsgBoxResult
    Dim unitNames As Variant
    Dim numRecs As Integer
    
    'Let the user pick the workbook to analyze
    Dim wbName As String
    wbName = PickWorkbook("Select an existing tissue summary workbook")
    If wbName = "" Then
        result = MsgBox("No workbook selected.", vbOKOnly)
        Exit Sub
    End If
    
    'Process the selected workbook (it will be left open)
    Call processTissueWorkbook(wbName, numRecs)
            
    'Display log information (time taken and how many sheets were processed)
    Dim log As New LogForm
    With log.MainListBox
        Dim tempStr As String
        tempStr = numRecs & " recordings analyzed in " & wbName
        .AddItem tempStr
        .AddItem ""
        .AddItem TIME_TAKEN_STR & Format(ProgramDuration(), "hh:mm:ss")
        .AddItem COMPLETED_MESSAGE
    End With
    log.Show
End Sub
Public Sub ProcessExistingWorkbookFolder()
    Dim dialog As FileDialog
    Dim rootFolder  As Folder
    Dim result As VbMsgBoxResult
        
    'Create the folder-selection dialog box
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "Select directory with summary workbooks (all subdirectories will also be processed)"
    dialog.AllowMultiSelect = False
    
    'If the user didn't select anything, display a message and return
    If dialog.Show = False Then
        result = MsgBox("No folder selected.", vbOKOnly, "Routine complete")
        Exit Sub
    End If
    
    'Store the selected directory's path
    Dim fileSystem As New FileSystemObject
    Set rootFolder = fileSystem.GetFolder(dialog.SelectedItems(1))
                
    'Process each selected workbook
    'If none were found, then just display a message and return
    Dim numFiles As Integer
    numFiles = 0
    folderContents.RemoveAll    'I guess this is necessary since the Dictionary has Module scope? ...
    Call processDir(rootFolder, rootFolder.name, numFiles)
    If numFiles = 0 Then
        result = MsgBox("No workbooks found to process.", vbOKOnly, "Routine complete")
        Exit Sub
    End If
    
    'Show a Log form and clean up
    Call logFolderResults(ProgramDuration(), folderContents)

End Sub

Private Sub processDir(ByRef dir As Folder, ByVal pathAbove As String, ByRef numFiles As Integer)
    Dim wkbk As File
    Dim tempNumFiles, numRecs As Integer
    
    'Process each Workbook in the Folder
    tempNumFiles = 0
    For Each wkbk In dir.Files
        If wkbk.Type = "Microsoft Excel Worksheet" Then
            tempNumFiles = tempNumFiles + 1
            
            'Record how many recordings in this workbook were analyzed
            Call processTissueWorkbook(wkbk.path, numRecs)
            folderContents.Add pathAbove & "\" & wkbk.name, numRecs

            'Save/close the workbook
            Application.DisplayAlerts = False
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        End If
    Next wkbk
    
    'Recursively process Workbooks in all subdirectories as well
    Dim subDir As Folder
    For Each subDir In dir.SubFolders
        Call processDir(subDir, pathAbove & "\" & subDir.name, numFiles)
    Next subDir
    
    'Increment the total number of Files that have been processed
    numFiles = numFiles + tempNumFiles
End Sub

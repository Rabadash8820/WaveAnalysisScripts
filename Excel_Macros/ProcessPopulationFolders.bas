Attribute VB_Name = "ProcessPopulationFolders"
Option Explicit
Option Private Module

Private Const TIME_GENERATED_STR = "Time Generated"
Private Const TIME_TAKEN_STR = "Time taken (hh:mm:ss)> "
Private Const COMPLETED_MESSAGE = "Don't forget to add start/end times to each of the generated workbooks!"

Private Const TIME_COL = 5

Public Sub ProcessPopulationFolders()
    Dim dialog As FileDialog
    Dim rootFolder, popFolder, retFolder As Folder
    Dim recording As File
    Dim result As VbMsgBoxResult
    Dim retinaWb As Workbook
    Dim wbName As String
    Dim numFiles As Integer
    Dim tempStr As String
    
    Call setupOptimizations
        
    'Create the folder-selection dialog box
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "Select root directory (all subdirectories will also be processed)"
    dialog.AllowMultiSelect = False
    
    'If the user didn't select anything, display a message and return
    If dialog.Show = False Then
        result = MsgBox("No folder selected.", vbOKOnly, "Routine complete")
        GoTo ExitSub
    End If
    
    'Store the selected directory's path
    Dim fileSystem As New FileSystemObject
    Set rootFolder = fileSystem.GetFolder(dialog.SelectedItems(1))
    
    'Loop through each population folder in this root directory
    Dim folderContents As New Dictionary
    For Each popFolder In rootFolder.SubFolders
    
        'For each retinal folder in this population folder
        For Each retFolder In popFolder.SubFolders
                
            'Initialize a new summary workbook for this retina
            Set retinaWb = Workbooks.Add
                        
            'Add each text file to the Contents sheet and load them on a new sheet
            Call addContentsSheet
            numFiles = 0
            For Each recording In retFolder.Files
                If recording.Type = "TXT File" Or recording.Type = "Text Document" Then
                    numFiles = numFiles + 1
                    Call openFile(recording, numFiles)
                End If
            Next recording
            folderContents.Add retFolder.path, numFiles
            
            'If no files were found, just display a message and return
            If numFiles = 0 Then
                result = MsgBox("No files found to process.", vbOKOnly, "Routine complete")
                Application.DisplayAlerts = False
                retinaWb.Close
                Application.DisplayAlerts = True
                GoTo ExitSub
            End If
        
            'Finalize the Contents sheet
            With Worksheets(CONTENTS_SHEET_NAME)
                .Cells.VerticalAlignment = xlCenter
                .Cells.HorizontalAlignment = xlLeft
                .Columns.EntireColumn.AutoFit
                .Rows.EntireRow.AutoFit
            End With

            'Save the workbook as popFolderPath\retFolderName.xlsx (overwriting any previous one)
            wbName = popFolder.path & "\" & retFolder.name
            Application.DisplayAlerts = False
            retinaWb.SaveAs fileName:=wbName, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
            retinaWb.Close
            Application.DisplayAlerts = True
            
        Next retFolder
        
    Next popFolder
    
    'Show a Log form and clean up
    Call logFolderResults(ProgramDuration(), folderContents)

ExitSub:
    Call tearDownOptimizations
End Sub

Private Sub addContentsSheet()
    'Initialize Contents sheet
    ActiveSheet.name = CONTENTS_SHEET_NAME
    
    'Add the time generated info
    Dim timeGenRng As Range
    Set timeGenRng = Cells(1, 1)
    timeGenRng.offset(0, 0).value = TIME_GENERATED_STR
    timeGenRng.offset(0, 0).Font.Bold = True
    timeGenRng.offset(1, 0).value = Now
    timeGenRng.offset(1, 0).NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
    
    'Add the other summary info...
    Dim infoCell As Range
    Set infoCell = timeGenRng.offset(3, 0)
    infoCell.offset(0, 0).value = "FileName"
    infoCell.offset(0, 1).value = "SheetName"
    infoCell.offset(0, 2).value = "StartTime"
    infoCell.offset(0, 3).value = "EndTime"
        
    '...and put it in a table
    Dim summaryTbl As ListObject
    Set summaryTbl = ActiveSheet.ListObjects.Add(xlSrcRange, infoCell.CurrentRegion, , xlYes)
    summaryTbl.name = "SummaryTbl"
    
    'Delete any extra sheets if this workbook was generated in Excel 2010 or earlier
    Application.DisplayAlerts = False
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Worksheets
        If sh.name <> ActiveSheet.name Then _
            sh.Delete
    Next sh
    Application.DisplayAlerts = True
End Sub

Private Sub openFile(recording As File, ByVal fileIndex As Integer)
    Dim header As String
    Dim col, numCols As Integer
    Dim numValues As Long
    Dim firstBadCell As Range
    Dim nameCell As Range
    
    'Add this recording to the Contents sheet
    Dim summaryTbl As ListObject, rng As Range
    Set summaryTbl = Worksheets(CONTENTS_SHEET_NAME).ListObjects("SummaryTbl")
    Set rng = summaryTbl.ListRows.Add.Range
    rng.Cells(1, 1) = recording.name
    rng.Cells(1, 2) = RECORDING_STR & fileIndex

    'Load data into a new sheet of the new workbook and format it
    Worksheets.Add After:=Sheets(Worksheets.Count)
    ActiveSheet.name = RECORDING_STR & fileIndex
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & recording.path, Destination:=Cells(1, 1))
        .name = RECORDING_STR & fileIndex
        .FieldNames = True
        .RefreshOnFileOpen = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = True
        .Refresh
    End With
    ActiveSheet.Rows(1).Font.Bold = True
            
    'Delete columns for the A1 electrode (if it exists) and the All File interval
    numCols = Cells(1, 1).End(xlToRight).Column
    For col = 1 To numCols
'        Cells(1, col).Select
        header = Cells(1, col).value
        If InStr(1, header, "A1") Or InStr(1, header, "AllFile") Then
            Columns(col).Delete
            col = col - 1
            numCols = numCols - 1
        End If
    Next
    
    'Delete cells with spaces generated by NeuroExplorer
    For col = 1 To numCols
        numValues = WorksheetFunction.Count(Columns(col)) + 1   '+1 is for the unit headers
        Set firstBadCell = Cells(1, col).offset((numValues + 1) - 1, 0)
        Range(firstBadCell, firstBadCell.End(xlDown)).Delete shift:=xlUp
    Next col
    ActiveSheet.UsedRange   'Refresh used range by getting this property
End Sub

Public Sub logFolderResults(ByVal duration As Double, ByRef folderContents As Dictionary)
    Dim tempStr As String
    
    'Display log information (time taken and how many files were processed per folder)
    Dim f As Variant
    Dim log As New LogForm
    With log.MainListBox
        For Each f In folderContents.Keys
            tempStr = folderContents(f) & " recordings processed in " & f
            .AddItem tempStr
        Next f
        .AddItem ""
        .AddItem TIME_TAKEN_STR & Format(duration, "hh:mm:ss")
        .AddItem COMPLETED_MESSAGE
    End With
    
    log.Show
    
End Sub

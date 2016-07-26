Attribute VB_Name = "ProcessPopulations"
Option Explicit

Private Const TIME_COL = 5

Dim tissueWbs As New Dictionary

Public Sub ProcessPopulations()
    Call setupOptimizations
    Dim outputStrs As New Collection
    
    'Define the Tissue/Recording/Population objects, etc.
    Dim Success As Boolean
    Call GetConfigVars
    Success = DefineObjects()
    If Not Success Then _
        GoTo ExitSub
    
    'Load each provided RecordingView's text files
    'If any errors occur, log their messages and exit
    Set outputStrs = checkTextFiles
    If outputStrs.Count > 0 Then _
        GoTo ExitSub
        
    'Define the types of bursts to use
    Dim burstUseTypes As New Dictionary
    burstUseTypes.Add "WAB", BurstUseType.WABs
    burstUseTypes.Add "NonWAB", BurstUseType.NonWABs
    
    'Load each provided RecordingView's text files,
    'then perform wave analyses on each TissueView!
    Dim fs As New FileSystemObject
    Dim p As Integer, r As Integer, t As Integer, bt As Integer, bType As String, wbPath As String
    Dim pop As cPopulation, rv As cRecordingView, tv As cTissueView
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        For t = 1 To pop.TissueViews.Count
            Set tv = pop.TissueViews.item(t)
            For r = 1 To tv.RecordingViews.Count
                Set rv = tv.RecordingViews.item(r)
                Call loadRecording(rv, r)
            Next r
            For bt = 1 To UBound(BURST_TYPES, 2)
                bType = BURST_TYPES(1, bt)
                wbPath = tv.WorkbookPaths(bType)
                Call processTissueWorkbook(wbPath, tv.Tissue, burstUseTypes(bType))
            Next bt
        Next t
    Next p
    
    'Combine data into a single workbook
    Dim combineWb As Workbook
    Set combineWb = Workbooks.Add
    Call CombineDataIntoWorkbook(combineWb)
    
    'If requested, remove bursts with durations of zero from the totals
    If MARK_BURST_DUR_UNITS Then _
        Call DeleteZeroBurstDurUnits(combineWb)
    
    'Remove the no-longer-needed Tissue summary workbooks
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        For t = 1 To pop.TissueViews.Count
            Set tv = pop.TissueViews.item(t)
            For bt = 1 To UBound(BURST_TYPES, 2)
                bType = BURST_TYPES(1, bt)
                wbPath = tv.WorkbookPaths(bType)
                If fs.FileExists(wbPath) Then
                    fs.DeleteFile (wbPath)
                End If
            Next bt
        Next t
    Next p
    
    'Store the success strings to be logged
    Set outputStrs = successStrings

ExitSub:
    'Log results (errors or success) and tear things down
    Call showLog(outputStrs)
    Call tearDownOptimizations
End Sub

Private Sub loadRecording(ByRef rv As cRecordingView, ByVal rvIndex As Integer)
    Dim pop As cPopulation, tv As cTissueView
    
    'For each burst type...
    Dim t As Integer, bType As String, wbPath As String, wb As Workbook
    Set tv = rv.TissueView
    For t = 1 To UBound(BURST_TYPES, 2)
        bType = BURST_TYPES(1, t)
        
        'Open the summary workbook for this recording's tissue (replacing any old one)
        wbPath = tv.WorkbookPaths(bType)
        Dim fs As New FileSystemObject
        If fs.FileExists(wbPath) Then _
            fs.DeleteFile (wbPath)
        Set wb = Workbooks.Add
        Call addContentsSheet
        wb.SaveAs (wbPath)
        
        'Add each text file to the Contents sheet and load them on a new sheet
        Dim txtFile As File
        Set txtFile = fs.GetFile(rv.TextPath)
        Call openFile(rv, txtFile)
                       
        'Clean up and close the summary workbook
        With Worksheets(CONTENTS_NAME)
            .Cells.VerticalAlignment = xlCenter
            .Cells.HorizontalAlignment = xlLeft
            .Columns.EntireColumn.AutoFit
            .Rows.EntireRow.AutoFit
        End With
        wb.Close (True)
    Next t

End Sub

Private Sub addContentsSheet()
    'Initialize Contents sheet
    ActiveSheet.Name = CONTENTS_NAME
    
    'Add the time generated info
    Dim timeGenRng As Range
    Set timeGenRng = Cells(1, 1)
    timeGenRng.offset(0, 0).Value = TIME_GENERATED_STR
    timeGenRng.offset(0, 0).Font.Bold = True
    timeGenRng.offset(1, 0).Value = Now
    timeGenRng.offset(1, 0).NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
    
    'Add the other summary info...
    Dim infoCell As Range
    Set infoCell = timeGenRng.offset(3, 0)
    infoCell.offset(0, 0).Value = "FileName"
    infoCell.offset(0, 1).Value = "SheetName"
    infoCell.offset(0, 2).Value = "StartTime"
    infoCell.offset(0, 3).Value = "EndTime"
        
    '...and put it in a table
    Dim contentsTbl As ListObject
    Set contentsTbl = ActiveSheet.ListObjects.Add(xlSrcRange, infoCell.CurrentRegion, , xlYes)
    contentsTbl.Name = CONTENTS_NAME
    
    'Delete any extra sheets if this workbook was generated in Excel 2010 or earlier
    Application.DisplayAlerts = False
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Worksheets
        If sh.Name <> ActiveSheet.Name Then _
            sh.Delete
    Next sh
    Application.DisplayAlerts = True
End Sub

Private Sub openFile(ByRef rec As cRecordingView, ByRef recFile As File)
    Dim header As String
    Dim col, numCols As Integer
    Dim numValues As Long
    Dim firstBadCell As Range
    Dim nameCell As Range
    
    'Add this recording to the Contents sheet
    Dim contentsTbl As ListObject, rng As Range
    Set contentsTbl = Worksheets(CONTENTS_NAME).ListObjects(CONTENTS_NAME)
    Set rng = contentsTbl.ListRows.Add.Range
    rng.Cells(1, 1) = recFile.Name
    rng.Cells(1, 2) = RECORDING_STR & Worksheets.Count
    rng.Cells(1, 3) = rec.Recording.StartTime
    rng.Cells(1, 4) = rec.Recording.StartTime + rec.Recording.Duration

    'Load data into a new sheet of the new workbook and format it
    Worksheets.Add After:=Sheets(Worksheets.Count)
    ActiveSheet.Name = RECORDING_STR & Worksheets.Count - 1
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & recFile.path, Destination:=Cells(1, 1))
        .Name = RECORDING_STR & Worksheets.Count - 1
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
        header = Cells(1, col).Value
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
        Range(firstBadCell, firstBadCell.End(xlDown)).Delete Shift:=xlUp
    Next col
    ActiveSheet.UsedRange   'Refresh used range by getting this property
End Sub

Private Function checkTextFiles() As Collection
    
    'Load each provided RecordingView's text files (if they exist)
    'then perform wave analyses on each TissueView!
    Dim fs As New FileSystemObject, unfound As New Collection, notGiven As New Collection
    Dim p As Integer, r As Integer, t As Integer, bt As Integer, bType As String, wbPath As String
    Dim pop As cPopulation, rv As cRecordingView, tv As cTissueView
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        For t = 1 To pop.TissueViews.Count
            Set tv = pop.TissueViews.item(t)
            For r = 1 To tv.RecordingViews.Count
                Set rv = tv.RecordingViews.item(r)
                If rv.TextPath = "" Then
                    notGiven.Add "Recording " & rv.Recording.ID & " in Population """ & rv.TissueView.Population.Name & """"
                ElseIf Not fs.FileExists(rv.TextPath) Then
                    unfound.Add "Recording " & rv.Recording.ID & " in Population """ & rv.TissueView.Population.Name & """  (" & rv.TextPath & ")"
                End If
            Next r
        Next t
    Next p
    
    'If there was an error opening any of the text files then store error messages
    Dim errorOccurred As Boolean, item As Variant, errorStrs As New Collection
    errorOccurred = (unfound.Count > 0 Or notGiven.Count > 0)
    If errorOccurred Then
        errorStrs.Add "Please correct the following errors before running again."
        If unfound.Count > 0 Then
            errorStrs.Add ""
            errorStrs.Add "The provided text files could not be found for the following Recordings:"
            For Each item In unfound
                errorStrs.Add "     " & item
            Next item
        End If
        If notGiven.Count > 0 Then
            errorStrs.Add ""
            errorStrs.Add "No text file was provided for the following Recordings:"
            For Each item In notGiven
                errorStrs.Add "     " & item
            Next item
        End If
    End If

    'Return these error messages
    Set checkTextFiles = errorStrs
End Function

Private Function successStrings() As Collection
    Dim outputStrs As New Collection, numRecs As Integer
    Dim p As Integer, t As Integer, r As Integer
    Dim pop As cPopulation, tv As cTissueView, rv As cRecordingView
        
    'For each Tissue of each Population...
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        outputStrs.Add "Tissues loaded for population " & pop.Name & ":"
        For t = 1 To pop.TissueViews.Count
            
            'Add which of its Recordings were successfuly loaded
            Set tv = pop.TissueViews.item(t)
            numRecs = tv.RecordingViews.Count
            outputStrs.Add "    Attempted to load " & numRecs & " recording" & IIf(numRecs = 1, "", "s") & " in Tissue " & tv.Tissue.ID
            For r = 1 To tv.RecordingViews.Count
                Set rv = tv.RecordingViews.item(r)
                outputStrs.Add "        " & "Recording " & rv.Recording.ID & " successfully loaded"
            Next r
            
        Next t
        outputStrs.Add ""
    Next p

    'Return these error messages
    Set successStrings = outputStrs
End Function

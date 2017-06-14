Attribute VB_Name = "Main"
Option Explicit

Private Const TIME_COL = 5

Dim tissueWbs As New Dictionary

Public Sub ProcessPopulations()
    Call setupOptimizations
    Dim outputStrs As New Collection
    
    'Define the Tissue/Recording/Population objects, etc.
    Dim success As Boolean
    Call GetConfigVars
    Call DefineObjects(outputStrs)
    If outputStrs.Count > 0 Then _
        GoTo Finally
    
    'Load each provided RecordingView's text files
    'If any errors occur, log their messages and exit
    Call checkTextFilesExist(outputStrs)
    If outputStrs.Count > 0 Then _
        GoTo Finally
    
    'Load each provided RecordingView's text files,
    'then perform wave analyses on each TissueView!
    Dim fs As New FileSystemObject
    Dim p As Integer, r As Integer, t As Integer, bType As Variant, wbPath As String
    Dim pop As cPopulation, rv As cRecordingView, tv As cTissueView
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        For t = 1 To pop.TissueViews.Count
            Set tv = pop.TissueViews.item(t)
            For r = 1 To tv.RecordingViews.Count
                Set rv = tv.RecordingViews.item(r)
                success = loadRecording(rv, r)
                If Not success Then
                    outputStrs.Add "Recording " & rv.Recording.ID & " from Tissue """ & rv.TissueView.Tissue.Name & """ did not contain any burst start/end timestamps."
                    outputStrs.Add "Make sure that you exported Interval data from NeuroExplorer for EVERY Recording's text files before running again."
                    GoTo Finally
                End If
            Next r
            For Each bType In BURST_TYPES.Keys()
                wbPath = tv.WorkbookPaths(bType)
                Call AnalyzeTissueWorkbook(wbPath, tv, bType)
            Next bType
        Next t
    Next p
    
    'Combine data into a single workbook
    Dim combineWb As Workbook
    Set combineWb = Workbooks.Add
    Call CombineDataIntoWorkbook(combineWb)
    
    'If requested, delete bursts with bad durations
    'Also delete any remaining units marked for "exclusion"
    If EXCLUDE_BURST_DUR_UNITS Then _
        Call DeleteZeroBurstDurUnits(combineWb)
    Call ExcludeUnits(combineWb)
    
    'Remove the no-longer-needed Tissue summary workbooks
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        For t = 1 To pop.TissueViews.Count
            Set tv = pop.TissueViews.item(t)
            For Each bType In BURST_TYPES.Keys()
                wbPath = tv.WorkbookPaths(bType)
                If fs.FileExists(wbPath) Then
                    fs.DeleteFile (wbPath)
                End If
            Next bType
        Next t
    Next p
    
    'Store the success strings to be logged
    Set outputStrs = successStrings
    
    GoTo Finally

Finally:
    'Log results (errors or success) and tear things down
    Call showLog(outputStrs)
    Call tearDownOptimizations
End Sub

Private Function loadRecording(ByRef rv As cRecordingView, ByVal rvIndex As Integer) As Boolean
    Dim pop As cPopulation, tv As cTissueView
    
    'For each burst type...
    Dim bType As Variant, wbPath As String, wb As Workbook
    Set tv = rv.TissueView
    For Each bType In BURST_TYPES.Keys()
        'Open the summary workbook for this recording's tissue (replacing any old one)
        wbPath = tv.WorkbookPaths(bType)
        Dim fs As New FileSystemObject
        If fs.FileExists(wbPath) Then _
            fs.DeleteFile (wbPath)
        Set wb = Workbooks.Add
        Call addContentsSheet
        wb.SaveAs (wbPath)
        
        'Add each text file to the Contents sheet and open them onto a new sheet
        'If the text file could not be opened then return with an error
        Dim txtFile As File, success As Boolean
        Set txtFile = fs.GetFile(rv.TextPath)
        success = openFile(rv, txtFile)
        If Not success Then
            loadRecording = False
            wb.Close SaveChanges:=False
            fs.DeleteFile (wbPath)
            Exit Function
        End If
                       
        'Clean up and close the summary workbook
        With Worksheets(CONTENTS_NAME)
            .Cells.VerticalAlignment = xlCenter
            .Cells.HorizontalAlignment = xlLeft
            .Columns.EntireColumn.AutoFit
            .Rows.EntireRow.AutoFit
        End With
        wb.Close (True)
    Next bType

    loadRecording = True
End Function

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

Private Function openFile(ByRef rec As cRecordingView, ByRef recFile As File) As Boolean
    'Add this recording to the Contents sheet
    Dim contentsTbl As ListObject, rng As Range
    Set contentsTbl = Worksheets(CONTENTS_NAME).ListObjects(CONTENTS_NAME)
    Set rng = contentsTbl.ListRows.Add.Range
    rng.Cells(1, 1) = recFile.Name
    rng.Cells(1, 2) = RECORDING_STR & Worksheets.Count
    rng.Cells(1, 3) = rec.Recording.startTime
    rng.Cells(1, 4) = rec.Recording.startTime + rec.Recording.Duration
    
    Dim success As Boolean
    On Error GoTo Finally

    'Load data into a new sheet of the new workbook and format it
    success = False
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
        .TextFileCommaDelimiter = False
        .TextFileConsecutiveDelimiter = True
        .TextFileOtherDelimiter = ""
        .TextFileSemicolonDelimiter = False
        .TextFileSpaceDelimiter = False
        .Refresh
    End With
    ActiveSheet.Rows(1).Font.Bold = True
    
    'Only continue if there are burst timestamp columns
    Dim header As String, numCols As Integer, col As Integer
    numCols = Cells(1, 1).End(xlToRight).Column
    For col = 1 To numCols
        header = Cells(1, col).Value
        If InStr(1, header, "burst") Then
            success = True
            Exit For
        End If
    Next
    success = True
'    If Not success Then _
'        GoTo Finally
            
    'Delete columns for the A1 electrode (if it exists) and the All File interval
    For col = 1 To numCols
        header = Cells(1, col).Value
        If InStr(1, header, "A1") Or InStr(1, header, "AllFile") Then
            Columns(col).Delete
            col = col - 1
            numCols = numCols - 1
        End If
    Next
    
    'Delete cells with spaces generated by NeuroExplorer
    Dim numValues As Long, firstBadCell As Range
    For col = 1 To numCols
        numValues = WorksheetFunction.Count(Columns(col)) + 1   '+1 is for the unit headers
        Set firstBadCell = Cells(1, col).offset((numValues + 1) - 1, 0)
        Range(firstBadCell, firstBadCell.End(xlDown)).Delete Shift:=xlUp
    Next col
    ActiveSheet.UsedRange   'Refresh used range by getting this property
    
    GoTo Finally
    
Catch:
    success = False
    GoTo Finally
    
Finally:
    openFile = success
    
End Function

Private Sub checkTextFilesExist(ByRef errorStrs As Collection)
    
    'Check whether each provided RecordingView has a text file that exists
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
    
    'If any of them do not then return some error messages
    Dim errorOccurred As Boolean, item As Variant
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
    
End Sub

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
            outputStrs.Add "    Attempted to load " & numRecs & " recording" & IIf(numRecs = 1, "", "s") & " in Tissue " & tv.Tissue.Name
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

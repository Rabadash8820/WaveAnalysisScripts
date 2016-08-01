Attribute VB_Name = "GlobalThings"
Option Explicit
Option Private Module

'ABSTRACT DATA TYPES
Public Enum BurstUseType
    All
    WABs
    NonWABs
End Enum
Public Enum ReportStatsType
    MeanSEM
    MedianIQR
End Enum

'CONFIG VARIABLES
Public MEA_ROWS, MEA_COLS As Integer
Public NUM_CHANNELS As Integer
Public Const GROUND_CHANNEL = 4     'adch_15
Public Const MAX_UNITS_PER_CHANNEL = 10
Public MAX_POSSIBLE_UNITS As Integer
Public CORRELATION_DT As Double              'seconds
Public NUM_BINS, MIN_BINS As Double
Public MIN_ASSOC_UNITS As Integer
Public MIN_DURATION As Double
Public MAX_DURATION As Double

'SHEET/TABLE NAMES
'Since these could be table names, they should use only alphanumeric characters and NO spaces
Public Const POPS_NAME = "Populations"
Public Const TISSUES_NAME = "Tissues"
Public Const ANALYZE_NAME = "Analyze"
Public Const COMBINE_NAME = "Combine_Results"
Public Const INVALIDS_NAME = "Unit_Removal"
Public Const CONFIG_NAME = "Config"
Public Const CONTENTS_NAME = "Contents"
Public Const ALL_AVGS_NAME = "All_Avgs"
Public Const BURST_AVGS_NAME = "Burst_Avgs"
Public Const STATS_NAME = "Stats"
Public Const PROPERTIES_NAME = "Properties"
Public Const STTC_NAME = "STTC"
Public Const RECORDINGS_NAME = "Recordings"
Public Const RECORDING_VIEWS_NAME = "Recordings"

'STRINGS
Public Const TIME_GENERATED_STR = "Time Generated"
Public Const CELL_STR = "Cell"
Public Const RECORDING_STR = "Recording_"
Public Const STTC_HEADER_STR = "Spike Time Tiling Coefficient Values for Every Cell Pair"
Public Const INTER_ELECTRODE_DIST_STR = "Inter-Electrode Distance"
Public Const STTC_STR = "STTC"
Public Const CHANNEL_PREFIX = "adch_"

'FLAGS
Public ASSOC_SAME_CHANNEL_UNITS As Boolean
Public ASSOC_MULTIPLE_UNITS As Boolean
Public DELETE_BAD_SPIKES As Boolean
Public DELETE_BAD_BURSTS As Boolean
Public EXCLUDE_BURST_DUR_UNITS As Boolean
Public DATA_PAIRED As Boolean
Public REPORT_PROPS_TYPE As ReportStatsType
Public REPORT_STTC_TYPE As ReportStatsType

'Arrays/collections and associated values
Public NUM_PROPERTIES As Integer
Public NUM_BKGRD_PROPERTIES As Integer
Public NUM_BURST_PROPERTIES As Integer
Public PROPERTIES() As String
Public PROP_UNITS() As String
Public BURST_TYPES As New Dictionary
Public CTRL_POP As cPopulation
Public POPULATIONS As New Dictionary
Public TISSUES As New Dictionary
Public Recordings As New Dictionary

'OTHER VALUES
Public Const MAX_EXCEL_ROWS = 1048576

Public Function DefineObjects() As Boolean
    Dim Success As Boolean
    Success = False
    
    'Open the Data Summary workbook
    Dim summaryFile As File, result As VbMsgBoxResult, summaryWb As Workbook
    Set summaryFile = PickWorkbook("Select the Data Summary workbook")
    If summaryFile Is Nothing Then
        result = MsgBox("No workbook selected.", vbOKOnly, "Routine complete")
        GoTo ExitFunc
    End If
    Set summaryWb = Workbooks.Open(summaryFile.path)
    
    'Open the Population-definition workbook
    Dim popFile As File, popWb As Workbook
    Set popFile = PickWorkbook("Select the workbook that defines your experimental populations")
    If popFile Is Nothing Then
        result = MsgBox("No workbook selected.", vbOKOnly, "Routine complete")
        GoTo ExitFunc
    End If
    Set popWb = Workbooks.Open(popFile.path)
    
    'Get Tissue/Recording and experimental Population info
    'Wrap these objects in Views associated with the appropriate Population
    Call defineRecordings(summaryWb)
    Call definePopulations(summaryWb)
    Call definePopulationViews(popWb)
    Call defineUnits(summaryWb)
    Application.DisplayAlerts = False
    summaryWb.Close
    popWb.Close
    Application.DisplayAlerts = True
    
    Success = True
    
ExitFunc:
    DefineObjects = Success
End Function

Private Sub defineRecordings(ByRef summaryWb As Workbook)
    'Make sure parent objects are defined first
    Call defineTissues(summaryWb)

    'Get the Recordings table
    Dim recSht As Worksheet, recTbl As ListObject
    Set recSht = summaryWb.Worksheets(RECORDINGS_NAME)
    Set recTbl = recSht.ListObjects(RECORDINGS_NAME)
    
    'Store the population info (or just return if none was provided)
    Dim lsRow As ListRow
    Dim rec As cRecording, tissueID As Integer
    Recordings.RemoveAll
    For Each lsRow In recTbl.ListRows
        Set rec = New cRecording
        rec.ID = lsRow.Range(1, recTbl.ListColumns("ID").index).Value
        rec.startTime = lsRow.Range(1, recTbl.ListColumns("StartStamp").index).Value
        rec.Duration = lsRow.Range(1, recTbl.ListColumns("Duration").index).Value
        tissueID = lsRow.Range(1, recTbl.ListColumns("Tissue ID").index).Value
        Set rec.Tissue = TISSUES(tissueID)
        TISSUES(tissueID).Recordings.Add rec
        Recordings.Add rec.ID, rec
    Next lsRow

End Sub

Private Sub defineTissues(ByRef summaryWb As Workbook)
    'Get the Tissues table
    Dim tissueSht As Worksheet, tissueTbl As ListObject
    Set tissueSht = summaryWb.Worksheets(TISSUES_NAME)
    Set tissueTbl = tissueSht.ListObjects(TISSUES_NAME)
    
    'Store the population info (or just return if none was provided)
    Dim lsRow As ListRow
    Dim tiss As cTissue
    TISSUES.RemoveAll
    For Each lsRow In tissueTbl.ListRows
        Set tiss = New cTissue
        tiss.ID = lsRow.Range(1, tissueTbl.ListColumns("ID").index).Value
        tiss.Name = lsRow.Range(1, tissueTbl.ListColumns("Name").index).Value
        tiss.DatePrepared = lsRow.Range(1, tissueTbl.ListColumns("Date Prepared").index).Value
        TISSUES.Add tiss.ID, tiss
    Next lsRow

End Sub

Private Sub definePopulations(ByRef summaryWb As Workbook)
    'Get the Populations and Recordings tables
    Dim popsSht As Worksheet, popsTbl As ListObject
    Set popsSht = summaryWb.Worksheets(POPS_NAME)
    Set popsTbl = popsSht.ListObjects(POPS_NAME)
    
    'Get burst types
    BURST_TYPES.RemoveAll
    BURST_TYPES.Add BurstUseType.WABs, "WAB"
    BURST_TYPES.Add BurstUseType.NonWABs, "NonWAB"
    
    'Store the population info (or just return if none was provided)
    Dim lsRow As ListRow, result As VbMsgBoxResult
    If popsTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No experimental populations have been defined.  Provide this info on the " & POPS_NAME & " sheet", vbOKOnly)
        Exit Sub
    End If
    Dim pop As cPopulation
    POPULATIONS.RemoveAll
    For Each lsRow In popsTbl.ListRows
        Set pop = New cPopulation
        pop.ID = lsRow.Range(1, popsTbl.ListColumns("Population ID").index).Value
        pop.Name = lsRow.Range(1, popsTbl.ListColumns("Name").index).Value
        pop.Abbreviation = lsRow.Range(1, popsTbl.ListColumns("Abbreviation").index).Value
        pop.IsControl = (lsRow.Range(1, popsTbl.ListColumns("Control?").index).Value <> "")
        pop.ForeColor = lsRow.Range(1, popsTbl.ListColumns("Population ID").index).Font.Color
        pop.BackColor = lsRow.Range(1, popsTbl.ListColumns("Population ID").index).Interior.Color
        POPULATIONS.Add pop.ID, pop
    Next lsRow
    
    'Identify the control population
    Dim numCtrlPops As Integer, p As Integer
    numCtrlPops = 0
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        If pop.IsControl Then
            Set CTRL_POP = pop
            numCtrlPops = numCtrlPops + 1
        End If
    Next p
    If numCtrlPops <> 1 Then
        result = MsgBox("You must identify one (and only one) experimental population as the control.", vbOKOnly)
        Exit Sub
    End If
    
End Sub

Private Sub definePopulationViews(ByRef popRecWb As Workbook)
    'Get the Populations and Recordings tables
    Dim recSht As Worksheet, recTbl As ListObject
    Set recSht = popRecWb.Worksheets(RECORDING_VIEWS_NAME)
    Set recTbl = recSht.ListObjects(RECORDING_VIEWS_NAME)

    'If no Recording info was provided on the Combine sheet, then just return
    Dim numRecs As Integer, result As VbMsgBoxResult
    numRecs = recTbl.ListRows.Count
    If recTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No recording-population associations have been specified.  Provide this info on the " & RECORDING_VIEWS_NAME & " sheet", vbOKOnly)
        Exit Sub
    End If
    
    'For each Population, associate each of its Tissues with a TissueView
    Dim tvs As New Dictionary, tv As cTissueView, p As Integer, pop As cPopulation
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        tvs.Add pop.ID, New Dictionary
    Next p
    
    'Build Views...
    Dim popID As Integer, recID As Integer, tID As Integer, rv As cRecordingView
    Dim txtPath As String, wbPath As String, lsRow As ListRow, bType As String, bt As Variant
    For Each lsRow In recTbl.ListRows
        'Create the TissueView object (if it doesn't already exist)
        'This includes defining its summary workbook paths
        popID = lsRow.Range(1, recTbl.ListColumns("Population ID").index).Value
        recID = lsRow.Range(1, recTbl.ListColumns("Recording ID").index).Value
        tID = Recordings(recID).Tissue.ID
        txtPath = lsRow.Range(1, recTbl.ListColumns("Text File").index).Value
        If tvs(popID).exists(tID) Then
            Set tv = tvs(popID)(tID)
        Else
            Set tv = New cTissueView
            Set tv.Tissue = TISSUES(tID)
            tvs(popID).Add tID, tv
            For Each bt In BURST_TYPES.Keys()
                bType = BURST_TYPES(bt)
                wbPath = Left(txtPath, InStrRev(txtPath, "\"))
                wbPath = wbPath & tID & "_" & Format(tv.Tissue.Name, "yyyy-mm-dd") & "_" & bType & ".xlsx"
                tv.WorkbookPaths.Add bt, wbPath
            Next bt
        End If
        
        'Create the RecordingView object
        Set rv = New cRecordingView
        Set rv.Recording = Recordings(recID)
        Set rv.TissueView = tv
        tv.RecordingViews.Add rv
        Set tv.Population = POPULATIONS(popID)
        POPULATIONS(popID).TissueViews.Add tv
        rv.TextPath = txtPath
    Next lsRow
End Sub

Private Sub defineUnits(ByRef summaryWb As Workbook)
    'Get all the provided invalid unit info
    Dim invalidsTbl As ListObject, invalidRng As Range
    Set invalidsTbl = summaryWb.Worksheets(INVALIDS_NAME).ListObjects(INVALIDS_NAME)
    Set invalidRng = invalidsTbl.DataBodyRange
    
    'For each provided unit...
    Dim lr As ListRow, unit As cUnit, popID As Integer, tissID As Integer, findTissue As Boolean, tissView As cTissueView, unitName As String, del As Boolean, exclude As Boolean
    If Not invalidRng Is Nothing Then
        For Each lr In invalidsTbl.ListRows
            Set unit = New cUnit
            
            'Fetch its data from the table
            tissID = lr.Range(1, invalidsTbl.ListColumns("Tissue ID").index).Value
            unitName = lr.Range(1, invalidsTbl.ListColumns("Unit").index).Value
            del = (lr.Range(1, invalidsTbl.ListColumns("Delete?").index).Value <> "")
            exclude = (lr.Range(1, invalidsTbl.ListColumns("Exclude?").index).Value <> "")
            
            'Get its associated TissueView
            findTissue = True
            If Not tissView Is Nothing Then _
                findTissue = (tissID <> tissView.Tissue.ID)
            If findTissue Then
                Dim p As Integer, pop As cPopulation, tv As cTissueView
                Set tissView = Nothing
                For p = 0 To POPULATIONS.Count - 1
                    Set pop = POPULATIONS.Items()(p)
                    For Each tv In pop.TissueViews
                        If tv.Tissue.ID = tissID Then
                            Set tissView = tv
                            Exit For
                        End If
                    Next tv
                    If Not tissView Is Nothing Then Exit For
                Next p
            End If
            
            'Wrap its data in a Unit object
            Set unit.TissueView = tissView
            tissView.BadUnits.Add unit
            unit.Name = unitName
            unit.ShouldDelete = del
            unit.ShouldExclude = exclude
        Next lr
    End If
End Sub

Public Sub GetConfigVars()
    'Prepare the property units array
    NUM_PROPERTIES = 16
    NUM_BKGRD_PROPERTIES = 10
    ReDim PROP_UNITS(1 To NUM_PROPERTIES)
    
    'Get the config parameters from the Params table
    Dim analyzeSht As Worksheet, configSht As Worksheet, configTbl As ListObject, params As Variant
    Set analyzeSht = Worksheets(ANALYZE_NAME)
    Set configSht = Worksheets(CONFIG_NAME)
    Set configTbl = configSht.ListObjects(CONFIG_NAME)
    params = configTbl.DataBodyRange.Resize(, 2).Value
    
    'Loop through each of its rows to cache parameter values
    Dim p As Integer
    Dim Name As String
    Dim val As Variant
    For p = 1 To UBound(params, 1)
        Name = params(p, 1)
        val = params(p, 2)
        Call storeParam(Name, val)
    Next p
    
    'Initialize parameters that depend on other parameters
    NUM_CHANNELS = MEA_ROWS * MEA_COLS
    MAX_POSSIBLE_UNITS = MAX_UNITS_PER_CHANNEL * NUM_CHANNELS  'Theoretically, no recording could possibly yield more sorted units than this
    NUM_BURST_PROPERTIES = NUM_PROPERTIES - NUM_BKGRD_PROPERTIES
    
    'Set property name strings
    'Try to just use alphanumeric characters w/o spaces since these will be Excel table headers later
    ReDim PROPERTIES(1 To NUM_PROPERTIES)
    PROPERTIES(1) = "NumSpikes"
    PROPERTIES(2) = "FiringRateOutsideAllBursts"
    PROPERTIES(3) = "FiringRateOutsideWABs"
    PROPERTIES(4) = "ISIOutsideAllBursts"
    PROPERTIES(5) = "ISIOutsideWABs"
    PROPERTIES(6) = "PercentSpikesOutsideAllBursts"
    PROPERTIES(7) = "PercentSpikesOutsideWABs"
    PROPERTIES(8) = "BurstFrequency"
    PROPERTIES(9) = "IBI"
    PROPERTIES(10) = "PercentBurstsInWaves"
    PROPERTIES(11) = "NumBursts"
    PROPERTIES(12) = "BurstDuration"
    PROPERTIES(13) = "BurstFiringRate"
    PROPERTIES(14) = "BurstISI"
    PROPERTIES(15) = "PercentBurstTimeAbove10Hz"
    PROPERTIES(16) = "SpikesPerBurst"
    
    'Get some other config flags set by the user
    Dim propMedIQR As Boolean, sttcMedIQR As Boolean
    DATA_PAIRED = (analyzeSht.Shapes("DataPairedChk").OLEFormat.Object.Value = 1)
    ASSOC_SAME_CHANNEL_UNITS = (analyzeSht.Shapes("SameChannelAssocChk").OLEFormat.Object.Value = 1)
    ASSOC_MULTIPLE_UNITS = (analyzeSht.Shapes("MultipleUnitsAssocChk").OLEFormat.Object.Value = 1)
    DELETE_BAD_SPIKES = (analyzeSht.Shapes("DeleteBadSpikesChk").OLEFormat.Object.Value = 1)
    DELETE_BAD_BURSTS = (analyzeSht.Shapes("DeleteBadBurstsChk").OLEFormat.Object.Value = 1)
    EXCLUDE_BURST_DUR_UNITS = (analyzeSht.Shapes("ExcludeBurstDurChk").OLEFormat.Object.Value = 1)
    
    propMedIQR = (analyzeSht.Shapes("PropMedIQRChk").OLEFormat.Object.Value = 1)
    sttcMedIQR = (analyzeSht.Shapes("SttcMedIQRChk").OLEFormat.Object.Value = 1)
    REPORT_PROPS_TYPE = IIf(propMedIQR, ReportStatsType.MedianIQR, ReportStatsType.MeanSEM)
    REPORT_STTC_TYPE = IIf(sttcMedIQR, ReportStatsType.MedianIQR, ReportStatsType.MeanSEM)
End Sub

Private Sub storeParam(ByVal Name As String, ByVal Value As Variant)
    If Name = "MEA Rows" Then
        MEA_ROWS = CInt(Value)
    ElseIf Name = "MEA Columns" Then
        MEA_COLS = CInt(Value)
    ElseIf Name = "Min Burst Duration" Then
        MIN_DURATION = CDbl(Value)
    ElseIf Name = "Max Burst Duration" Then
        MAX_DURATION = CDbl(Value)
    ElseIf Name = "Correlation dT" Then
        CORRELATION_DT = CDbl(Value)
    ElseIf Name = "Min Correlated Units" Then
        MIN_ASSOC_UNITS = CInt(Value)
    ElseIf Name = "Min Correlated Bins" Then
        MIN_BINS = CInt(Value)
    ElseIf Name = "Num Bins" Then
        NUM_BINS = CInt(Value)
        
    ElseIf Name = "NumSpikes Units" Then
        PROP_UNITS(1) = CStr(Value)
    ElseIf Name = "FiringRateOutsideAllBursts Units" Then
        PROP_UNITS(2) = CStr(Value)
    ElseIf Name = "FiringRateOutsideWABs Units" Then
        PROP_UNITS(3) = CStr(Value)
    ElseIf Name = "ISIOutsideAllBursts Units" Then
        PROP_UNITS(4) = CStr(Value)
    ElseIf Name = "ISIOutsideWABs Units" Then
        PROP_UNITS(5) = CStr(Value)
    ElseIf Name = "PercentSpikesOutsideAllBursts Units" Then
        PROP_UNITS(6) = CStr(Value)
    ElseIf Name = "PercentSpikesOutsideWABs Units" Then
        PROP_UNITS(7) = CStr(Value)
    ElseIf Name = "BurstFrequency Units" Then
        PROP_UNITS(8) = CStr(Value)
    ElseIf Name = "IBI Units" Then
        PROP_UNITS(9) = CStr(Value)
    ElseIf Name = "PercentBurstsInWaves Units" Then
        PROP_UNITS(10) = CStr(Value)
    ElseIf Name = "NumBursts Units" Then
        PROP_UNITS(11) = CStr(Value)
    ElseIf Name = "BurstDuration Units" Then
        PROP_UNITS(12) = CStr(Value)
    ElseIf Name = "BurstFiringRate Units" Then
        PROP_UNITS(13) = CStr(Value)
    ElseIf Name = "BurstISI Units" Then
        PROP_UNITS(14) = CStr(Value)
    ElseIf Name = "PercentBurstTimeAbove10Hz Units" Then
        PROP_UNITS(15) = CStr(Value)
    ElseIf Name = "SpikesPerBurst Units" Then
        PROP_UNITS(16) = CStr(Value)
    End If
    
End Sub

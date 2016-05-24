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
Public Const INVALIDS_NAME = "Invalid_Units"
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
Public MARK_BURST_DUR_UNITS As Boolean
Public DATA_PAIRED As Boolean
Public REPORT_PROPS_TYPE As ReportStatsType
Public REPORT_STTC_TYPE As ReportStatsType

'Arrays/collections and associated values
Public NUM_PROPERTIES As Integer
Public NUM_BKGRD_PROPERTIES As Integer
Public NUM_BURST_PROPERTIES As Integer
Public PROPERTIES() As String
Public PROP_UNITS() As String
Public BURST_TYPES As Variant
Public CTRL_POP As Population
Public POPULATIONS As New Dictionary
Public TISSUES As New Dictionary
Public Recordings As New Dictionary
Public INVALIDS As Variant
'OTHER VALUES
Public Const MAX_EXCEL_ROWS = 1048576

Public Function DefineObjects() As Boolean
    Dim Success As Boolean
    Success = False
    
    'Open the Data Summary workbook
    Dim summaryFile As File, result As VbMsgBoxResult
    Set summaryFile = PickWorkbook("Select the Data Summary workbook")
    If summaryFile Is Nothing Then
        result = MsgBox("No workbook selected.", vbOKOnly, "Routine complete")
        GoTo ExitFunc
    End If
    
    'Open the Population-definition workbook
    Dim popFile As File
    Set popFile = PickWorkbook("Select the workbook that defines your experimental populations")
    If popFile Is Nothing Then
        result = MsgBox("No workbook selected.", vbOKOnly, "Routine complete")
        GoTo ExitFunc
    End If
    
    'Get Tissue/Recording and experimental Population info
    Application.DisplayAlerts = False
    Workbooks.Open (summaryFile.Path)
    Call DefineRecordings
    Workbooks(summaryFile.Name).Close
    
    Workbooks.Open (popFile.Path)
    Call DefinePopulations
    Call getInvalidUnits
    Workbooks(popFile.Name).Close
    Application.DisplayAlerts = True
    
    Success = True
    
ExitFunc:
    DefineObjects = Success
End Function

Private Sub DefineTissues()
    'Get the Tissues table
    Dim tissueSht As Worksheet, tissueTbl As ListObject
    Set tissueSht = Worksheets(TISSUES_NAME)
    Set tissueTbl = tissueSht.ListObjects(TISSUES_NAME)
    
    'Store the population info (or just return if none was provided)
    Dim lsRow As ListRow
    Dim tiss As Tissue
    TISSUES.RemoveAll
    For Each lsRow In tissueTbl.ListRows
        Set tiss = New Tissue
        tiss.ID = lsRow.Range(1, tissueTbl.ListColumns("ID").index).value
        tiss.DatePrepared = lsRow.Range(1, tissueTbl.ListColumns("Date Prepared").index).value
        TISSUES.Add tiss.ID, tiss
    Next lsRow

End Sub

Private Sub DefineRecordings()
    'Make sure parent objects are defined first
    Call DefineTissues

    'Get the Recordings table
    Dim recSht As Worksheet, recTbl As ListObject
    Set recSht = Worksheets(RECORDINGS_NAME)
    Set recTbl = recSht.ListObjects(RECORDINGS_NAME)
    
    'Store the population info (or just return if none was provided)
    Dim lsRow As ListRow
    Dim rec As Recording, tissueID As Integer
    Recordings.RemoveAll
    For Each lsRow In recTbl.ListRows
        Set rec = New Recording
        rec.ID = lsRow.Range(1, recTbl.ListColumns("ID").index).value
        rec.StartTime = lsRow.Range(1, recTbl.ListColumns("StartStamp").index).value
        rec.Duration = lsRow.Range(1, recTbl.ListColumns("Duration").index).value
        tissueID = lsRow.Range(1, recTbl.ListColumns("Tissue ID").index).value
        Set rec.Tissue = TISSUES(tissueID)
        TISSUES(tissueID).Recordings.Add rec
        Recordings.Add rec.ID, rec
    Next lsRow

End Sub

Private Sub DefinePopulations()
    'Get the Populations and Recordings tables
    Dim popsSht As Worksheet, recSht As Worksheet, popsTbl As ListObject, recTbl As ListObject
    Set popsSht = Worksheets(POPS_NAME)
    Set popsTbl = popsSht.ListObjects(POPS_NAME)
    Set recSht = Worksheets(RECORDING_VIEWS_NAME)
    Set recTbl = recSht.ListObjects(RECORDING_VIEWS_NAME)
    
    'Get burst types
    Dim numBurstTypes As Integer, t As Integer, bType As String
    ReDim BURST_TYPES(1 To 1, 1 To 2)
    BURST_TYPES(1, 1) = "WAB"
    BURST_TYPES(1, 2) = "NonWAB"
    numBurstTypes = UBound(BURST_TYPES, 2)
    
    'Store the population info (or just return if none was provided)
    Dim lsRow As ListRow, result As VbMsgBoxResult
    If popsTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No experimental populations have been defined.  Provide this info on the " & POPS_NAME & " sheet", vbOKOnly)
        Exit Sub
    End If
    Dim pop As Population
    POPULATIONS.RemoveAll
    For Each lsRow In popsTbl.ListRows
        Set pop = New Population
        pop.ID = lsRow.Range(1, popsTbl.ListColumns("Population ID").index).value
        pop.Name = lsRow.Range(1, popsTbl.ListColumns("Name").index).value
        pop.Abbreviation = lsRow.Range(1, popsTbl.ListColumns("Abbreviation").index).value
        pop.IsControl = (lsRow.Range(1, popsTbl.ListColumns("Control?").index).value <> "")
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

    'If no Recording info was provided on the Combine sheet, then just return
    Dim numRecs As Integer
    numRecs = recTbl.ListRows.Count
    If recTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No recording-population associations have been specified.  Provide this info on the " & RECORDING_VIEWS_NAME & " sheet", vbOKOnly)
        Exit Sub
    End If
    
    'For each Population, associate each of its Tissues with a TissueView
    Dim tvs As New Dictionary, tv As TissueView
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        tvs.Add pop.ID, New Dictionary
    Next p
    
    'Build Views...
    Dim popID As Integer, recID As Integer, tID As Integer, rv As RecordingView
    Dim txtPath As String, wbPath As String
    For Each lsRow In recTbl.ListRows
        'Create the TissueView object (if it doesn't already exist)
        'This includes defining its summary workbook paths
        popID = lsRow.Range(1, recTbl.ListColumns("Population ID").index).value
        recID = lsRow.Range(1, recTbl.ListColumns("Recording ID").index).value
        tID = Recordings(recID).Tissue.ID
        txtPath = lsRow.Range(1, recTbl.ListColumns("Text File").index).value
        If tvs(popID).Exists(tID) Then
            Set tv = tvs(popID)(tID)
        Else
            Set tv = New TissueView
            Set tv.Tissue = TISSUES(tID)
            tvs(popID).Add tID, tv
            For t = 1 To UBound(BURST_TYPES, 2)
                bType = BURST_TYPES(1, t)
                wbPath = Left(txtPath, InStrRev(txtPath, "\"))
                wbPath = wbPath & lsRow.index & "_" & Format(tv.Tissue.DatePrepared, "yyyy-mm-dd") & "_" & bType & ".xlsx"
                tv.WorkbookPaths.Add bType, wbPath
            Next t
        End If
        
        'Create the RecordingView object
        Set rv = New RecordingView
        Set rv.Recording = Recordings(recID)
        Set rv.TissueView = tv
        tv.RecordingViews.Add rv
        Set tv.Population = POPULATIONS(popID)
        POPULATIONS(popID).TissueViews.Add tv
        rv.TextPath = txtPath
    Next lsRow
End Sub

Private Sub getInvalidUnits()
    'Get all the provided invalid unit info
    Dim invalidsTbl As ListObject, invalidRng As Range
    Set invalidsTbl = Worksheets(INVALIDS_NAME).ListObjects(INVALIDS_NAME)
    Set invalidRng = invalidsTbl.DataBodyRange
    If Not invalidRng Is Nothing Then
        INVALIDS = invalidRng.value
    Else
        ReDim INVALIDS(1 To 1, 1 To 1)
        INVALIDS(1, 1) = -1
    End If
End Sub

Public Sub GetConfigVars()
    'Prepare the property units array
    NUM_PROPERTIES = 11
    NUM_BKGRD_PROPERTIES = 6
    ReDim PROP_UNITS(1 To NUM_PROPERTIES)
    
    'Get the config parameters from the Params table
    Dim analyzeSht As Worksheet, configSht As Worksheet, configTbl As ListObject, params As Variant
    Set analyzeSht = Worksheets(ANALYZE_NAME)
    Set configSht = Worksheets(CONFIG_NAME)
    Set configTbl = configSht.ListObjects(CONFIG_NAME)
    params = configTbl.DataBodyRange.Resize(, 2).value
    
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
    PROPERTIES(1) = "BkgrdFiringRate"
    PROPERTIES(2) = "BkgrdISI"
    PROPERTIES(3) = "PercentSpikesInBursts"
    PROPERTIES(4) = "BurstFrequency"
    PROPERTIES(5) = "IBI"
    PROPERTIES(6) = "PercentBurstsInWaves"
    PROPERTIES(7) = "BurstDuration"
    PROPERTIES(8) = "BurstFiringRate"
    PROPERTIES(9) = "BurstISI"
    PROPERTIES(10) = "PercentBurstTimeAbove10Hz"
    PROPERTIES(11) = "SpikesPerBurst"
    
    'Get some other config flags set by the user
    Dim propMedIQR As Boolean, sttcMedIQR As Boolean
    DATA_PAIRED = (analyzeSht.Shapes("DataPairedChk").OLEFormat.Object.value = 1)
    ASSOC_SAME_CHANNEL_UNITS = (analyzeSht.Shapes("SameChannelAssocChk").OLEFormat.Object.value = 1)
    ASSOC_MULTIPLE_UNITS = (analyzeSht.Shapes("MultipleUnitsAssocChk").OLEFormat.Object.value = 1)
    MARK_BURST_DUR_UNITS = (analyzeSht.Shapes("ExcludeBurstDurChk").OLEFormat.Object.value = 1)
    
    propMedIQR = (analyzeSht.Shapes("PropMedIQRChk").OLEFormat.Object.value = 1)
    sttcMedIQR = (analyzeSht.Shapes("SttcMedIQRChk").OLEFormat.Object.value = 1)
    REPORT_PROPS_TYPE = IIf(propMedIQR, ReportStatsType.MedianIQR, ReportStatsType.MeanSEM)
    REPORT_STTC_TYPE = IIf(sttcMedIQR, ReportStatsType.MedianIQR, ReportStatsType.MeanSEM)
End Sub

Private Sub storeParam(ByVal Name As String, ByVal value As Variant)
    If Name = "MEA Rows" Then
        MEA_ROWS = CInt(value)
    ElseIf Name = "MEA Columns" Then
        MEA_COLS = CInt(value)
    ElseIf Name = "Min Burst Duration" Then
        MIN_DURATION = CDbl(value)
    ElseIf Name = "Max Burst Duration" Then
        MAX_DURATION = CDbl(value)
    ElseIf Name = "Correlation dT" Then
        CORRELATION_DT = CDbl(value)
    ElseIf Name = "Min Correlated Units" Then
        MIN_ASSOC_UNITS = CInt(value)
    ElseIf Name = "Min Correlated Bins" Then
        MIN_BINS = CInt(value)
    ElseIf Name = "Num Bins" Then
        NUM_BINS = CInt(value)
    ElseIf Name = "BkgrdFiringRate Units" Then
        PROP_UNITS(1) = CStr(value)
    ElseIf Name = "BkgrdISI Units" Then
        PROP_UNITS(2) = CStr(value)
    ElseIf Name = "PercentSpikesInBursts Units" Then
        PROP_UNITS(3) = CStr(value)
    ElseIf Name = "BurstFrequency Units" Then
        PROP_UNITS(4) = CStr(value)
    ElseIf Name = "IBI Units" Then
        PROP_UNITS(5) = CStr(value)
    ElseIf Name = "PercentBurstsInWaves Units" Then
        PROP_UNITS(6) = CStr(value)
    ElseIf Name = "BurstDuration Units" Then
        PROP_UNITS(7) = CStr(value)
    ElseIf Name = "BurstFiringRate Units" Then
        PROP_UNITS(8) = CStr(value)
    ElseIf Name = "BurstISI Units" Then
        PROP_UNITS(9) = CStr(value)
    ElseIf Name = "PercentBurstTimeAbove10Hz Units" Then
        PROP_UNITS(10) = CStr(value)
    ElseIf Name = "SpikesPerBurst Units" Then
        PROP_UNITS(11) = CStr(value)
    End If
End Sub

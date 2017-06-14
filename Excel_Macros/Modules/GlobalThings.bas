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
Public Const RECORDING_VIEWS_NAME = "Associated_Recordings"

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

'Global variables for this Module
Dim errorStrs As New Collection

Public Sub DefineObjects(ByRef errorStrCollection As Collection)

    Set errorStrs = errorStrCollection

    'Let the user choose the Data Summary workbook
    Dim summaryFile As File, summaryWb As Workbook
    Set summaryFile = PickWorkbook("Select the Data Summary workbook")
    If summaryFile Is Nothing Then
        errorStrs.Add "No Summary workbook selected."
        Exit Sub
    End If
    
    'Let the user choose the Population-Recording association workbook
    Dim popRecFile As File, popRecWb As Workbook
    Set popRecFile = PickWorkbook("Select the Population-Recording association workbook")
    If popRecFile Is Nothing Then
        errorStrs.Add "No PopRecordings workbook selected."
        Exit Sub
    End If
    
    'Open the workbooks that they chose
    Set summaryWb = Workbooks.Open(summaryFile.path)
    Set popRecWb = Workbooks.Open(popRecFile.path)
    
    'Get Tissue/Recording and experimental Population info
    'Wrap these objects in Views associated with the appropriate Population
    Call defineTissues(summaryWb)
    If errorStrs.Count > 0 Then GoTo Finally
    
    Call defineRecordings(summaryWb)
    If errorStrs.Count > 0 Then GoTo Finally
    
    Call definePopulations(popRecWb)
    If errorStrs.Count > 0 Then GoTo Finally
    
    Call associateViews(popRecWb)
    If errorStrs.Count > 0 Then GoTo Finally
    
    Call defineUnits(summaryWb)
    If errorStrs.Count > 0 Then GoTo Finally
    
    GoTo Finally
    
Finally:
    Application.DisplayAlerts = False
    summaryWb.Close
    popRecWb.Close
    Application.DisplayAlerts = True
    
End Sub

Private Sub defineTissues(ByRef summaryWb As Workbook)
    'Get the Tissues table
    Dim tissueSht As Worksheet, tissueTbl As ListObject
    Set tissueSht = summaryWb.Worksheets(TISSUES_NAME)
    Set tissueTbl = tissueSht.ListObjects(TISSUES_NAME)
    
    'Store the Tissue info (or just return if none was provided)
    Dim lsRow As ListRow
    If tissueTbl.DataBodyRange Is Nothing Then
        errorStrs.Add "No Tissues have been defined.  Provide this info on the " & TISSUES_NAME & " sheet of the Summary workbook."
        Exit Sub
    End If
    TISSUES.RemoveAll
    Dim tiss As cTissue
    For Each lsRow In tissueTbl.ListRows
        Set tiss = New cTissue
        tiss.Name = lsRow.Range(1, tissueTbl.ListColumns("Name").index).Value
        tiss.DatePrepared = lsRow.Range(1, tissueTbl.ListColumns("Date Prepared").index).Value
        TISSUES.Add tiss.Name, tiss
    Next lsRow

End Sub

Private Sub defineRecordings(ByRef summaryWb As Workbook)

    'Get the Recordings table
    Dim recSht As Worksheet, recTbl As ListObject
    Set recSht = summaryWb.Worksheets(RECORDINGS_NAME)
    Set recTbl = recSht.ListObjects(RECORDINGS_NAME)
    
    'Make sure Recording info was provided
    Dim lsRow As ListRow
    If recTbl.DataBodyRange Is Nothing Then
        errorStrs.Add "No Recordings have been defined.  Provide this info on the " & RECORDINGS_NAME & " sheet of the Summary workbook."
        Exit Sub
    End If
    
    'Store Recording info
    'If a Recording doesn't have a corresponding Tissue then return an error message
    Recordings.RemoveAll
    Dim rec As cRecording, tissName As String
    For Each lsRow In recTbl.ListRows
        Set rec = New cRecording
        rec.ID = lsRow.Range(1, recTbl.ListColumns("ID").index).Value
        rec.startTime = lsRow.Range(1, recTbl.ListColumns("StartStamp").index).Value
        rec.Duration = lsRow.Range(1, recTbl.ListColumns("Duration").index).Value
        tissName = lsRow.Range(1, recTbl.ListColumns("Tissue Name").index).Value
        If Not TISSUES.Exists(tissName) Then
            errorStrs.Add "Could not find a Tissue named " & tissName & " for Recording " & rec.ID & "."
            Exit Sub
        End If
        Set rec.Tissue = TISSUES(tissName)
        TISSUES(tissName).Recordings.Add rec
        Recordings.Add rec.ID, rec
    Next lsRow

End Sub

Private Sub definePopulations(ByRef popRecWb As Workbook)
    'Get the Populations and Recordings tables
    Dim popsSht As Worksheet, popsTbl As ListObject
    Set popsSht = popRecWb.Worksheets(POPS_NAME)
    Set popsTbl = popsSht.ListObjects(POPS_NAME)
    
    'Get burst types
    BURST_TYPES.RemoveAll
    BURST_TYPES.Add BurstUseType.WABs, "WAB"
    BURST_TYPES.Add BurstUseType.NonWABs, "NonWAB"
    
    'Store the Population info (or just return if none was provided)
    Dim lsRow As ListRow
    If popsTbl.DataBodyRange Is Nothing Then
        errorStrs.Add "No experimental Populations have been defined.  Provide this info on the " & POPS_NAME & " sheet of the PopRecordings workbook."
        Exit Sub
    End If
    Dim pop As cPopulation
    POPULATIONS.RemoveAll
    For Each lsRow In popsTbl.ListRows
        Set pop = New cPopulation
        pop.Name = lsRow.Range(1, popsTbl.ListColumns("Name").index).Value
        pop.Abbreviation = lsRow.Range(1, popsTbl.ListColumns("Abbreviation").index).Value
        pop.IsControl = (lsRow.Range(1, popsTbl.ListColumns("Control?").index).Value <> "")
        pop.ForeColor = lsRow.Range(1, popsTbl.ListColumns("Name").index).Font.Color
        pop.BackColor = lsRow.Range(1, popsTbl.ListColumns("Name").index).Interior.Color
        POPULATIONS.Add pop.Name, pop
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
        errorStrs.Add "You must identify one (and only one) experimental Population as the control."
        Exit Sub
    End If
    
End Sub

Private Sub associateViews(ByRef popRecWb As Workbook)
    'Get the Populations and Recordings tables
    Dim recSht As Worksheet, recTbl As ListObject
    Set recSht = popRecWb.Worksheets(RECORDING_VIEWS_NAME)
    Set recTbl = recSht.ListObjects(RECORDING_VIEWS_NAME)

    'If no Recording info was provided on the Combine sheet, then just return
    Dim numRecs As Integer
    numRecs = recTbl.ListRows.Count
    If recTbl.DataBodyRange Is Nothing Then
        errorStrs.Add "No recording-population associations have been specified.  Provide this info on the " & RECORDING_VIEWS_NAME & " sheet of the PopRecordings workbook."
        Exit Sub
    End If
    
    'For each Population, associate each of its Tissues with a TissueView
    Dim tvs As New Dictionary, tv As cTissueView, p As Integer, pop As cPopulation
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.Items()(p)
        tvs.Add pop.Name, New Dictionary
    Next p
    
    'Build Views...
    Dim popName As String, recID As Integer, tName As String, rv As cRecordingView
    Dim txtPath As String, wbPath As String, lsRow As ListRow, bType As String, bt As Variant
    For Each lsRow In recTbl.ListRows
        
        'Create the RecordingView object
        recID = lsRow.Range(1, recTbl.ListColumns("Recording ID").index).Value
        txtPath = lsRow.Range(1, recTbl.ListColumns("Text File Path").index).Value
        Set rv = New cRecordingView
        Set rv.Recording = Recordings(recID)
        rv.TextPath = txtPath
        
        'Create the TissueView object (if it doesn't already exist)
        'This includes defining its summary workbook paths
        popName = lsRow.Range(1, recTbl.ListColumns("Associated Population Name").index).Value
        tName = Recordings(recID).Tissue.Name
        If tvs(popName).Exists(tName) Then
            Set tv = tvs(popName)(tName)
        Else
            Set tv = New cTissueView
            Set tv.Tissue = TISSUES(tName)
            tvs(popName).Add tName, tv
            For Each bt In BURST_TYPES.Keys()
                bType = BURST_TYPES(bt)
                wbPath = Left(txtPath, InStrRev(txtPath, "\"))
                wbPath = wbPath & tName & "_" & Format(tv.Tissue.Name, "yyyy-mm-dd") & "_" & bType & ".xlsx"
                tv.WorkbookPaths.Add bt, wbPath
            Next bt
        End If
        
        'Associate View objects
        Set rv.TissueView = tv
        tv.RecordingViews.Add rv
        Set tv.Population = POPULATIONS(popName)
        POPULATIONS(popName).TissueViews.Add tv
        
    Next lsRow
    
    'Remove Populations with no associated Tissues
    Dim temp As New Collection
    For p = 0 To POPULATIONS.Count - 1
        temp.Add POPULATIONS.Items()(p)
    Next p
    For p = 1 To temp.Count
        Set pop = temp.item(p)
        If pop.TissueViews.Count = 0 Then _
            POPULATIONS.Remove (pop.Name)
    Next p
    
    'Remove Tissues with no associated Recordings
    Set temp = New Collection
    Dim tiss As cTissue, t As Integer
    For t = 0 To TISSUES.Count - 1
        temp.Add TISSUES.Items()(t)
    Next t
    For t = 1 To temp.Count
        Set tiss = temp.item(t)
        If tiss.Recordings.Count = 0 Then _
            TISSUES.Remove (tiss.Name)
    Next t
    
End Sub

Private Sub defineUnits(ByRef summaryWb As Workbook)
    'Only continue if invalid unit info was provided
    Dim invalidsTbl As ListObject
    Set invalidsTbl = summaryWb.Worksheets(INVALIDS_NAME).ListObjects(INVALIDS_NAME)
    If invalidsTbl.DataBodyRange Is Nothing Then Exit Sub
    
    'For each provided unit...
    Dim lr As ListRow, findTissue As Boolean
    Dim tissName As String, oldTissName As String, unitName As String, del As Boolean, exclude As Boolean
    Dim unit As cUnit, tissView As cTissueView
    For Each lr In invalidsTbl.ListRows
        
        'Fetch its data from the table
        tissName = lr.Range(1, invalidsTbl.ListColumns("Tissue Name").index).Value
        unitName = lr.Range(1, invalidsTbl.ListColumns("Unit").index).Value
        del = (lr.Range(1, invalidsTbl.ListColumns("Delete?").index).Value <> "")
        exclude = (lr.Range(1, invalidsTbl.ListColumns("Exclude?").index).Value <> "")
        
        'Try to find its associated TissueView
        findTissue = True
        If tissView Is Nothing Then
            findTissue = (tissName <> oldTissName)
        Else
            findTissue = (tissName <> tissView.Tissue.Name)
        End If
        If findTissue Then
            Dim p As Integer, pop As cPopulation, tv As cTissueView
            Set tissView = Nothing
            For p = 0 To POPULATIONS.Count - 1
                Set pop = POPULATIONS.Items()(p)
                For Each tv In pop.TissueViews
                    If tv.Tissue.Name = tissName Then
                        Set tissView = tv
                        Exit For
                    End If
                Next tv
                If Not tissView Is Nothing Then Exit For
            Next p
        End If
        
        'If the associated Tissue was found, then wrap this Unit's data in a Unit object
        'Otherwise, all remaining Units of this (not found) Tissue will be skipped
        If tissView Is Nothing Then
            oldTissName = tissName
        Else
            Set unit = New cUnit
            Set unit.TissueView = tissView
            tissView.BadUnits.Add unit
            unit.Name = unitName
            unit.ShouldDelete = del
            unit.ShouldExclude = exclude
        End If
    Next lr
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

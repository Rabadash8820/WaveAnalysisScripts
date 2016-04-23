Attribute VB_Name = "GlobalThings"
Option Explicit
Option Private Module

'ABSTRACT DATA TYPES
Public Enum BurstUseType
    All
    WABs
    NonWABs
End Enum
Public Enum wbComboType
    PropertyWkbk
    SttcWkbk
End Enum

'CONFIG VARIABLES
Public BURSTS_TO_USE As BurstUseType
Public ASSOC_SAME_CHANNEL_UNITS As Boolean
Public MEA_ROWS, MEA_COLS As Integer
Public NUM_CHANNELS As Integer
Public Const GROUND_CHANNEL = 4     'adch_15
Public Const MAX_UNITS_PER_CHANNEL = 10
Public MAX_POSSIBLE_UNITS As Integer
Public CORRELATION_DT As Double              'seconds
Public NUM_BINS, MIN_BINS As Double
Public MIN_ASSOC_UNITS As Integer
Public MIN_DURATION, MAX_DURATION As Double
Public Const MEAN_FREQ_DIFF = 3, PEAK_FREQ_DIFF = 10     'I.e., 300% and 1000%

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

'STRINGS
Public Const CELL_STR = "Cell"
Public Const RECORDING_STR = "Recording_"
Public Const STTC_HEADER_STR = "Spike Time Tiling Coefficient Values for Every Cell Pair"
Public Const INTER_ELECTRODE_DIST_STR = "Inter-Electrode Distance"
Public Const STTC_STR = "STTC"
Public Const CHANNEL_PREFIX = "adch_"

'Arrays/collections and associated values
Public NUM_PROPERTIES As Integer
Public NUM_BKGRD_PROPERTIES As Integer
Public NUM_BURST_PROPERTIES As Integer
Public PROPERTIES() As String
Public PROP_UNITS() As String
Public POPULATIONS As New Dictionary
Public CTRL_POP As Population
Public BURST_TYPES As Variant

'OTHER VALUES
Public Const MAX_EXCEL_ROWS = 1048576

'GLOBAL VARIABLES FOR THIS MODULE
Dim configTbl As ListObject
Dim popsTbl As ListObject
Dim tissueTbl As ListObject

Public Function PickWorkbook(ByVal pickMsg As String) As String
    Dim wbName As Workbook
    
    'Create the file-selection dialog box
    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.Title = pickMsg
    dialog.AllowMultiSelect = False
    
    'If the user didn't select anything, then return an empty string
    If dialog.Show = False Then
        PickWorkbook = ""
        Exit Function
    End If
    
    'Make sure that the selected file was actually an Excel workbook
    Dim fileSystem As New FileSystemObject
    Dim wbFile As File
    Set wbFile = fileSystem.GetFile(dialog.SelectedItems(1))
    Dim correctType As Boolean
    correctType = wbFile.Type = "Microsoft Excel Worksheet" Or wbFile.Type = "Microsoft Excel Macro-Enabled Worksheet"
    If Not correctType Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Selected file is not an Excel workbook.", vbOKOnly, "Routine complete")
        Exit Function
    End If
    
    'If it was then return its name
    PickWorkbook = wbFile.name
End Function

Public Sub DefinePopulations()
    
    'Get the Populations and Tissues tables
    Dim popsSht As Worksheet, tissueSht As Worksheet
    Set popsSht = Worksheets(POPS_NAME)
    Set popsTbl = popsSht.ListObjects(POPS_NAME)
    Set tissueSht = Worksheets(TISSUES_NAME)
    
    'Get burst types
    Dim numBurstTypes As Integer, t As Integer, bType As String
    Set tissueTbl = tissueSht.ListObjects(TISSUES_NAME)
    BURST_TYPES = tissueTbl.HeaderRowRange(1, 3).Resize(1, tissueTbl.ListColumns.Count - 2).value
    numBurstTypes = UBound(BURST_TYPES, 2)
    For t = 1 To numBurstTypes
        bType = BURST_TYPES(1, t)
        bType = Left(bType, Len(bType) - Len(" Workbook"))
        BURST_TYPES(1, t) = bType
    Next t
    
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
        pop.ID = lsRow.Range(1, popsTbl.ListColumns("Population ID").Index).value
        pop.name = lsRow.Range(1, popsTbl.ListColumns("Name").Index).value
        pop.Abbreviation = lsRow.Range(1, popsTbl.ListColumns("Abbreviation").Index).value
        pop.IsControl = (lsRow.Range(1, popsTbl.ListColumns("Control?").Index).value <> "")
        pop.ForeColor = lsRow.Range(1, popsTbl.ListColumns("Population ID").Index).Font.Color
        pop.BackColor = lsRow.Range(1, popsTbl.ListColumns("Population ID").Index).Interior.Color
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

    'If no Tissue info was provided on the Combine sheet, then just return
    Dim numTissues As Integer
    numTissues = tissueTbl.ListRows.Count
    If tissueTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No tissues have been defined.  Provide this info on the " & TISSUES_NAME & " sheet", vbOKOnly)
        Exit Sub
    End If
    
    'Otherwise, create the Tissue objects
    Dim popID As Integer, wbPath As String, tiss As Tissue
    For Each lsRow In tissueTbl.ListRows
        Set tiss = New Tissue
        tiss.ID = lsRow.Range(1, tissueTbl.ListColumns("Tissue ID").Index).value
        popID = lsRow.Range(1, tissueTbl.ListColumns("Population ID").Index).value
        Set tiss.Population = POPULATIONS(popID)
        For t = 1 To numBurstTypes
            wbPath = lsRow.Range(1, tissueTbl.ListColumns(BURST_TYPES(1, t) & " Workbook").Index).value
            tiss.WorkbookPaths.Add BURST_TYPES(1, t), wbPath
        Next t
        POPULATIONS(popID).Tissues.Add tiss
    Next lsRow

End Sub

Public Sub GetConfigVars()
    'Prepare the property units array
    NUM_PROPERTIES = 11
    NUM_BKGRD_PROPERTIES = 6
    ReDim PROP_UNITS(1 To NUM_PROPERTIES)
    
    'Get the config parameters from the Params table
    Dim analyzeSht As Worksheet, configSht As Worksheet, params As Variant
    Set analyzeSht = Worksheets(ANALYZE_NAME)
    Set configSht = Worksheets(CONFIG_NAME)
    Set configTbl = configSht.ListObjects(CONFIG_NAME)
    params = configTbl.DataBodyRange.Resize(, 2).value
    
    'Loop through each of its rows to cache parameter values
    Dim p As Integer
    Dim name As String
    Dim val As Variant
    For p = 1 To UBound(params, 1)
        name = params(p, 1)
        val = params(p, 2)
        Call storeParam(name, val)
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
    
End Sub

Private Sub storeParam(ByVal name As String, ByVal value As Variant)
    If name = "MEA Rows" Then
        MEA_ROWS = CInt(value)
    ElseIf name = "MEA Columns" Then
        MEA_COLS = CInt(value)
    ElseIf name = "Min Burst Duration" Then
        MIN_DURATION = CDbl(value)
    ElseIf name = "Max Burst Duration" Then
        MAX_DURATION = CDbl(value)
    ElseIf name = "Correlation dT" Then
        CORRELATION_DT = CDbl(value)
    ElseIf name = "Min Correlated Units" Then
        MIN_ASSOC_UNITS = CInt(value)
    ElseIf name = "Min Correlated Bins" Then
        MIN_BINS = CInt(value)
    ElseIf name = "Num Bins" Then
        NUM_BINS = CInt(value)
    ElseIf name = "BkgrdFiringRate Units" Then
        PROP_UNITS(1) = CStr(value)
    ElseIf name = "BkgrdISI Units" Then
        PROP_UNITS(2) = CStr(value)
    ElseIf name = "PercentSpikesInBursts Units" Then
        PROP_UNITS(3) = CStr(value)
    ElseIf name = "BurstFrequency Units" Then
        PROP_UNITS(4) = CStr(value)
    ElseIf name = "IBI Units" Then
        PROP_UNITS(5) = CStr(value)
    ElseIf name = "PercentBurstsInWaves Units" Then
        PROP_UNITS(6) = CStr(value)
    ElseIf name = "BurstDuration Units" Then
        PROP_UNITS(7) = CStr(value)
    ElseIf name = "BurstFiringRate Units" Then
        PROP_UNITS(8) = CStr(value)
    ElseIf name = "BurstISI Units" Then
        PROP_UNITS(9) = CStr(value)
    ElseIf name = "PercentBurstTimeAbove10Hz Units" Then
        PROP_UNITS(10) = CStr(value)
    ElseIf name = "SpikesPerBurst Units" Then
        PROP_UNITS(11) = CStr(value)
    End If
End Sub

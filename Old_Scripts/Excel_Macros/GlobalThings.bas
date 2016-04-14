Attribute VB_Name = "GlobalThings"
Option Explicit
Option Private Module

Public Enum BurstUseType
    All
    WABs
    NonWABs
End Enum
Public Enum wbComboType
    PropertyWkbk
    SttcWkbk
End Enum
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

Public Const CHANNEL_PREFIX = "adch_"

Public Const CONFIG_SHEET_NAME = "Analyze"
Public Const COMBINE_SHEET_NAME = "Combine Results"
Public Const INVALIDS_SHEET_NAME = "Invalid Units"
Public Const POPULATIONS_SHEET_NAME = "Populations"
Public Const CONTENTS_SHEET_NAME = "Contents"
Public Const ALL_AVGS_SHEET_NAME = "All Avgs"
Public Const BURST_AVGS_SHEET_NAME = "Burst Avgs"
Public Const STTC_SHEET_NAME = "STTC"

Public Const CELL_STR = "Cell"
Public Const RECORDING_STR = "Recording_"
Public Const STTC_HEADER_STR = "Spike Time Tiling Coefficient Values for Every Cell Pair"
Public Const INTER_ELECTRODE_DIST_STR = "Inter-Electrode Distance"
Public Const STTC_STR = "STTC"
    
Public NUM_PROPERTIES As Integer
Public NUM_BKGRD_PROPERTIES As Integer
Public NUM_BURST_PROPERTIES As Integer
Public PROPERTIES() As String
Public PROP_UNITS() As String

Dim configSht As Worksheet

Public Function pickWorkbook(ByVal pickMsg As String) As String
    Dim wbName As Workbook
    
    'Create the file-selection dialog box
    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.title = pickMsg
    dialog.AllowMultiSelect = False
    
    'If the user didn't select anything, then return an empty string
    If dialog.Show = False Then
        pickWorkbook = ""
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
    pickWorkbook = wbFile.name
End Function
    
Public Sub initParams()
    Set configSht = Worksheets(CONFIG_SHEET_NAME)

    Call getPropertyNames
    Call getParams
    Call getBurstsToUse
End Sub

Private Sub getParams()
    ReDim PROP_UNITS(1 To NUM_PROPERTIES)

    'Get the config parameters from the Params table
    Dim params As Variant
    Dim paramsTbl As ListObject
    Set paramsTbl = configSht.ListObjects("ParamsTbl")
    params = paramsTbl.DataBodyRange.Resize(, 2).value
    
    'Loop through each of its rows to cache parameter values
    Dim p As Integer
    Dim name As String
    Dim val As Variant
    For p = 1 To UBound(params, 1)
        name = params(p, 1)
        val = params(p, 2)
        Call storeParam(name, val)
    Next p
    
    'Get config info from Form Controls on sheet
    Call getBurstsToUse
    ASSOC_SAME_CHANNEL_UNITS = (configSht.Shapes("SameChannelAssocChk").OLEFormat.Object.value = 1)
    
    'Initialize parameters that depend on other parameters
    NUM_CHANNELS = MEA_ROWS * MEA_COLS
    MAX_POSSIBLE_UNITS = MAX_UNITS_PER_CHANNEL * NUM_CHANNELS  'Theoretically, no recording could possibly yield more sorted units than this
    
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

Private Sub getPropertyNames()
    'Set property "constants"
    NUM_PROPERTIES = 11
    NUM_BKGRD_PROPERTIES = 6
    NUM_BURST_PROPERTIES = NUM_PROPERTIES - NUM_BKGRD_PROPERTIES
    
    'Set property name strings
    'Try to just use alphanumeric characters w/o spaces since these will be Excel table headers later
    ReDim PROPERTIES(1 To NUM_PROPERTIES) As String
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

Private Sub getBurstsToUse()
    Dim allChecked As Boolean, wabsChecked As Boolean
    allChecked = (configSht.Shapes("AllRadio").OLEFormat.Object.value = 1)
    wabsChecked = (configSht.Shapes("WabsRadio").OLEFormat.Object.value = 1)
    
    BURSTS_TO_USE = IIf(allChecked, BurstUseType.All, IIf(wabsChecked, BurstUseType.WABs, BurstUseType.NonWABs))
End Sub

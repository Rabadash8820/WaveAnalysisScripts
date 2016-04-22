Attribute VB_Name = "InvalidUnits"
Option Explicit

Const NUM_NONPOP_SHEETS = 3
Const BURST_DUR_COL = 3
Const MARK_STYLE = "Bad"
Const NORMAL_STYLE = "Normal"

Dim invalids As Variant
Dim markBurstDurUnits As Boolean, deleteAlso As Boolean, keepOpen As Boolean

'THE DATA VALIDATION WORKSHEET SHOULD LIST UNITS THAT WILL BE *MARKED FOR REMOVAL*

Public Sub markInvalidUnits()
    Call setupOptimizations
    
    Call DefinePopulations
    
    'Only continue if at least one invalidation operation was selected
    Dim result As VbMsgBoxResult
    Dim thisWb As Workbook
    Set thisWb = ActiveWorkbook
    Dim invalidSht As Worksheet
    Set invalidSht = thisWb.Worksheets(INVALIDS_SHT_NAME)
    
    'Get all the provided invalid unit info
    Dim invalidUnitsTbl As ListObject
    Set invalidSht = Worksheets(INVALIDS_SHT_NAME)
    Set invalidUnitsTbl = invalidSht.ListObjects(INVALIDS_TBL_NAME)
    If invalidUnitsTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No invalid units provided.", vbOKOnly)
        GoTo ExitSub
    Else
        invalids = invalidUnitsTbl.DataBodyRange.value
    End If
    
    'Get some other config flags set by the user
    markBurstDurUnits = (invalidSht.Shapes("MarkBurstDurChk").OLEFormat.Object.value = 1)
    deleteAlso = (invalidSht.Shapes("InvalidDeleteChk").OLEFormat.Object.value = 1)
    keepOpen = (invalidSht.Shapes("KeepOpenChk").OLEFormat.Object.value = 1)
    
    'Let the user pick the workbook in which to mark invalid units
    Dim wbName As String, sttcWbName As String
    wbName = PickWorkbook("Select the Summary workbook in which to mark invalid units")

    'Open this workbook and mark any invalid units therein
    If wbName = "" Then
        result = MsgBox("No workbook selected.", vbOKOnly)
        GoTo ExitSub
    Else
        Dim wb As Workbook, numMarkedUnits As Long
        numMarkedUnits = 0
        Set wb = Workbooks.Open(wbName)
        Call markDataOnWbType(wb, wbComboType.PropertyWkbk, numMarkedUnits)
        Set wb = Workbooks.Open(wbName)
        Call markDataOnWbType(wb, wbComboType.SttcWkbk, numMarkedUnits)
    End If
    
    'Warn user to remove zero property values also
    Dim markedStr As String, msg As String
    markedStr = IIf(deleteAlso, "deleted", "marked")
    Dim numInvalids As Integer
    numInvalids = UBound(invalids, 1)
    msg = numInvalids & " invalid units provided." & vbCr & _
          numMarkedUnits & " units actually " & markedStr & "." & vbCr & _
          "Time taken: " & Format(ProgramDuration(), "hh:mm:ss")
    result = MsgBox(msg, vbOKOnly)

ExitSub:
    Call tearDownOptimizations
End Sub

Private Sub markDataOnWbType(ByRef wkbk As Workbook, ByVal wbType As wbComboType, ByRef numMarkedUnits As Long)
    Call resetRanges
            
    Select Case wbType
        'If we're marking a Properties workbook...
        'These workbooks have the option to also mark units with zero burst duration
        Case wbComboType.PropertyWkbk
            Call markPropWbData(wkbk, numMarkedUnits)
            If markBurstDurUnits Then _
                Call markZeroBurstDurUnits(wkbk)
        
        'If we're marking an STTC workbook...
        Case wbComboType.SttcWkbk
            Call markSttcWbData(wkbk, numMarkedUnits)
        
    End Select
End Sub

Private Sub markPropWbData(ByRef wkbk As Workbook, ByRef numMarkedUnits As Long)
    'For each invalid unit...
    'Find the data sheets that match its population name,
    'Find the table rows on those sheets that match its retina and Unit IDs,
    'And mark that row with a noticeable style, or delete it if requested
    Dim numPops As Integer
    numPops = Worksheets.Count - NUM_NONPOP_SHEETS
    Dim i As Long, lr As Long, sh As Integer, sht As Worksheet
    Dim marked As Boolean
    Dim popName As String, retinaID As String, unitID As String
    Dim shtMatches As Boolean, rowMatches As Boolean
    Dim lsRows As ListRows, lsRng As Range
    Dim numInvalids As Integer
    numInvalids = UBound(invalids, 1)
    For i = 1 To numInvalids
        popName = invalids(i, 1)
        retinaID = invalids(i, 2)
        unitID = invalids(i, 3)
        marked = False
        For sh = 1 To numPops
            Set sht = Worksheets(NUM_NONPOP_SHEETS + sh)
            shtMatches = (InStr(1, sht.name, popName) <> 0)
            If shtMatches Then
                Set lsRows = sht.ListObjects(sht.name).ListRows
                For lr = 1 To lsRows.Count
                    Set lsRng = lsRows(lr).Range
                    rowMatches = lsRng.Cells(1, 1).value = retinaID And lsRng.Cells(1, 2).value = unitID
                    If rowMatches Then
                        Call markUnit(lsRows(lr))
                        marked = True
                        Exit For
                    End If
                Next lr
            End If
        Next sh
        If marked Then _
            numMarkedUnits = numMarkedUnits + 1
    Next i
End Sub

Private Sub markSttcWbData(ByRef wkbk As Workbook, ByRef numMarkedUnits As Long)
    'For each invalid unit...
    'Find the _STTC sheets that match its population name,
    'Find the table rows on those sheets that match its retina and Unit IDs,
    'And mark those rows with a noticeable style, or delete it (if requested)
    Dim numPops As Integer
    numPops = Worksheets.Count - NUM_NONPOP_SHEETS
    Dim i As Long, lr As Long, currRow As Long, sh As Integer, sht As Worksheet
    Dim marked As Boolean
    Dim popName As String, retinaID As String, unitID As String
    Dim shtMatches As Boolean, rowMatches As Boolean
    Dim lsRows As ListRows, lsRng As Range
    Dim numInvalids As Integer
    numInvalids = UBound(invalids, 1)
    For i = 1 To numInvalids
        popName = invalids(i, 1)
        retinaID = invalids(i, 2)
        unitID = invalids(i, 3)
        marked = False
        For sh = 1 To numPops
            Set sht = Worksheets(NUM_NONPOP_SHEETS + sh)
            shtMatches = (InStr(1, sht.name, popName) <> 0)
            If shtMatches Then
                currRow = 1
                Set lsRows = sht.ListObjects(sht.name).ListRows
                For lr = 1 To lsRows.Count
                    Set lsRng = lsRows(currRow).Range
                    rowMatches = (lsRng.Cells(1, 1).value = retinaID And _
                                 (lsRng.Cells(1, 2).value = unitID Or lsRng.Cells(1, 3).value = unitID))
                    If rowMatches Then
                        Call markUnit(lsRows(currRow))
                        marked = True
                        If Not deleteAlso Then _
                            currRow = currRow + 1
                    Else
                        currRow = currRow + 1
                    End If
                Next lr
            End If
        Next sh
        If marked Then _
            numMarkedUnits = numMarkedUnits + 1
    Next i

End Sub

Private Sub markZeroBurstDurUnits(propsWb As Workbook)
    'Sheets with the following keywords in their names will have units with burst durations of 0 marked
    Dim keywords(1 To 2) As String, keyword As Variant
    keywords(1) = "_WABs"
    keywords(2) = "_NonWABs"
    
    'Mark units with burst durations of 0 on all applicable sheets
    Dim sht As Worksheet, tbl As ListObject, lsRows As ListRows
    Dim lr As Long, currRow As Long
    Dim marked As Boolean
    Dim shtMatches As Boolean, rowMatches As Boolean
    For Each sht In propsWb.Worksheets
        For Each keyword In keywords
            shtMatches = (InStr(1, sht.name, keyword) > 0)
            If shtMatches Then
                currRow = 1
                Set lsRows = sht.ListObjects(sht.name).ListRows
                For lr = 1 To lsRows.Count
                    rowMatches = (lsRows(currRow).Range(1, BURST_DUR_COL).value = 0)
                    If rowMatches Then
                        Call markUnit(lsRows(currRow))
                        marked = True
                        If Not deleteAlso Then _
                            currRow = currRow + 1
                    Else
                        currRow = currRow + 1
                    End If
                Next lr
            End If
        Next keyword
    Next sht
End Sub

Private Sub resetRanges()
    'Clear any old mark styles
    Dim numPops As Integer
    numPops = Worksheets.Count - NUM_NONPOP_SHEETS
    Dim sh As Integer, sht As Worksheet
    Dim lsRng As Range
    For sh = 1 To numPops
        Set sht = Worksheets(NUM_NONPOP_SHEETS + sh)
        Set lsRng = sht.ListObjects(sht.name).DataBodyRange
        Call resetRange(lsRng)
    Next sh
End Sub

Private Sub resetRange(ByRef rng As Range)
    rng.Style = NORMAL_STYLE
    rng.VerticalAlignment = xlCenter
    rng.HorizontalAlignment = xlCenter
End Sub

Private Sub markUnit(ByRef lsRow As ListRow)
'   Debug.Print ("Marking <" & popName & ", " & retinaID & ", " & unitID & "> on sheet " & sht.Name)

    If deleteAlso Then
        lsRow.Delete
    Else
        lsRow.Range.Style = MARK_STYLE
    End If
End Sub

Attribute VB_Name = "InvalidUnits"
Option Explicit

Const NUM_NONPOP_SHEETS = 4
Const BURST_DUR_COL = 3
Const MARK_STYLE = "Bad"
Const NORMAL_STYLE = "Normal"

Dim invalids As Variant
Dim markBurstDurUnits As Boolean, deleteAlso As Boolean, keepOpen As Boolean

Public Sub markInvalidUnits()
    Call setupOptimizations
    
    Call DefinePopulations
    
    'Only continue if at least one invalidation operation was selected
    Dim result As VbMsgBoxResult
    Dim thisWb As Workbook
    Set thisWb = ActiveWorkbook
    Dim invalidSht As Worksheet
    Set invalidSht = thisWb.Worksheets(INVALIDS_NAME)
    
    'Get all the provided invalid unit info
    Dim invalidUnitsTbl As ListObject
    Set invalidSht = Worksheets(INVALIDS_NAME)
    Set invalidUnitsTbl = invalidSht.ListObjects(INVALIDS_NAME)
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
    wbName = PickWorkbook("Select the Results workbook in which to mark invalid units")
    If wbName = "" Then
        result = MsgBox("No workbook selected.", vbOKOnly)
        GoTo ExitSub
    End If
    
    'Open this workbook and mark any invalid units therein
    Dim wb As Workbook, numMarkedUnits As Long
    numMarkedUnits = 0
    Set wb = Workbooks.Open(wbName)
    Call resetRanges
    Call markPropWbData(wb, numMarkedUnits)
    If markBurstDurUnits Then _
        Call markZeroBurstDurUnits(wb)
    Call markSttcWbData(wb, numMarkedUnits)

    'Save/close the workbook if the user doesn't want to keep it open
    If Not keepOpen Then _
        Call wb.Close(True)
    
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

Private Sub markPropWbData(ByRef wkbk As Workbook, ByRef numMarkedUnits As Long)
    'For each invalid unit...
    'Find the data sheets that match its population name,
    'Find the table rows on those sheets that match its tissue and Unit IDs,
    'And mark that row with a noticeable style, or delete them if requested
    Dim numPops As Integer
    numPops = Worksheets.Count - NUM_NONPOP_SHEETS
    Dim i As Long, lr As Long, sh As Integer, sht As Worksheet
    Dim marked As Boolean
    Dim pop As Population, tissueID As String, unitID As String
    Dim shtMatches As Boolean, rowMatches As Boolean
    Dim lsRows As ListRows, lsRng As Range
    Dim numInvalids As Integer
    numInvalids = UBound(invalids, 1)
    For i = 1 To numInvalids
        marked = False
        Set pop = POPULATIONS(invalids(i, 1))
        tissueID = invalids(i, 2)
        unitID = invalids(i, 3)
        
        For sh = 1 To numPops
            Set sht = Worksheets(NUM_NONPOP_SHEETS + sh)
            shtMatches = (InStr(1, sht.name, pop.name) <> 0) And (InStr(1, sht.name, "STTC") = 0)
            If shtMatches Then
                Set lsRows = sht.ListObjects(sht.name).ListRows
                For lr = 1 To lsRows.Count
                    Set lsRng = lsRows(lr).Range
                    rowMatches = lsRng.Cells(1, 1).value = tissueID And lsRng.Cells(1, 2).value = unitID
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
    'Find the STTC sheet that matches its population name,
    'Find the table rows on those sheets that match its tissue and Unit IDs,
    'And mark those rows with a noticeable style, or delete them if requested
    Dim numPops As Integer
    numPops = Worksheets.Count - NUM_NONPOP_SHEETS
    Dim i As Long, lr As Long, currRow As Long, sht As Worksheet
    Dim marked As Boolean
    Dim pop As Population, tissueID As String, unitID As String
    Dim rowMatches As Boolean, lsRows As ListRows, lsRng As Range
    Dim numInvalids As Integer
    numInvalids = UBound(invalids, 1)
    For i = 1 To numInvalids
        marked = False
        Set pop = POPULATIONS(invalids(i, 1))
        tissueID = invalids(i, 2)
        unitID = invalids(i, 3)
        
        Set sht = Worksheets(pop.name & "_STTC")
        currRow = 1
        Set lsRows = sht.ListObjects(sht.name).ListRows
        For lr = 1 To lsRows.Count
            Set lsRng = lsRows(currRow).Range
            rowMatches = (lsRng.Cells(1, 1).value = tissueID And _
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
    Dim numPops As Integer, sh As Integer, sht As Worksheet, lsRng As Range
    numPops = Worksheets.Count - NUM_NONPOP_SHEETS
    For sh = 1 To numPops
        Set sht = Worksheets(NUM_NONPOP_SHEETS + sh)
        Set lsRng = sht.ListObjects(sht.name).DataBodyRange
        With lsRng
            .Style = NORMAL_STYLE
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
    Next sh
End Sub

Private Sub markUnit(ByRef lsRow As ListRow)
    If deleteAlso Then
        lsRow.Delete
    Else
        lsRow.Range.Style = MARK_STYLE
    End If
End Sub

Attribute VB_Name = "InvalidUnits"
Option Explicit
Option Private Module

Const DELETE_PREFIX = "XxxX_"
Const NUM_NONPOP_SHEETS = 4
Const BURST_DUR_COL = 3

Public Sub invalidateUnits(ByRef sht As Worksheet, ByVal tiss As cTissue, ByRef unitNames As Variant)
    'For each invalid unit on this Tissue, mark the columns that contains its data
    'These include the spike timestamp and burst start/end timestamp columns
    Dim cornerCell As Range
    Set cornerCell = Cells(1, 1)
    Dim numInvalids As Long, numUnits As Integer, iu As Long, u As Integer
    Dim ID As Integer, unitName As String, burstOffset As Integer
    numInvalids = UBound(INVALIDS)
    numUnits = UBound(unitNames)
    For iu = 1 To numInvalids
        ID = INVALIDS(iu, 2)
        If ID = tiss.ID Then
            unitName = INVALIDS(iu, 3)
            For u = 1 To numUnits
                If unitNames(u, 1) = unitName Then
                    burstOffset = numUnits + 2 * (u - 1)
                    cornerCell.offset(0, u - 1).Value = DELETE_PREFIX & cornerCell.offset(0, u - 1).Value
                    cornerCell.offset(0, burstOffset).Value = DELETE_PREFIX & cornerCell.offset(0, burstOffset).Value
                    cornerCell.offset(0, burstOffset + 1).Value = DELETE_PREFIX & cornerCell.offset(0, burstOffset + 1).Value
                End If
            Next u
        End If
    Next iu
    
    'Go back through the columns and delete the ones that were marked
    Dim doDelete As Boolean, headerRng As Range
    For u = 3 * numUnits - 1 To 0 Step -1
        Set headerRng = cornerCell.offset(0, u)
        doDelete = (Left(headerRng.Value, Len(DELETE_PREFIX)) = DELETE_PREFIX)
        If doDelete Then _
            Columns(headerRng.Column).Delete
    Next u
End Sub

Public Sub deleteZeroBurstDurUnits(ByRef wb As Workbook)
    'Sheets with the following keywords in their names will have units with burst durations of 0 marked
    Dim keywords(1 To 2) As String, keyword As Variant
    keywords(1) = "_WABs"
    keywords(2) = "_NonWABs"
    
    'Mark units with burst durations of 0 on all applicable sheets
    Dim sht As Worksheet, tbl As ListObject, lsRows As ListRows
    Dim lr As Long, currRow As Long
    Dim shtMatches As Boolean, rowMatches As Boolean
    For Each sht In wb.Worksheets
        For Each keyword In keywords
            shtMatches = (InStr(1, sht.Name, keyword) > 0)
            If shtMatches Then
                currRow = 1
                Set lsRows = sht.ListObjects(sht.Name).ListRows
                For lr = 1 To lsRows.Count
                    rowMatches = (lsRows(currRow).Range(1, BURST_DUR_COL).Value = 0)
                    If rowMatches Then
                        lsRows(currRow).Delete
                    Else
                        currRow = currRow + 1
                    End If
                Next lr
            End If
        Next keyword
    Next sht
End Sub

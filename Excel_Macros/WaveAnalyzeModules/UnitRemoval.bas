Attribute VB_Name = "UnitRemoval"
Option Explicit
Option Private Module

Const DELETE_PREFIX = "XxxX_"
Const NUM_NONPOP_SHEETS = 4
Const BURST_DUR_COL = 3

Public Sub DeleteUnits(ByRef sht As Worksheet, ByVal tiss As cTissue, ByRef unitNames As Variant)
    'For each invalid unit on this Tissue, mark the columns that contains its data
    'These include the spike timestamp and burst start/end timestamp columns
    Dim cornerCell As Range
    Set cornerCell = sht.Cells(1, 1)
    Dim numUnits As Integer, u As Integer, unit As cUnit, burstOffset As Integer
    numUnits = UBound(unitNames)
    For Each unit In DELETE_UNITS
        If unit.Tissue.ID = tiss.ID Then
            For u = 1 To numUnits
                If unitNames(u, 1) = unit.Name Then
                    burstOffset = numUnits + 2 * (u - 1)
                    cornerCell.offset(0, u - 1).Value = DELETE_PREFIX & cornerCell.offset(0, u - 1).Value
                    cornerCell.offset(0, burstOffset).Value = DELETE_PREFIX & cornerCell.offset(0, burstOffset).Value
                    cornerCell.offset(0, burstOffset + 1).Value = DELETE_PREFIX & cornerCell.offset(0, burstOffset + 1).Value
                    Exit For
                End If
            Next u
        End If
    Next unit
    
    'Go back through the columns and delete the ones that were marked
    Dim doDelete As Boolean, headerRng As Range
    For u = 3 * numUnits - 1 To 0 Step -1
        Set headerRng = cornerCell.offset(0, u)
        doDelete = (Left(headerRng.Value, Len(DELETE_PREFIX)) = DELETE_PREFIX)
        If doDelete Then _
            sht.Columns(headerRng.Column).Delete
    Next u
End Sub

Public Sub DeleteZeroBurstDurUnits(ByRef wb As Workbook)
    'Sheets with the following keywords in their names will have units with bad burst durationss deleted
    Dim keywords(1 To 2) As String, keyword As Variant
    keywords(1) = "_WABs"
    keywords(2) = "_NonWABs"
    
    'Find all sheets with names that match the above keywords
    Dim sht As Worksheet, shtMatches As Boolean
    For Each keyword In keywords
        For Each sht In wb.Worksheets
            shtMatches = (InStr(1, sht.Name, keyword) > 0)
            If shtMatches Then
            
                'Delete units with bad burst durations from those sheets
                Dim lsRows As ListRows, currRow As Long, lr As Long, badDur As Boolean
                Set lsRows = sht.ListObjects(sht.Name).ListRows
                currRow = 1
                For lr = 1 To lsRows.Count
                    badDur = (lsRows(currRow).Range(1, BURST_DUR_COL).Value = 0)
                    If badDur Then
                        lsRows(currRow).Delete
                    Else
                        currRow = currRow + 1
                    End If
                Next lr
                
            End If
        Next sht
    Next keyword
End Sub

Public Sub ExcludeUnits(ByRef wb As Workbook)
    'Sheets with the following keywords in their names will have units with bad burst durationss deleted
    Dim keywords(1 To 2) As String, keyword As Variant
    keywords(1) = "_WABs"
    keywords(2) = "_NonWABs"
    
    'Find all sheets with names that match the above keywords
    Dim sht As Worksheet, shtMatches As Boolean
    For Each keyword In keywords
        For Each sht In wb.Worksheets
            shtMatches = (InStr(1, sht.Name, keyword) > 0)
            If shtMatches Then
            
                'Delete units with bad burst durations from those sheets
                Dim lsRows As ListRows, currRow As Long, lr As Long, badDur As Boolean
                Set lsRows = sht.ListObjects(sht.Name).ListRows
                currRow = 1
                For lr = 1 To lsRows.Count
                    badDur = (lsRows(currRow).Range(1, BURST_DUR_COL).Value = 0)
                    If badDur Then
                        lsRows(currRow).Delete
                    Else
                        currRow = currRow + 1
                    End If
                Next lr
                
            End If
        Next sht
    Next keyword
End Sub

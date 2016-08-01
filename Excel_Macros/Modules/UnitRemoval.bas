Attribute VB_Name = "UnitRemoval"
Option Explicit
Option Private Module

Const DELETE_PREFIX = "XxxX_"
Const NUM_NONPOP_SHEETS = 4
Const BURST_DUR_COL = 4

Public Sub DeleteUnits(ByRef sht As Worksheet, ByVal tissView As cTissueView, ByRef unitNames As Variant)
    'For each invalid unit on this Tissue, mark the columns that contains its data
    'These include the spike timestamp and burst start/end timestamp columns
    Dim cornerCell As Range
    Set cornerCell = sht.Cells(1, 1)
    Dim numUnits As Integer, u As Integer, unit As cUnit, burstOffset As Integer
    numUnits = UBound(unitNames)
    For Each unit In tissView.BadUnits
        For u = 1 To numUnits
            If unitNames(u, 1) = unit.Name And unit.ShouldDelete Then
                burstOffset = numUnits + 2 * (u - 1)
                cornerCell.offset(0, u - 1).Value = DELETE_PREFIX & cornerCell.offset(0, u - 1).Value
                cornerCell.offset(0, burstOffset).Value = DELETE_PREFIX & cornerCell.offset(0, burstOffset).Value
                cornerCell.offset(0, burstOffset + 1).Value = DELETE_PREFIX & cornerCell.offset(0, burstOffset + 1).Value
                Exit For
            End If
        Next u
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
                Dim tbl As ListObject, durs As Variant, currRow As Long, lr As Long, doDelete As Boolean
                Set tbl = sht.ListObjects(sht.Name)
                durs = tbl.DataBodyRange.Columns(BURST_DUR_COL).Value
                currRow = 1
                For lr = 1 To UBound(durs)
                    doDelete = (durs(lr, 1) = 0)
                    If doDelete Then
                        tbl.ListRows(currRow).Delete
                    Else
                        currRow = currRow + 1
                    End If
                Next lr
                
            End If
        Next sht
    Next keyword
End Sub

Public Sub ExcludeUnits(ByRef wb As Workbook)
    'For each data sheet...
    Dim sh As Integer, sht As Worksheet, isSttcSht As Boolean
    Dim numCols As Integer, currRow As Long, lr As Long
    Dim tbl As ListObject, Units As Variant, doExclude As Boolean
    Dim tissID As Integer, unitName1 As String, unitName2 As String, findTissue As Boolean, tissView As cTissueView, unit As cUnit
    For sh = NUM_NONPOP_SHEETS + 1 To wb.Worksheets.Count
        
        'Loop through each unit...
        Set sht = wb.Worksheets(sh)
        isSttcSht = (Right(sht.Name, Len(STTC_NAME)) = STTC_NAME)
        Set tbl = sht.ListObjects(sht.Name)
        numCols = IIf(isSttcSht, 3, 2)
        Units = tbl.DataBodyRange.Resize(tbl.ListRows.Count, numCols).Value
        currRow = 1
        For lr = 1 To UBound(Units)
            tissID = Units(lr, 1)
            unitName1 = Units(lr, 2)
            unitName2 = unitName1
            If isSttcSht Then unitName2 = Units(lr, 3)
            
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
            
            'Delete the unit if it was marked for exclusion
            doExclude = False
            For Each unit In tissView.BadUnits
                doExclude = ((unit.Name = unitName1 Or unit.Name = unitName2) And unit.ShouldExclude)
                If doExclude Then Exit For
            Next unit
            If doExclude Then
                tbl.ListRows(currRow).Delete
            Else
                currRow = currRow + 1
            End If
        Next lr
                        
    Next sh
End Sub

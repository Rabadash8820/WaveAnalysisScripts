Attribute VB_Name = "CombineData"
Option Explicit

Private Const NUM_CONTENTS_COLS = 4

Dim keepOpen As Boolean, dataPaired As Boolean
Dim numAllTissues As Integer
Dim pops As New Dictionary, wbTypes As Variant
Dim ctrlPop As Population
Dim propNames() As String, wbTypePropNames() As String
Dim propWb As Workbook, sttcWb As Workbook
Dim propContentsTbl As ListObject, sttcContentsTbl As ListObject

Public Sub CombineData()
    Call setupOptimizations
    Call initParams
    
    'Only continue if at least one combination operation was selected
    Dim result As VbMsgBoxResult
    Dim thisWb As Workbook
    Set thisWb = ActiveWorkbook
    Dim combineSht As Worksheet, popsSht As Worksheet
    Set combineSht = thisWb.Worksheets(COMBINE_SHEET_NAME)
    Set popsSht = thisWb.Worksheets(POPULATIONS_SHEET_NAME)
    Dim combineProps As Boolean, combineSttc As Boolean
    combineProps = (combineSht.Shapes("CombinePropsChk").OLEFormat.Object.value = 1)
    combineSttc = (combineSht.Shapes("CombineSttcChk").OLEFormat.Object.value = 1)
    If Not combineProps And Not combineSttc Then
        result = MsgBox("No combination operations selected.", vbOKOnly)
        GoTo ExitSub
    End If
    
    'Store the population info (or just return if none was provided)
    Dim popsTbl As ListObject, lsRow As ListRow
    Set popsTbl = popsSht.ListObjects("PopTbl")
    If popsTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No experimental populations have been defined.  Provide this info on the " & POPULATIONS_SHEET_NAME & " sheet", vbOKOnly)
        GoTo ExitSub
    End If
    Dim pop As Population
    For Each lsRow In popsTbl.ListRows
        Set pop = New Population
        pop.ID = lsRow.Range(1, popsTbl.ListColumns("Population ID").Index).value
        pop.name = lsRow.Range(1, popsTbl.ListColumns("Population Name").Index).value
        pop.IsControl = (lsRow.Range(1, popsTbl.ListColumns("Control?").Index).value <> "")
        pops.Add pop.ID, pop
    Next lsRow
    
    'Identify the control population
    Dim numCtrlPops As Integer, p As Integer
    numCtrlPops = 0
    For p = 0 To pops.Count - 1
        Set pop = pops.items()(p)
        If pop.IsControl Then
            Set ctrlPop = pop
            numCtrlPops = numCtrlPops + 1
        End If
    Next p
    If numCtrlPops <> 1 Then
        result = MsgBox("You must identify one (and only one) experimental population as the control.", vbOKOnly)
        GoTo ExitSub
    End If

    'If no Tissue info was provided on the Combine sheet, then just return
    Dim tissueTbl As ListObject
    Dim numTissues As Integer
    Set tissueTbl = combineSht.ListObjects("TissuesTbl")
    numTissues = tissueTbl.ListRows.Count
    If tissueTbl.DataBodyRange Is Nothing Then
        result = MsgBox("No tissue workbook paths were provided.", vbOKOnly)
        GoTo ExitSub
    End If
    
    'Otherwise, create the Tissue objects
    Dim popID As Integer, t As Integer, wbType As String, wbPath As String, numWbTypes As Integer
    wbTypes = tissueTbl.HeaderRowRange(1, 3).Resize(1, tissueTbl.ListColumns.Count - 2).value
    numWbTypes = UBound(wbTypes, 2)
    numAllTissues = tissueTbl.ListRows.Count
    For t = 1 To numWbTypes
        wbType = wbTypes(1, t)
        wbType = Left(wbType, Len(wbType) - Len(" Workbook"))
        wbTypes(1, t) = wbType
    Next t
    Dim tiss As Tissue
    For Each lsRow In tissueTbl.ListRows
        Set tiss = New Tissue
        tiss.ID = lsRow.Range(1, tissueTbl.ListColumns("Tissue ID").Index).value
        popID = lsRow.Range(1, tissueTbl.ListColumns("Population ID").Index).value
        Set tiss.Population = pops(popID)
        For t = 1 To numWbTypes
            wbPath = lsRow.Range(1, tissueTbl.ListColumns(wbTypes(1, t) & " Workbook").Index).value
            tiss.WorkbookPaths.Add wbTypes(1, t), wbPath
        Next t
        pops(popID).Tissues.Add tiss
    Next lsRow
    
    'Let the user pick the workbooks in which to combine data
    Dim propWbName As String, sttcWbName As String
    If combineProps Then _
        propWbName = buildPropsWorkbook.name
    If combineSttc Then
        sttcWbName = pickWorkbook("Select the workbook in which to combine STTC data")
        If sttcWbName = "" Then combineSttc = False
    End If
    If Not combineProps And Not combineSttc Then
        result = MsgBox("No combination operations could be completed.", vbOKOnly)
        GoTo ExitSub
    End If
    If combineProps And combineSttc And propWbName = sttcWbName Then
        result = MsgBox("Property and STTC data must be combined into distinct workbooks.", vbOKOnly)
        GoTo ExitSub
    End If
    
    'Determine if the user wants to keep workbooks open after combining
    keepOpen = (combineSht.Shapes("KeepOpenChk").OLEFormat.Object.value = 1)
    
    'Open these workbooks and combine data into them (they will be left open)
    Dim wb As Workbook, numGoodTissues As Integer
    numGoodTissues = 0
    If combineProps Then
        Dim propKeywords(4) As String
        propKeywords(2) = "_Bursts"
        propKeywords(3) = "_WABs"
        propKeywords(4) = "_NonWABs"
        Set wb = Workbooks.Open(propWbName)
        Call combineIntoWb(wb, wbComboType.PropertyWkbk, propKeywords, numGoodTissues)
    End If
    If combineSttc Then
        Dim sttcKeywords(2) As String
        sttcKeywords(2) = "_STTC"
        Set wb = Workbooks.Open(sttcWbName)
        Call combineIntoWb(wb, wbComboType.SttcWkbk, sttcKeywords, numGoodTissues)
    End If
    
    'Display a succeessful completion dialog
    Dim combineMsg As String
    combineMsg = ""
    If numGoodTissues = 0 Then
        combineMsg = "Data could not be combined into the selected workbooks."
    ElseIf combineProps And combineSttc Then
        combineMsg = "Data was combined into Property and STTC workbooks."
    ElseIf combineProps Then
        combineMsg = "Data was only combined into a Property workbook."
    ElseIf combineSttc Then
        combineMsg = "Data was only combined into an STTC workbook."
    End If
    result = MsgBox(numGoodTissues & " data workbooks for " & numAllTissues & " tissues were successfully opened." & vbCr & _
                    combineMsg & vbCr & _
                    "Time taken: " & Format(ProgramDuration(), "hh:mm:ss"), _
                    vbOKOnly)
                    
ExitSub:
    Call tearDownOptimizations
End Sub

Private Function buildPropsWorkbook() As Workbook
    'Create the new Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    'Build Contents and Stats sheets
    Call buildContentsSheet
    Worksheets.Add After:=Worksheets("Contents")
    Call buildStatsSheet
    
    'Store base property names
    ReDim propNames(1 To 6) As String
    propNames(1) = "Background Firing Rate"
    propNames(2) = "Background Interspike Interval"
    propNames(3) = "Percent of Spikes Occurring in Bursts"
    propNames(4) = "Burst Frequency"
    propNames(5) = "Interburst Interval"
    propNames(6) = "Percent of Bursts Occurring In Waves"
    ReDim wbTypePropNames(1 To 5, 1 To 2) As String
    wbTypePropNames(1, 2) = " Burst Duration"
    wbTypePropNames(2, 2) = " Firing Rate"
    wbTypePropNames(3, 2) = " Interspike Interval"
    wbTypePropNames(4, 1) = "Percent "
    wbTypePropNames(4, 2) = " Time >10 Hz"
    wbTypePropNames(5, 1) = "Spikes Per "
    
    'Build data sheets (one per workbook type per experimental population)
    Dim popV As Variant, pop As Population, wbType As Integer
    For Each popV In pops.items
        Set pop = popV
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        Call buildDataSheet(pop, "Burst", propNames)
        For wbType = 1 To UBound(wbTypes, 2)
            Worksheets.Add After:=Worksheets(Worksheets.Count)
            Call buildDataSheet(pop, wbTypes(1, wbType), wbTypePropNames)
        Next wbType
    Next popV
    
    'Build Figures sheet (must be built last so that table references are valid)
    Worksheets.Add After:=Worksheets("Stats")
    Call buildFiguresSheet
    
    Set buildPropsWorkbook = wb
End Function

Private Sub buildContentsSheet()
                
    'Define some boilerplate variables
    Dim numWbTypes As Integer, t As Integer
    Dim numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    Dim headers() As Variant, data() As Variant
    numWbTypes = UBound(wbTypes, 2)
        
    'Build the Contents sheet
    ActiveSheet.name = "Contents"
        
    'Create the contents table on the Contents sheet
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1"), , xlYes)
    tbl.name = "ContentsTbl"
    
    'Add its columns
    For col = 1 To NUM_CONTENTS_COLS - 1
        tbl.ListColumns.Add
    Next col
    ReDim headers(1 To 1, 1 To NUM_CONTENTS_COLS)
    headers(1, 1) = "Tissue ID"
    headers(1, 2) = "Population ID"
    For t = 1 To numWbTypes
        headers(1, 2 + t) = wbTypes(1, t) & " Workbook"
    Next t
    tbl.HeaderRowRange.value = headers
    
    'Copy data to its DataBodyRange
    Dim contents As Variant, popV As Variant, pop As Population, tiss As Tissue
    ReDim contents(1 To numAllTissues, 1 To NUM_CONTENTS_COLS)
    row = 0
    For Each popV In pops.items
        Set pop = popV
        For Each tiss In pop.Tissues
            row = row + 1
            tbl.ListRows.Add
            contents(row, 1) = tiss.ID
            contents(row, 2) = pop.ID
            For t = 1 To numWbTypes
                contents(row, 2 + t) = tiss.WorkbookPaths(wbTypes(1, t))
            Next t
        Next tiss
    Next popV
    tbl.DataBodyRange.value = contents
    
    'Add the retina count row
    tbl.ShowTotals = True
    tbl.TotalsRowRange(1, 1).value = "Count"
    tbl.ListColumns(2).TotalsCalculation = xlTotalsCalculationCount
    tbl.ListColumns(NUM_CONTENTS_COLS).TotalsCalculation = xlTotalsCalculationNone
    
    'Format sheet
    Cells.VerticalAlignment = xlCenter
    Cells.HorizontalAlignment = xlLeft
    tbl.ListColumns(1).DataBodyRange.HorizontalAlignment = xlCenter
    tbl.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
    Columns.EntireColumn.AutoFit
    Rows.EntireRow.AutoFit

End Sub

Private Sub buildStatsSheet()
                
    'Define some boilerplate variables
    Dim maxTissues As Integer, numWbTypes As Integer, p As Integer, t As Integer
    Dim numCols As Integer, numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    Dim headers() As Variant, data() As Variant
    numWbTypes = UBound(wbTypes, 2)
    
    'Build the Stats sheet
    ActiveSheet.name = "Stats"
        
    'Create the contents table on the Contents sheet
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1"), , xlYes)
    tbl.name = "StatsTbl"
    
    'Add its columns
    numCols = 3
    For col = 1 To numCols - 1
        tbl.ListColumns.Add
    Next col
    ReDim headers(1 To 1, 1 To numCols)
    headers(1, 1) = "Property"
    headers(1, 2) = "Value"
    headers(1, 3) = "Comments"
    tbl.HeaderRowRange.value = headers
    
    'Add data to its DataBodyRange
    numRows = 3
    ReDim data(1 To numRows, 1 To numCols)
    data(1, 1) = "P-Value"
    data(1, 2) = 0.05
    data(2, 1) = "T-Test Tails"
    data(2, 2) = 2
    data(2, 3) = "1 - One-tailed distribution" & Chr(10) & "2 - Two-tailed distribution"
    data(3, 1) = "T-Test Tails"
    data(3, 2) = 3
    data(3, 3) = "1 - Paired" & Chr(10) & "2 - Two-sample equal variance (homoscedastic)" & Chr(10) & "3 - Two-sample unequal variance (heteroscedastic)"
    For row = 1 To numRows
        tbl.ListRows.Add
    Next row
    tbl.DataBodyRange.value = data
    
    'Name the value cells
    Dim valueCol As ListColumn
    Set valueCol = tbl.ListColumns(2)
    valueCol.DataBodyRange(1, 1).name = "PValue"
    valueCol.DataBodyRange(2, 1).name = "TTTails"
    valueCol.DataBodyRange(3, 1).name = "TTType"
    
    'Format sheet
    Cells.VerticalAlignment = xlCenter
    Cells.HorizontalAlignment = xlLeft
    tbl.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
    tbl.ListColumns(3).Range.ColumnWidth = 100  'Just a crazy high number so autofit will work correctly
    Columns.AutoFit
    Rows.AutoFit

End Sub

Private Sub buildDataSheet(ByRef pop As Population, ByVal wbTypeName As String, ByRef headers() As String)
    'Build the Data sheet
    Dim name As String
    name = pop.name & "_" & wbTypeName & "s"
    ActiveSheet.name = name
        
    'Create the Data table on the Data sheet
    Dim cornerCell As Range, tbl As ListObject
    Set cornerCell = Cells(1, 1)
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, cornerCell.offset(1, 0), , xlYes)
    tbl.name = name
    
    'Add its columns
    Dim numCols As Integer, col As Integer
    numCols = UBound(headers)
    For col = 1 To numCols + 2 - 1
        tbl.ListColumns.Add
    Next col
    tbl.HeaderRowRange.Cells(1, 1).value = "Tissue"
    tbl.HeaderRowRange.Cells(1, 2) = "Cell"
    If numDimensions(headers) = 1 Then
        tbl.HeaderRowRange.Cells(1, 3).Resize(1, numCols) = headers
    Else
        Dim headerStrs() As String
        ReDim headerStrs(1 To 1, 1 To numCols)
        For col = 1 To numCols
            headerStrs(1, col) = wbTypePropNames(col, 1) & wbTypeName & wbTypePropNames(col, 2)
        Next col
        tbl.HeaderRowRange.Cells(1, 3).Resize(1, numCols) = headerStrs
    End If

    'Add sheet "headers"
    Application.DisplayAlerts = False
    With Cells(1, 1).Resize(1, 2)
        .value = pop.name
        .Merge
        .Font.Bold = True
        .Font.Size = 16
    End With
    With Cells(1, 3).Resize(1, numCols)
        .value = wbTypeName & "s"
        .Merge
        .Font.Bold = True
        .Font.Size = 16
    End With
    Application.DisplayAlerts = True
    
    'Format sheet
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    tbl.HeaderRowRange.HorizontalAlignment = xlLeft
    ActiveSheet.Visible = xlSheetHidden

End Sub

Private Sub buildFiguresSheet()

    'Define some boilerplate variables
    Dim numWbTypes As Integer, p As Integer, t As Integer
    Dim numCols As Integer, numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    Dim headers() As Variant, data() As Variant
    numWbTypes = UBound(wbTypes, 2)
    Dim cornerCell As Range
    Set cornerCell = Cells(2, 1)
    
    'Build the Figures sheet
    ActiveSheet.name = "Property Figures"
    
    'Store column headers
    Dim rOffset As Integer, cOffset As Integer
    numRows = 1 + NUM_BKGRD_PROPERTIES + numWbTypes * NUM_BURST_PROPERTIES
    numCols = 1 + 7 * pops.Count
    ReDim data(1 To numRows, 1 To numCols)
    data(1, 1) = "Property"
    Dim pop As Population
    For p = 0 To pops.Count - 1
        Set pop = pops.items()(p)
        cOffset = 1 + 4 * p + 1
        data(1, cOffset + 0) = pop.name & "_Avg"
        data(1, cOffset + 1) = pop.name & "_SEM"
        data(1, cOffset + 2) = pop.name & "_%Change"
        data(1, cOffset + 3) = pop.name & "_%Change_SEM"
        data(1, 1 + 4 * pops.Count + 3 * p + 1) = pop.name & "_Value"
        data(1, 1 + 4 * pops.Count + 3 * p + 2) = pop.name & "_%Change"
        data(1, 1 + 4 * pops.Count + 3 * p + 3) = pop.name & "_pValue"
    Next p
    cOffset = 1 + 5 * pops.Count + 1
    
    'Store row headers
    For row = 1 To NUM_BKGRD_PROPERTIES
        data(row + 1, 1) = propNames(row)
    Next row
    Dim wbType As String
    For t = 0 To numWbTypes - 1
        wbType = wbTypes(1, t + 1)
        rOffset = 1 + NUM_BKGRD_PROPERTIES + t * NUM_BURST_PROPERTIES
        For row = 1 To NUM_BURST_PROPERTIES
            data(rOffset + row, 1) = wbTypePropNames(row, 1) & wbType & wbTypePropNames(row, 2)
        Next row
    Next t

    cornerCell.Resize(numRows, numCols).value = data
        
    'Store hidden chart titles
    Dim chartTitles As Variant
    ReDim chartTitles(1 To 1, 1 To numRows - 1)
    For col = 1 To NUM_BKGRD_PROPERTIES
        chartTitles(1, col) = "=" & cornerCell.offset(col, 0).Address & " & "" vs. Experimental Population"""
    Next col
    For t = 0 To numWbTypes - 1
        wbType = wbTypes(1, t + 1)
        cOffset = NUM_BKGRD_PROPERTIES + t * NUM_BURST_PROPERTIES
        For col = 1 To NUM_BURST_PROPERTIES
            chartTitles(1, cOffset + col) = "=" & cornerCell.offset(cOffset + col, 0).Address & " & "" vs. Experimental Population"""
        Next col
    Next t
    
    'Add formatting
    Dim wbTypeStyles() As String
    ReDim wbTypeStyles(1 To numWbTypes)
    wbTypeStyles(1) = "Good"
    wbTypeStyles(2) = "Bad"
    cornerCell.offset(1, 0).Resize(NUM_BKGRD_PROPERTIES, numCols).Style = "Neutral"
    For t = 1 To numWbTypes
        rOffset = NUM_BKGRD_PROPERTIES + (t - 1) * NUM_BURST_PROPERTIES + 1
        cornerCell.offset(rOffset, 0).Resize(NUM_BURST_PROPERTIES, numCols).Style = wbTypeStyles(t)
    Next t
    cornerCell.Resize(numRows, numCols).BorderAround Weight:=xlMedium
    cornerCell.Resize(1, numCols).Borders(xlEdgeBottom).Weight = xlMedium
    cornerCell.Resize(numRows, 1).Borders(xlEdgeRight).Weight = xlMedium
    cornerCell.offset(0, 1 + 4 * pops.Count).Resize(numRows, pops.Count + 2).Borders(xlEdgeLeft).Weight = xlMedium
    cornerCell.offset(NUM_BKGRD_PROPERTIES, 0).Resize(1, numCols).Borders(xlEdgeBottom).Weight = xlThin
    For t = 1 To numWbTypes - 1
        rOffset = NUM_BKGRD_PROPERTIES + t * NUM_BURST_PROPERTIES
        cornerCell.offset(rOffset, 0).Resize(1, numCols).Borders(xlEdgeBottom).Weight = xlThin
    Next t
    For p = 1 To pops.Count - 1
        cornerCell.offset(0, 4 * p).Resize(numRows, 1).Borders(xlEdgeRight).Weight = xlThin
        cornerCell.offset(0, 4 * pops.Count + 3 * p).Resize(numRows, 1).Borders(xlEdgeRight).Weight = xlThin
    Next p
    cornerCell.Resize(1, numCols).Font.Bold = True
    cornerCell.Resize(numRows, 1).Font.Bold = True
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    cornerCell.offset(1, 0).Resize(numRows, 1).HorizontalAlignment = xlLeft
    
    'Determine the max number of Tissues in any Population
    Dim maxTissues As Integer, popV As Variant
    For Each popV In pops.items
        Set pop = popV
        maxTissues = WorksheetFunction.Max(maxTissues, pop.Tissues.Count)
    Next popV
    
    'Build the main percent-change chart
    
    
    'Set up row areas (i.e., for All Bursts bursts from all other workbook types)
    Dim numChartRows As Integer, titleOffset As Integer, propColSpace As Integer
    Dim numPropRows As Integer, numPropCols As Integer
    numChartRows = 15
    titleOffset = 2
    propColSpace = 2
    numPropRows = 2 + numChartRows + 2 + 2 + maxTissues + 1 + 1     'Space + chart + space + headers + tissues + space + line
    numPropCols = 1 + 2 * numWbTypes + propColSpace   'rowHeader + wbTypes + space
    With cornerCell.offset(numPropRows - 1).EntireRow.Interior
        .Pattern = xlSolid
        .TintAndShade = -0.349986266670736
    End With
    With cornerCell.offset(numPropRows + titleOffset, 0)
        .value = "All"
        .Style = "Neutral"
        .Font.Size = 16
        .Font.Bold = True
        .Orientation = 90
        .Resize(numPropRows - 2 * titleOffset - 1, 1).Merge
    End With
    rOffset = numPropRows + titleOffset
    For t = 1 To numWbTypes
        rOffset = (t + 1) * numPropRows
        With cornerCell.offset(rOffset - 1, 0).EntireRow.Interior
            .Pattern = xlSolid
            .TintAndShade = -0.349986266670736
        End With
        With cornerCell.offset(rOffset + titleOffset, 0)
            .value = wbTypes(1, t)
            .Style = wbTypeStyles(t)
            .Font.Size = 16
            .Font.Bold = True
            .Orientation = 90
            .Resize(numPropRows - 2 * titleOffset - 1, 1).Merge
        End With
    Next t
    
    'Build property areas
    Dim prop As Integer
    Dim propCornerCell As Range, tblRowRng As Range, chartCell As Range
    rOffset = numPropRows + 2 + numChartRows + 1
    For prop = 1 To NUM_BKGRD_PROPERTIES
        Set tblRowRng = cornerCell.offset(prop, 0)
        cOffset = numCols + (prop - 1) * numPropCols
        Set propCornerCell = cornerCell.offset(rOffset, cOffset)
        Set chartCell = propCornerCell.offset(-(numChartRows + 1), 0)
        Call buildPropArea(propCornerCell, tblRowRng, chartCell, PROP_UNITS(prop), "Burst", maxTissues)
    Next prop
    For t = 1 To numWbTypes
        rOffset = (t + 1) * numPropRows + 2 + numChartRows + 1
        For prop = 1 To NUM_BURST_PROPERTIES
            Set tblRowRng = cornerCell.offset(NUM_BKGRD_PROPERTIES + (t - 1) * NUM_BURST_PROPERTIES + prop, 0)
            cOffset = numCols + (prop - 1) * numPropCols
            Set propCornerCell = cornerCell.offset(rOffset, cOffset)
            Set chartCell = propCornerCell.offset(-(numChartRows + 1), 0)
            Call buildPropArea(propCornerCell, tblRowRng, chartCell, PROP_UNITS(NUM_BKGRD_PROPERTIES + prop), wbTypes(1, t), maxTissues)
        Next prop
    Next t
    
    'Final formatting...
    Columns.AutoFit
    Rows.AutoFit
    cornerCell.offset(-1, 0).Resize(1, numRows - 1).value = chartTitles
    Rows(cornerCell.row - 1).Hidden = True
    Range(Columns(2), Columns(4 * pops.Count + 1)).Hidden = True

End Sub

Private Sub buildPropArea(ByRef cornerCell As Range, ByRef tblRowCell As Range, ByRef chartCell As Range, ByVal unitsStr As String, ByVal wbType As String, ByVal maxTissues As Integer)

    Dim numWbTypes As Integer, numHeaders As Integer
    Dim t As Integer, pop As Population, p As Integer
    Dim rOffset As Integer, cOffset As Integer
    numWbTypes = UBound(wbTypes, 2)
    numHeaders = 1 + 2 * numWbTypes
    Dim headers() As Variant
    ReDim headers(1 To 1, 1 To numHeaders)
    
    'Draw the property title
    cornerCell.offset(1, 0).Formula = "=" & tblRowCell.Address
    With cornerCell.offset(1, 0).Resize(1, numHeaders)
        .Merge
        .Font.Bold = True
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    'Write data summary headers
    headers(1, 1) = "Tissue"
    Dim numPopCols As Integer
    numPopCols = 2
    For p = 0 To pops.Count - 1
        Set pop = pops.items()(p)
        cOffset = 1 + p * numPopCols
        headers(1, cOffset + 1) = pop.name
        headers(1, cOffset + 2) = pop.name & "_%Change"
    Next p
    cornerCell.offset(2, 0).Resize(1, numHeaders).value = headers
    cornerCell.offset(2, 0).Resize(1, numHeaders).Font.Bold = True
    
    'Identify the control population's data range
    Dim ctrlRng As Range
    For p = 0 To pops.Count - 1
        Set pop = pops.items()(p)
        If pop.ID = ctrlPop.ID Then
            cOffset = numPopCols * p
            Set ctrlRng = cornerCell.offset(3, cOffset + 1).Resize(maxTissues, 1)
            Exit For
        End If
    Next p
    
    'Write tissue data
    Dim tissueRng As Range
    For t = 1 To maxTissues
        rOffset = 2 + t
        cornerCell.offset(rOffset, 0).value = t
        For p = 0 To pops.Count - 1
            Set pop = pops.items()(p)
            cOffset = numPopCols * p
            Set tissueRng = cornerCell.offset(rOffset, cOffset + 1)
            tissueRng.offset(0, 0).Formula = "=IFERROR(AVERAGEIF(" & pop.name & "_" & wbType & "s[Tissue]," & t & "," & pop.name & "_" & wbType & "s[" & tblRowCell.value & "]), """")"
            tissueRng.offset(0, 1).Formula = "=IFERROR(100*(" & tissueRng.Address & "-AVERAGE(" & ctrlRng.Address & "))/AVERAGE(" & ctrlRng.Address & "), """")"
        Next p
    Next t
    cornerCell.offset(3, 0).Resize(maxTissues, 1).Font.Bold = True
    
    'Add formulas to the main table
    Dim formulaRng As Range, propRng As Range, summaryRng As Range
    For p = 0 To pops.Count - 1
        Set formulaRng = tblRowCell.offset(0, 4 * p)
        Set propRng = cornerCell.offset(3, 2 * p).Resize(maxTissues, 1)
        formulaRng.offset(0, 1).Formula = "=IFERROR(AVERAGE(" & propRng.offset(0, 1).Address & "), """")"
        formulaRng.offset(0, 2).Formula = "=IFERROR(STDEV.S(" & propRng.offset(0, 1).Address & ")/SQRT(COUNT(" & propRng.offset(0, 1).Address & ")), """")"
        formulaRng.offset(0, 3).Formula = "=IFERROR(AVERAGE(" & propRng.offset(0, 2).Address & "), """")"
        formulaRng.offset(0, 4).Formula = "=IFERROR(STDEV.S(" & propRng.offset(0, 2).Address & ")/SQRT(COUNT(" & propRng.offset(0, 2).Address & ")), """")"
        Set summaryRng = tblRowCell.offset(0, 4 * pops.Count + 3 * p)
        summaryRng.offset(0, 1).Formula = "=TEXT(" & formulaRng.offset(0, 1).Address & ",""0.000"")&"" ± ""&TEXT(" & formulaRng.offset(0, 2).Address & ",""0.000"")"
        summaryRng.offset(0, 2).Formula = "=TEXT(" & formulaRng.offset(0, 3).Address & ",""0.0"")&"" ± ""&TEXT(" & formulaRng.offset(0, 4).Address & ",""0.0"")"
        summaryRng.offset(0, 3).Formula = "=IFERROR(TEXT(T.TEST(" & cornerCell.offset(3, 2 * p + 1).Resize(maxTissues, 1).Address & "," & ctrlRng.Address & ",TTTails,TTType), ""0.000""), """")"
    Next p
    
    'Add bar chart
    Dim chartShp As Shape
    Set chartShp = buildPropChart(tblRowCell, unitsStr)
    chartShp.Left = chartCell.Left
    chartShp.Top = chartCell.Top
    chartShp.Width = cornerCell.Resize(1, numHeaders).Width
    chartShp.Height = cornerCell.offset(-2, 0).Top - chartCell.Top

End Sub

Private Function buildPropChart(ByRef tblRowCell As Range, ByVal unitsStr As String) As Shape

    'Add the new chart object
    Dim chartShp As Shape
    Set chartShp = ActiveSheet.Shapes.AddChart(xlColumnClustered)
    With chartShp
        .name = Replace(tblRowCell.value, " ", "_") & "_Chart"
        .Line.Visible = False
    End With
    
    With chartShp.Chart
    
        'Select its data
        Dim errorRng As Range
        Set errorRng = tblRowCell.offset(0, 0 + 2)
        Dim s As Integer
        For s = 1 To .SeriesCollection.Count
            .SeriesCollection(1).Delete
        Next s
        .PlotVisibleOnly = False
        Dim p As Integer, pop As Population
        For p = 0 To pops.Count - 1
            Set pop = pops.items()(p)
            Set errorRng = tblRowCell.offset(0, 4 * p + 2)
            .SeriesCollection.Add Source:=tblRowCell.offset(0, 4 * p + 1)
            .SeriesCollection(p + 1).name = pop.name
            .SeriesCollection(p + 1).ErrorBar Direction:=xlY, include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeCustom, _
                Amount:="='" & tblRowCell.Worksheet.name & "'!" & errorRng.Address, _
                minusvalues:="='" & tblRowCell.Worksheet.name & "'!" & errorRng.Address
        Next p
    
        'Format the Chart
        .HasAxis(xlCategory) = True
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionNone
            .MajorTickMark = xlTickMarkNone
            .Format.Line.foreColor.RGB = vbBlack
            .Format.Line.Weight = 3
            .HasTitle = False
        End With
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .HasMajorGridlines = False
            .TickLabels.Font.Color = vbBlack
            .TickLabels.Font.Bold = True
            .TickLabels.Font.Size = 10
            .Format.Line.foreColor.RGB = vbBlack
            .Format.Line.Weight = 3
            .HasTitle = True
            .AxisTitle.Caption = unitsStr
            .AxisTitle.Font.Bold = True
            .AxisTitle.Font.Color = vbBlack
            .AxisTitle.Font.Size = 12
        End With
        .HasTitle = True
        .HasTitle = False   'I have ABSOLUTELY no idea why this f*cking toggle is necessary, but a runtime error occurs without it
        .HasTitle = True
        With .chartTitle
            .Text = "='" & tblRowCell.Worksheet.name & "'!" & Cells(1, tblRowCell.row - 2).Address
            .Font.Color = vbBlack
            .Font.Bold = True
            .Font.Size = 18
        End With
        .HasLegend = True
        With .Legend
            .Font.Color = vbBlack
            .Font.Bold = True
            .Font.Size = 10
            .Position = xlRight
        End With
        
        'Formats the Chart's Series
        Dim foreColors() As Long, numColors As Integer, c As Integer
        numColors = 2
        ReDim foreColors(0 To numColors - 1)
        foreColors(0) = vbBlack
        foreColors(1) = vbRed
        c = -1
        For p = 1 To pops.Count
            c = c + 1
            With .FullSeriesCollection(p)
                .Format.Fill.foreColor.RGB = foreColors(c Mod numColors)
                .Format.Line.foreColor.RGB = vbBlack
                .Format.Line.Weight = 2
                .ErrorBars.Format.Line.foreColor.RGB = vbBlack
                .ErrorBars.Format.Line.Weight = 2
            End With
        Next p
        
    End With
    
    Set buildPropChart = chartShp
End Function

Private Function uniqueItems(ByRef items As Variant) As Variant

    Dim unique() As String
    ReDim unique(1 To UBound(items))
    Dim numUniques As Integer, str As Variant, s As Integer, found As Boolean
    
    'For each provided string...
    numUniques = 0
    For Each str In items
        'Check if this string has already been stored
        found = False
        For s = 1 To numUniques
            If (unique(s) = str) Then
                found = True
                Exit For
            End If
        Next s
        
        'If not, then store it
        If Not found Then
            numUniques = numUniques + 1
            unique(numUniques) = str
        End If
    Next str
    
    'Trim up the unique item array and return it
    ReDim Preserve unique(1 To numUniques)
    uniqueItems = unique
End Function

Private Sub combineIntoWb(ByRef wb As Workbook, ByVal wbType As wbComboType, ByRef tblKeywords() As String, ByRef numGoodTissues As Integer)
    'Open combination workbook (it will be kept open)
    Dim contentsTbl As ListObject
    Set contentsTbl = wb.Worksheets(CONTENTS_SHEET_NAME).ListObjects(CONTENTS_SHEET_NAME)
    
    'Clear any existing data in that workbook
    Dim k As Integer
    For k = 1 To UBound(tblKeywords)
        Call clearTables(wb, tblKeywords(k))
    Next k
    
    'Open it, fetch its data, and re-close it
    Dim popV As Variant, pop As Population, t As Tissue
    For Each popV In pops.items
        Set pop = popV
        For Each t In pop.Tissues
            Call fetchRetina(wbType, wb, t, numGoodTissues)
        Next t
    Next popV
    
    'Pretty up the sheets now that data is present
    For k = 1 To UBound(tblKeywords)
        Call cleanSheets(wb, tblKeywords(k))
    Next k
    
    'Save/close the workbook if the user doesn't want to keep it open
    If Not keepOpen Then _
        Call wb.Close(True)

End Sub

Private Sub fetchRetina(ByVal wbType As wbComboType, ByRef summaryWb As Workbook, ByRef Tissue As Tissue, ByRef numGoodTissues As Integer)
    'Make sure an ID was provided for this retina
    Dim result As VbMsgBoxResult
    If Tissue.ID = "" Then
        result = MsgBox("A tissue in population " & Tissue.Population.name & " was not given an ID." & vbCr & _
                        "Data will not be loaded.")
        Exit Sub
    End If
    
    'If so, then Initialize some local variables
    Dim fs As New FileSystemObject
    Dim retinaWb As Workbook
    Dim numWbTypes As Integer
    numWbTypes = UBound(wbTypes, 2)
    
    'For each type of data, load that data from its provided workbook (if it exists),
    'and store it in the workbooks requested by the user
    Dim wbTypeName As Variant, wbPath As String
    For Each wbTypeName In wbTypes
        wbPath = Tissue.WorkbookPaths(wbTypeName)
        If fs.FileExists(wbPath) Then
            Set retinaWb = Workbooks.Open(wbPath)
            Call fetchRetinaData(wbType, Tissue.ID, summaryWb, retinaWb, Tissue.Population.name, wbTypeName)
            retinaWb.Close
            numGoodTissues = numGoodTissues + 1
        ElseIf wbPath = "" Then
            result = MsgBox("No " & wbType & " workbook provided for tissue " & Tissue.ID & " in population " & Tissue.Population.name, vbOKOnly)
        Else
            result = MsgBox("Workbook " & wbPath & " could not be found" & vbCr & _
                            "Make sure you provided the correct " & wbTypeName & " path.", vbOKOnly)
        End If
    Next wbTypeName

End Sub

Private Sub clearTables(ByRef wb As Workbook, ByVal keyword As String)
    Dim sht As Worksheet, tbl As ListObject
    Dim needsClearing As Boolean
    
    'Clear the data table on each sheet with the given keyword in the name
    For Each sht In wb.Worksheets
        needsClearing = (InStr(1, sht.name, keyword) > 0)
        If needsClearing Then
            Set tbl = sht.ListObjects(sht.name)
            If Not tbl.DataBodyRange Is Nothing Then _
                tbl.DataBodyRange.Delete
        End If
    Next sht
End Sub

Private Sub fetchRetinaData(ByVal wbType As wbComboType, ByVal retinaID, ByRef summaryWb As Workbook, ByRef retinaWb As Workbook, ByVal popName As String, ByVal typeName As String)
    Select Case wbType
    
        'Combine data for Property workbooks
        Case wbComboType.PropertyWkbk
            Select Case typeName
                Case "Processed WAB Workbook"
                    Call copyRetinaData(summaryWb, retinaWb, "All Avgs", "AllAvgsTbl", popName & "_All", popName & "_All", retinaID)
                    Call copyRetinaData(summaryWb, retinaWb, "Burst Avgs", "BurstAvgsTbl", popName & "_WABs", popName & "_WABs", retinaID)
                    
                Case "Processed NonWAB Workbook"
                    Call copyRetinaData(summaryWb, retinaWb, "Burst Avgs", "BurstAvgsTbl", popName & "_NonWABs", popName & "_NonWABs", retinaID)
            End Select
            
        'Combine data for STTC workbooks
        Case wbComboType.SttcWkbk
            Select Case typeName
                Case "Processed WAB Workbook"
                    Call copyRetinaData(summaryWb, retinaWb, "STTC", "SttcTbl", popName & "_STTC", popName & "_STTC", retinaID)
            End Select
            
    End Select
End Sub

Private Sub copyRetinaData(ByRef summaryWb As Workbook, ByRef retinaWb As Workbook, _
                           ByVal fetchSheetName As String, ByVal fetchTableName As String, ByVal outputSheetName As String, ByVal outputTableName As String, _
                           ByVal retinaID As String)
    'Set the Range of data to be copied from the retinal workbook
    Dim fetchRng As Range
    Set fetchRng = retinaWb.Worksheets(fetchSheetName).ListObjects(fetchTableName).DataBodyRange
    
    'Set the Range to be copied to in the summary workbook
    Dim outputTbl As ListObject
    Set outputTbl = summaryWb.Worksheets(outputSheetName).ListObjects(outputTableName)
    outputTbl.ListRows.Add
    Dim outputRng As Range
    Set outputRng = outputTbl.ListRows(outputTbl.ListRows.Count).Range.Cells(1, 2)
    
    'Copy the data, and add the provided retinaID to each row
    fetchRng.Copy Destination:=outputRng
    Dim idRng As Range
    Set idRng = outputRng.offset(0, -1).Resize(fetchRng.Rows.Count, 1)
    idRng.value = retinaID

End Sub

Private Sub cleanSheets(ByRef wb As Workbook, ByVal keyword As String)
    Dim sht As Worksheet
    Dim needsCleaning As Boolean
    
    'Clear the data table on each sheet with the given keyword in the name
    For Each sht In wb.Worksheets
        needsCleaning = (InStr(1, sht.name, keyword) > 0)
        If needsCleaning Then
            sht.Columns.EntireColumn.AutoFit
            sht.Rows.EntireRow.AutoFit
        End If
    Next sht
End Sub

Private Function numDimensions(ByRef arr As Variant)
    'Sets up the error handler.
    On Error GoTo FinalDimension
    
    'VBA arrays can have up to 60,000 dimensions
    'Do something with each dimension until an error is generated
    Dim numDims As Long, temp As Integer
    For numDims = 1 To 60000
       temp = LBound(arr, numDims)
    Next numDims
    
    ' The error routine.
FinalDimension:
    numDimensions = numDims - 1

End Function

Attribute VB_Name = "CombineData"
Option Explicit

Private Const NUM_CONTENTS_COLS = 4

Dim keepOpen As Boolean, dataPaired As Boolean
Dim numAllTissues As Integer
Dim propNames() As String, wbTypePropNames() As String
Dim combineWb As Workbook

Public Sub CombineData()
    Call setupOptimizations
    
    Call GetConfigVars
    Call DefinePopulations
    
    'Initialize some variables
    Dim combineSht As Worksheet, popsSht As Worksheet, tissueSht As Worksheet, tissueTbl As ListObject
    Set combineSht = ThisWorkbook.Worksheets(COMBINE_SHT_NAME)
    Set popsSht = ThisWorkbook.Worksheets(POPS_SHT_NAME)
    Set tissueSht = ThisWorkbook.Worksheets(TISSUES_SHT_NAME)
    Set tissueTbl = tissueSht.ListObjects(TISSUES_TBL_NAME)
    numAllTissues = tissueTbl.ListRows.Count
    
    'Get some user options from the Form Controls within the workbook
    keepOpen = (combineSht.Shapes("KeepOpenChk").OLEFormat.Object.value = 1)
    dataPaired = (popsSht.Shapes("DataPairedChk").OLEFormat.Object.value = 1)
    
    'Build the workbook in which to combine data, and combine data!
    Set combineWb = buildComboWorkbook()
    Dim numGoodTissues As Integer
    numGoodTissues = 0
    Call combineIntoWb(numGoodTissues)
    
    'Display a succeessful completion dialog
    Dim combineMsg As String, result As VbMsgBoxResult
    combineMsg = "Property and STTC data was combined into the provided workbook"
    result = MsgBox(numGoodTissues & " data workbooks for " & numAllTissues & " tissues were successfully opened." & vbCr & _
                    combineMsg & vbCr & _
                    "Time taken: " & Format(ProgramDuration(), "hh:mm:ss"), _
                    vbOKOnly)
                    
    Call tearDownOptimizations
End Sub

Private Function buildComboWorkbook() As Workbook
    'Create the new Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    'Build Contents and Stats sheets
    Call buildContentsSheet
    Worksheets.Add After:=Worksheets(CONTENTS_SHT_NAME)
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
    Dim sttcHeaders() As String
    ReDim sttcHeaders(1 To 5)
    sttcHeaders(1) = "Tissue"
    sttcHeaders(2) = "Cell1"
    sttcHeaders(3) = "Cell2"
    sttcHeaders(4) = "Unit Distance"
    sttcHeaders(5) = "STTC"
    
    'Build data sheets (one per workbook type per experimental population)
    Dim popV As Variant, pop As Population, bType As Integer
    For Each popV In POPULATIONS.items
        Set pop = popV
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        Call buildSttcDataSheet(pop, sttcHeaders)
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        Call buildDataSheet(pop, "Burst", propNames)
        For bType = 1 To UBound(BURST_TYPES, 2)
            Worksheets.Add After:=Worksheets(Worksheets.Count)
            Call buildDataSheet(pop, BURST_TYPES(1, bType), wbTypePropNames)
        Next bType
    Next popV

    'Build Figures sheets (must be built last so that table references are valid)
    Worksheets.Add After:=Worksheets(STATS_SHT_NAME)
    Call buildPropFiguresSheet
    Worksheets.Add After:=Worksheets(PROPFIGS_SHT_NAME)
    Call buildSttcFiguresSheet

    
    Set buildComboWorkbook = wb
End Function

Private Sub buildContentsSheet()
                
    'Define some boilerplate variables
    Dim numBurstTypes As Integer, t As Integer
    Dim numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    Dim headers() As Variant, data() As Variant
    numBurstTypes = UBound(BURST_TYPES, 2)
        
    'Build the Contents sheet
    ActiveSheet.name = CONTENTS_SHT_NAME
        
    'Create the contents table on the Contents sheet
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1"), , xlYes)
    tbl.name = CONTENTS_TBL_NAME
    
    'Add its columns
    For col = 1 To NUM_CONTENTS_COLS - 1
        tbl.ListColumns.Add
    Next col
    ReDim headers(1 To 1, 1 To NUM_CONTENTS_COLS)
    headers(1, 1) = "Tissue ID"
    headers(1, 2) = "Population ID"
    For t = 1 To numBurstTypes
        headers(1, 2 + t) = BURST_TYPES(1, t) & " Workbook"
    Next t
    tbl.HeaderRowRange.value = headers
    
    'Copy data to its DataBodyRange
    Dim contents As Variant, popV As Variant, pop As Population, tiss As Tissue
    ReDim contents(1 To numAllTissues, 1 To NUM_CONTENTS_COLS)
    row = 0
    For Each popV In POPULATIONS.items
        Set pop = popV
        For Each tiss In pop.Tissues
            row = row + 1
            tbl.ListRows.Add
            contents(row, 1) = tiss.ID
            contents(row, 2) = pop.ID
            For t = 1 To numBurstTypes
                contents(row, 2 + t) = tiss.WorkbookPaths(BURST_TYPES(1, t))
            Next t
        Next tiss
    Next popV
    tbl.DataBodyRange.value = contents
    
    'Add the tissue count row
    tbl.ShowTotals = True
    tbl.TotalsRowRange(1, 1).value = "Count"
    tbl.ListColumns(2).TotalsCalculation = xlTotalsCalculationCount
    tbl.ListColumns(NUM_CONTENTS_COLS).TotalsCalculation = xlTotalsCalculationNone
    
    'Format sheet
    'Columns/rows will be autofitted after combining data
    Cells.VerticalAlignment = xlCenter
    Cells.HorizontalAlignment = xlLeft
    tbl.ListColumns(1).DataBodyRange.HorizontalAlignment = xlCenter
    tbl.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter

End Sub

Private Sub buildStatsSheet()
                
    'Define some boilerplate variables
    Dim maxTissues As Integer, numBurstTypes As Integer, p As Integer, t As Integer
    Dim numCols As Integer, numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    Dim headers() As Variant, data() As Variant
    numBurstTypes = UBound(BURST_TYPES, 2)
    
    'Build the Stats sheet
    ActiveSheet.name = STATS_SHT_NAME
        
    'Create the contents table on the Contents sheet
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1"), , xlYes)
    tbl.name = STATS_TBL_NAME
    
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
    Dim numCols As Integer
    numCols = UBound(headers, 1)
    Dim cornerCell As Range, tbl As ListObject
    Set cornerCell = Cells(1, 1)
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, cornerCell.offset(1, 0).Resize(1, numCols + 2), , xlYes)
    tbl.name = name
    
    'Add its columns
    Dim col As Integer
    numCols = UBound(headers)
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
    Dim popNameCols As Integer
    popNameCols = 2
    Application.DisplayAlerts = False
    With Cells(1, 1).Resize(1, popNameCols)
        .value = pop.name
        .Merge
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = pop.ForeColor
        .Interior.Color = pop.BackColor
    End With
    With Cells(1, popNameCols + 1).Resize(1, numCols)
        .value = wbTypeName & "s"
        .Merge
        .Font.Bold = True
        .Font.Size = 16
    End With
    Application.DisplayAlerts = True
    
    'Format sheet
    'Columns/rows will be autofitted after combining data
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    tbl.HeaderRowRange.HorizontalAlignment = xlLeft
    ActiveSheet.Visible = xlSheetHidden

End Sub

Private Sub buildSttcDataSheet(ByRef pop As Population, ByRef sttcHeaders() As String)
    'Build the Data sheet
    Dim name As String
    name = pop.name & "_STTC"
    ActiveSheet.name = name
        
    'Create the Data table on the Data sheet
    Dim cornerCell As Range, tbl As ListObject
    Set cornerCell = Cells(1, 1)
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, cornerCell.offset(1, 0).Resize(1, UBound(sttcHeaders, 1)), , xlYes)
    tbl.name = name
    
    'Add its columns
    Dim numCols As Integer
    numCols = UBound(sttcHeaders)
    tbl.HeaderRowRange.value = sttcHeaders

    'Add sheet "headers"
    Dim popNameCols As Integer
    popNameCols = 3
    Application.DisplayAlerts = False
    With Cells(1, 1).Resize(1, popNameCols)
        .value = pop.name
        .Merge
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = pop.ForeColor
        .Interior.Color = pop.BackColor
    End With
    With Cells(1, popNameCols + 1).Resize(1, numCols - popNameCols)
        .value = "STTC"
        .Merge
        .Font.Bold = True
        .Font.Size = 16
    End With
    Application.DisplayAlerts = True
    
    'Format sheet
    'Columns/rows will be autofitted after combining data
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    tbl.HeaderRowRange.HorizontalAlignment = xlLeft
    ActiveSheet.Visible = xlSheetHidden

End Sub

Private Sub buildPropFiguresSheet()

    'Define some boilerplate variables
    Dim numBurstTypes As Integer, p As Integer, t As Integer
    Dim numCols As Integer, numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    Dim headers() As Variant, data() As Variant
    numBurstTypes = UBound(BURST_TYPES, 2)
    Dim cornerCell As Range
    Set cornerCell = Cells(2, 1)
    
    'Build the Figures sheet
    ActiveSheet.name = PROPFIGS_SHT_NAME

    'Store column headers
    Dim rOffset As Integer, cOffset As Integer
    numRows = 1 + NUM_BKGRD_PROPERTIES + numBurstTypes * NUM_BURST_PROPERTIES
    numCols = 1 + 7 * POPULATIONS.Count
    ReDim data(1 To numRows, 1 To numCols)
    data(1, 1) = "Property"
    Dim pop As Population
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items()(p)
        cOffset = 1 + 4 * p + 1
        data(1, cOffset + 0) = pop.Abbreviation & "_Avg"
        data(1, cOffset + 1) = pop.Abbreviation & "_SEM"
        data(1, cOffset + 2) = pop.Abbreviation & "_%Change"
        data(1, cOffset + 3) = pop.Abbreviation & "_%Change_SEM"
        data(1, 1 + 4 * POPULATIONS.Count + 3 * p + 1) = pop.Abbreviation & "_Avg"
        data(1, 1 + 4 * POPULATIONS.Count + 3 * p + 2) = pop.Abbreviation & "_%Change"
        data(1, 1 + 4 * POPULATIONS.Count + 3 * p + 3) = pop.Abbreviation & "_pValue"
    Next p
    cOffset = 1 + 5 * POPULATIONS.Count + 1

    'Store row headers
    For row = 1 To NUM_BKGRD_PROPERTIES
        data(row + 1, 1) = propNames(row)
    Next row
    Dim bType As String
    For t = 0 To numBurstTypes - 1
        bType = BURST_TYPES(1, t + 1)
        rOffset = 1 + NUM_BKGRD_PROPERTIES + t * NUM_BURST_PROPERTIES
        For row = 1 To NUM_BURST_PROPERTIES
            data(rOffset + row, 1) = wbTypePropNames(row, 1) & bType & wbTypePropNames(row, 2)
        Next row
    Next t

    cornerCell.Resize(numRows, numCols).value = data

    'Store hidden chart titles
    Dim chartTitles As Variant
    ReDim chartTitles(1 To 1, 1 To numRows - 1)
    For col = 1 To NUM_BKGRD_PROPERTIES
        chartTitles(1, col) = "=" & cornerCell.offset(col, 0).Address & " & "" vs. Experimental Population"""
    Next col
    For t = 0 To numBurstTypes - 1
        bType = BURST_TYPES(1, t + 1)
        cOffset = NUM_BKGRD_PROPERTIES + t * NUM_BURST_PROPERTIES
        For col = 1 To NUM_BURST_PROPERTIES
            chartTitles(1, cOffset + col) = "=" & cornerCell.offset(cOffset + col, 0).Address & " & "" vs. Experimental Population"""
        Next col
    Next t

    'Add formatting
    Dim wbTypeStyles() As String
    ReDim wbTypeStyles(1 To numBurstTypes)
    wbTypeStyles(1) = "Good"
    wbTypeStyles(2) = "Bad"
    cornerCell.offset(1, 0).Resize(NUM_BKGRD_PROPERTIES, numCols).Style = "Neutral"
    For t = 1 To numBurstTypes
        rOffset = NUM_BKGRD_PROPERTIES + (t - 1) * NUM_BURST_PROPERTIES + 1
        cornerCell.offset(rOffset, 0).Resize(NUM_BURST_PROPERTIES, numCols).Style = wbTypeStyles(t)
    Next t
    cornerCell.Resize(numRows, numCols).BorderAround Weight:=xlMedium
    cornerCell.Resize(1, numCols).Borders(xlEdgeBottom).Weight = xlMedium
    cornerCell.Resize(numRows, 1).Borders(xlEdgeRight).Weight = xlMedium
    cornerCell.offset(0, 1 + 4 * POPULATIONS.Count).Resize(numRows, POPULATIONS.Count + 2).Borders(xlEdgeLeft).Weight = xlMedium
    cornerCell.offset(NUM_BKGRD_PROPERTIES, 0).Resize(1, numCols).Borders(xlEdgeBottom).Weight = xlThin
    For t = 1 To numBurstTypes - 1
        rOffset = NUM_BKGRD_PROPERTIES + t * NUM_BURST_PROPERTIES
        cornerCell.offset(rOffset, 0).Resize(1, numCols).Borders(xlEdgeBottom).Weight = xlThin
    Next t
    cornerCell.offset(0, 1).Resize(1, 4).Interior.Color = POPULATIONS.items()(0).BackColor
    cornerCell.offset(0, 1).Resize(1, 4).Font.Color = POPULATIONS.items()(0).ForeColor
    cornerCell.offset(0, 4 * POPULATIONS.Count + 1).Resize(1, 3).Interior.Color = POPULATIONS.items()(0).BackColor
    cornerCell.offset(0, 4 * POPULATIONS.Count + 1).Resize(1, 3).Font.Color = POPULATIONS.items()(0).ForeColor
    For p = 1 To POPULATIONS.Count - 1
        cornerCell.offset(0, 4 * p + 1).Resize(1, 4).Interior.Color = POPULATIONS.items()(p).BackColor
        cornerCell.offset(0, 4 * p + 1).Resize(1, 4).Font.Color = POPULATIONS.items()(p).ForeColor
        cornerCell.offset(0, 4 * POPULATIONS.Count + 3 * p + 1).Resize(1, 3).Interior.Color = POPULATIONS.items()(p).BackColor
        cornerCell.offset(0, 4 * POPULATIONS.Count + 3 * p + 1).Resize(1, 3).Font.Color = POPULATIONS.items()(p).ForeColor
        cornerCell.offset(0, 4 * p).Resize(numRows, 1).Borders(xlEdgeRight).Weight = xlThin
        cornerCell.offset(0, 4 * POPULATIONS.Count + 3 * p).Resize(numRows, 1).Borders(xlEdgeRight).Weight = xlThin
    Next p
    cornerCell.Resize(1, numCols).Font.Bold = True
    cornerCell.Resize(numRows, 1).Font.Bold = True
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    cornerCell.offset(1, 0).Resize(numRows, 1).HorizontalAlignment = xlLeft

    'Determine the max number of Tissues in any Population
    Dim maxTissues As Integer, popV As Variant
    For Each popV In POPULATIONS.items
        Set pop = popV
        maxTissues = WorksheetFunction.Max(maxTissues, pop.Tissues.Count)
    Next popV

    'Set up row areas (i.e., for All Bursts bursts from all other workbook types)
    Dim numChartRows As Integer, titleOffset As Integer, propColSpace As Integer
    Dim numPropRows As Integer, numPropCols As Integer
    numChartRows = 15
    titleOffset = 2
    propColSpace = 2
    numPropRows = 2 + numChartRows + 2 + 2 + maxTissues + 1 + 1     'Space + chart + space + headers + tissues + space + line
    numPropCols = 1 + 2 * POPULATIONS.Count + propColSpace   'rowHeader + BURST_TYPES + space
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
    For t = 1 To numBurstTypes
        rOffset = (t + 1) * numPropRows
        With cornerCell.offset(rOffset - 1, 0).EntireRow.Interior
            .Pattern = xlSolid
            .TintAndShade = -0.349986266670736
        End With
        With cornerCell.offset(rOffset + titleOffset, 0)
            .value = BURST_TYPES(1, t)
            .Style = wbTypeStyles(t)
            .Font.Size = 16
            .Font.Bold = True
            .Orientation = 90
            .Resize(numPropRows - 2 * titleOffset - 1, 1).Merge
        End With
    Next t

    'Add the new column chart object for percent changes
    Dim chartShp As Shape, chartRng As Range
    Set chartRng = cornerCell.offset(0, numCols + 2).Resize(numPropRows - 2, 7)
    Set chartShp = ActiveSheet.Shapes.AddChart(xlBarClustered, chartRng.Left, chartRng.Top, chartRng.Width, chartRng.Height)
    With chartShp
        .name = "Percent_Change_Chart"
        .Line.Visible = False
    End With

    With chartShp.Chart

        'Remove default Series
        Dim s As Integer
        For s = 1 To .SeriesCollection.Count
            .SeriesCollection(1).Delete
        Next s

        'Set the new population Series (showing future hidden cells too)
        .PlotVisibleOnly = False
        Dim errorRng As Range, numSeries As Integer
        numSeries = 0
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            If pop.ID <> CTRL_POP.ID Then
                numSeries = numSeries + 1
                .SeriesCollection.Add Source:=cornerCell.offset(0, 4 * p + 3).Resize(numRows, 1)
                .SeriesCollection(numSeries).name = pop.name
                .SeriesCollection(numSeries).XValues = cornerCell.offset(1, 0).Resize(numRows - 1, 1)
                .ApplyDataLabels Type:=xlDataLabelsShowValue
                Set errorRng = cornerCell.offset(1, 4 * p + 4).Resize(numRows - 1, 1)
                .SeriesCollection(numSeries).ErrorBar Direction:=xlX, include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeCustom, _
                    Amount:="='" & cornerCell.Worksheet.name & "'!" & errorRng.Address, _
                    minusvalues:="='" & cornerCell.Worksheet.name & "'!" & errorRng.Address
            End If
        Next p

        'Format the Chart
        .HasAxis(xlCategory) = True
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.offset = 500
            .ReversePlotOrder = True
            .MajorTickMark = xlTickMarkNone
            .TickLabels.Font.Color = vbBlack
            .TickLabels.Font.Bold = True
            .TickLabels.Font.Size = 10
            .Format.Line.ForeColor.RGB = vbBlack
            .Format.Line.Weight = 3
            .HasTitle = False
        End With
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .HasMajorGridlines = False
            .TickLabelPosition = xlTickMarkNone
            .Format.Line.Visible = False
        End With
        .HasTitle = True
        .HasTitle = False   'I have ABSOLUTELY no idea why this f*cking toggle is necessary, but a runtime error occurs without it
        .HasTitle = True
        With .ChartTitle
            .Text = "Percent Change in RGC Firing Properties"
            .Font.Color = vbBlack
            .Font.Bold = True
            .Font.Size = 18
        End With
        If numSeries > 1 Then
            .HasLegend = True
            With .Legend
                .Font.Color = vbBlack
                .Font.Bold = True
                .Font.Size = 10
                .Position = xlRight
            End With
        Else
            .HasLegend = False
        End If

        'Formats the Chart's Series
        numSeries = 0
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            If pop.ID <> CTRL_POP.ID Then
                numSeries = numSeries + 1
                With .FullSeriesCollection(numSeries)
                    .Format.Fill.ForeColor.RGB = pop.BackColor
                    .Format.Line.ForeColor.RGB = vbBlack
                    .Format.Line.Weight = 2
                    .ErrorBars.Format.Line.ForeColor.RGB = vbBlack
                    .ErrorBars.Format.Line.Weight = 2
                    .HasLeaderLines = False
                    .DataLabels.NumberFormat = "0.0%"
                    .DataLabels.Position = xlLabelPositionOutsideEnd
                    .DataLabels.Font.Size = 10
                    .DataLabels.Font.Bold = True
                    .DataLabels.Font.Color = vbBlack
                End With
            End If
        Next p

    End With

    'Build property areas
    Dim prop As Integer
    Dim propCornerCell As Range, tblRowRng As Range
    rOffset = numPropRows + 2 + numChartRows + 1
    For prop = 1 To NUM_BKGRD_PROPERTIES
        Set tblRowRng = cornerCell.offset(prop, 0)
        cOffset = numCols + (prop - 1) * numPropCols
        Set propCornerCell = cornerCell.offset(rOffset, cOffset)
        Set chartRng = propCornerCell.offset(-(numChartRows + 1), 0).Resize(numChartRows, 1 + 2 * POPULATIONS.Count)
        Call buildPropArea(propCornerCell, tblRowRng, chartRng, PROP_UNITS(prop), "Burst", maxTissues)
    Next prop
    For t = 1 To numBurstTypes
        rOffset = (t + 1) * numPropRows + 2 + numChartRows + 1
        For prop = 1 To NUM_BURST_PROPERTIES
            Set tblRowRng = cornerCell.offset(NUM_BKGRD_PROPERTIES + (t - 1) * NUM_BURST_PROPERTIES + prop, 0)
            cOffset = numCols + (prop - 1) * numPropCols
            Set propCornerCell = cornerCell.offset(rOffset, cOffset)
            Set chartRng = propCornerCell.offset(-(numChartRows + 1), 0).Resize(numChartRows, 1 + 2 * POPULATIONS.Count)
            Call buildPropArea(propCornerCell, tblRowRng, chartRng, PROP_UNITS(NUM_BKGRD_PROPERTIES + prop), BURST_TYPES(1, t), maxTissues)
        Next prop
    Next t

    'Final formatting...
    Columns.AutoFit
    Rows.AutoFit
    cornerCell.offset(-1, 0).Resize(1, numRows - 1).value = chartTitles
    Rows(cornerCell.row - 1).Hidden = True
    Range(Columns(2), Columns(4 * POPULATIONS.Count + 1)).Hidden = True

End Sub

Private Sub buildSttcFiguresSheet()

    'Define some boilerplate variables
    Dim numBurstTypes As Integer, p As Integer, t As Integer
    Dim numCols As Integer, numRows As Integer, col As Integer, row As Integer
    Dim tbl As ListObject
    numBurstTypes = UBound(BURST_TYPES, 2)

    'Build the Figures sheet
    ActiveSheet.name = STTCFIGS_SHT_NAME

    'Store named distances
    Dim cornerCell As Range
    Set cornerCell = Cells(1, 1)
    Dim distVals As Variant
    ReDim distVals(1 To 4, 1 To 1)
    distVals(1, 1) = "Inter-Electrode Distance (µm)"
    distVals(2, 1) = "200"
    distVals(3, 1) = "Ignore Cutoff Distance (µm)"
    distVals(4, 1) = "800"
    cornerCell.Resize(4, 1).value = distVals
    cornerCell.offset(1, 0).name = "InterElectrodeDist"
    cornerCell.offset(3, 0).name = "IgnoreDist"
        
    'Create the contents table on the Contents sheet
    Dim numChartRows As Integer, numChartCols As Integer, numSpaceRows As Integer
    numChartRows = 20
    numChartCols = 10
    numSpaceRows = 5
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, cornerCell.offset(numChartRows + numSpaceRows, 0), , xlYes)
    tbl.name = STTC_TBL_NAME
    
    'Add its columns
    Dim pop As Population, tiss As Tissue
    numCols = 2
    tbl.ListColumns.Add
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items()(p)
        tbl.ListColumns.Add     'For mean
        tbl.ListColumns.Add     'For SEM
        numCols = numCols + 2
        For t = 0 To pop.Tissues.Count - 1
            tbl.ListColumns.Add
            numCols = numCols + 1
        Next t
    Next p
    ReDim headers(1 To 1, 1 To numCols)
    headers(1, 1) = "Unit Distance"
    headers(1, 2) = "Real Distance"
    col = 2
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items()(p)
        headers(1, col + 1) = pop.Abbreviation & "_Avg"
        headers(1, col + 2) = pop.Abbreviation & "_SEM"
        col = col + 2
        For t = 0 To pop.Tissues.Count - 1
            Set tiss = pop.Tissues.Item(t + 1)
            headers(1, col + 1) = pop.Abbreviation & "_" & tiss.ID
            col = col + 1
        Next t
    Next p
    tbl.HeaderRowRange.value = headers
    
    'Add inter-electrode distances
    numRows = NUM_CHANNELS * (NUM_CHANNELS - 1) / 2
    row = 0
    Dim ch1 As Integer, ch2 As Integer
    For ch1 = 0 To NUM_CHANNELS - 1
        For ch2 = ch1 + 1 To NUM_CHANNELS - 1
            tbl.ListRows.Add
            row = row + 1
            tbl.ListRows(row).Range(1, 1).value = interElectrodeDistance(ch1, ch2)
        Next ch2
    Next ch1
    
    'Remove duplicate inter-electrode distances and sort
    tbl.ListColumns(1).DataBodyRange.RemoveDuplicates Columns:=1, header:=xlYes
    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(1).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With tbl.Sort
        .header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'Add formulas to the remaining rows/columns
    Dim firstTiss As Tissue, lastTiss As Tissue, popRngStr As String, popSttcTblStr As String
    numCols = 2
    tbl.ListColumns(2).DataBodyRange.Formula = "=IF(InterElectrodeDist*[@Unit Distance]<=IgnoreDist,InterElectrodeDist*[@[Unit Distance]],NA())"
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items()(p)
        Set firstTiss = pop.Tissues(1)
        Set lastTiss = pop.Tissues(pop.Tissues.Count)
        popRngStr = STTC_TBL_NAME & "[@[" & pop.Abbreviation & "_" & firstTiss.ID & "]:[" & pop.Abbreviation & "_" & lastTiss.ID & "]]"
        tbl.ListColumns(numCols + 1).DataBodyRange.Formula = "=AVERAGE(" & popRngStr & ")"
        tbl.ListColumns(numCols + 2).DataBodyRange.Formula = "=STDEV.S(" & popRngStr & ")/SQRT(COUNT(" & popRngStr & "))"
        For t = 1 To pop.Tissues.Count
            Set tiss = pop.Tissues.Item(t)
            popSttcTblStr = pop.name & "_STTC"
            tbl.ListColumns(numCols + 2 + t).DataBodyRange.Formula = "=AVERAGEIFS(" & popSttcTblStr & "[STTC]," & popSttcTblStr & "[Tissue],""" & tiss.ID & """," & popSttcTblStr & "[Unit Distance],[@Unit Distance])"
        Next t
        numCols = numCols + 2 + pop.Tissues.Count
    Next p
    
    'Format cells
    With cornerCell.Resize(4, 1)
        .Font.Size = 11
        .Font.Bold = True
    End With
    cornerCell.offset(0, 0).HorizontalAlignment = xlLeft
    cornerCell.offset(1, 0).HorizontalAlignment = xlCenter
    cornerCell.offset(2, 0).HorizontalAlignment = xlLeft
    cornerCell.offset(3, 0).HorizontalAlignment = xlCenter
    tbl.HeaderRowRange.HorizontalAlignment = xlLeft
    tbl.DataBodyRange.HorizontalAlignment = xlCenter
    tbl.DataBodyRange.NumberFormat = "0.000"
    numCols = 3
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items(p)
        tbl.ListColumns(numCols).Range.Borders(xlEdgeLeft).Weight = xlMedium
        tbl.ListColumns(numCols + 2).Range.Borders(xlEdgeLeft).Weight = xlThin
        numCols = numCols + pop.Tissues.Count + 2
    Next p

    'Add the new column chart object for percent changes
    Dim numPropRows As Integer, numPropCols
    numPropRows = 10
    Dim chartShp As Shape, chartRng As Range
    Set chartRng = cornerCell.offset(0, 2).Resize(numChartRows, numChartCols)
    Set chartShp = ActiveSheet.Shapes.AddChart(xlXYScatter, chartRng.Left, chartRng.Top, chartRng.Width, chartRng.Height)
    With chartShp
        .name = "STTC_Chart"
        .Line.Visible = False
    End With

    With chartShp.Chart

        'Remove default Series
        Dim s As Integer
        For s = 1 To .SeriesCollection.Count
            .SeriesCollection(1).Delete
        Next s

        'Set the new population Series (showing future hidden cells too)
        Dim errorRng As Range, numSeries As Integer
        numSeries = 0
        numCols = 2
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            numSeries = numSeries + 1
            .SeriesCollection.Add Source:=tbl.ListColumns(numCols + 1).DataBodyRange
            .SeriesCollection(numSeries).name = pop.name
            .SeriesCollection(numSeries).XValues = tbl.ListColumns(2).DataBodyRange
            Set errorRng = tbl.ListColumns(numCols + 2).DataBodyRange
            .SeriesCollection(numSeries).ErrorBar Direction:=xlY, include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeCustom, _
                Amount:="='" & cornerCell.Worksheet.name & "'!" & errorRng.Address, _
                minusvalues:="='" & cornerCell.Worksheet.name & "'!" & errorRng.Address
            numCols = numCols + 2 + pop.Tissues.Count
        Next p

        'Format the Chart
        .HasAxis(xlCategory) = True
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.Font.Color = vbBlack
            .TickLabels.Font.Bold = True
            .TickLabels.Font.Size = 10
            .TickLabels.NumberFormat = "0"
            .Format.Line.ForeColor.RGB = vbBlack
            .Format.Line.Weight = 3
            .HasTitle = True
            .AxisTitle.Caption = "Distance (µm)"
            .AxisTitle.Font.Color = vbBlack
            .AxisTitle.Font.Bold = True
            .AxisTitle.Font.Size = 12
        End With
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .HasMajorGridlines = False
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabels.Font.Color = vbBlack
            .TickLabels.Font.Bold = True
            .TickLabels.Font.Size = 10
            .TickLabels.NumberFormat = "0.0"
            .Format.Line.ForeColor.RGB = vbBlack
            .Format.Line.Weight = 3
            .HasTitle = True
            .AxisTitle.Caption = "STTC"
            .AxisTitle.Font.Color = vbBlack
            .AxisTitle.Font.Bold = True
            .AxisTitle.Font.Size = 12
        End With
        .HasTitle = True
        .HasTitle = False   'I have ABSOLUTELY no idea why this f*cking toggle is necessary, but a runtime error occurs without it
        .HasTitle = True
        With .ChartTitle
            .Text = "Spike Time Tiling Coefficients vs. Inter-Electrode Distance"
            .Font.Color = vbBlack
            .Font.Bold = True
            .Font.Size = 18
        End With
        .HasLegend = True
        With .Legend
            .Font.Color = vbBlack
            .Font.Bold = True
            .Font.Size = 12
            .Position = xlTop
        End With

        'Formats the Chart's Series
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            With .FullSeriesCollection(p + 1)
                .Trendlines.Add Type:=xlPolynomial, Order:=3, name:=""
                .Trendlines(1).Format.Line.DashStyle = msoLineDash
                .Trendlines(1).Format.Line.ForeColor.RGB = pop.BackColor
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
                .MarkerBackgroundColor = pop.BackColor
                .MarkerForegroundColor = vbBlack
                .Format.Line.Weight = 1.5
                .ErrorBars.Format.Line.ForeColor.RGB = vbBlack
                .ErrorBars.Format.Line.Weight = 2
            End With
            .Legend.LegendEntries(POPULATIONS.Count + 1).Delete
        Next p
        
    End With

    'Final formatting...
    Columns.AutoFit
    Rows.AutoFit

End Sub

Private Sub buildPropArea(ByRef cornerCell As Range, ByRef tblRowCell As Range, ByRef chartRng As Range, ByVal unitsStr As String, ByVal bType As String, ByVal maxTissues As Integer)

    Dim numBurstTypes As Integer, numHeaders As Integer
    Dim t As Integer, pop As Population, p As Integer
    Dim rOffset As Integer, cOffset As Integer
    numBurstTypes = UBound(BURST_TYPES, 2)
    numHeaders = 1 + 2 * numBurstTypes
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
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items()(p)
        cOffset = 1 + p * numPopCols
        headers(1, cOffset + 1) = pop.Abbreviation & "_Avg"
        headers(1, cOffset + 2) = pop.Abbreviation & "_%Change"
        cornerCell.offset(2, 2 * p + 1).Resize(1, 2).Interior.Color = pop.BackColor
        cornerCell.offset(2, 2 * p + 1).Resize(1, 2).Font.Color = pop.ForeColor
    Next p
    cornerCell.offset(2, 0).Resize(1, numHeaders).value = headers
    cornerCell.offset(2, 0).Resize(1, numHeaders).Font.Bold = True
    
    'Identify the control population's data range
    Dim ctrlRng As Range
    For p = 0 To POPULATIONS.Count - 1
        Set pop = POPULATIONS.items()(p)
        If pop.ID = CTRL_POP.ID Then
            cOffset = numPopCols * p
            Set ctrlRng = cornerCell.offset(3, cOffset + 1).Resize(maxTissues, 1)
            Exit For
        End If
    Next p
    
    'Write tissue data (formula for percent change depends on whether data is paired)
    Dim tissueCell As Range, ctrlValueStr As String
    For t = 1 To maxTissues
        cornerCell.offset(2 + t, 0).value = t
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            Set tissueCell = cornerCell.offset(2 + t, p * numPopCols + 1)
            ctrlValueStr = IIf(dataPaired, ctrlRng.Cells(t, 1).Address, "AVERAGE(" & ctrlRng.Address & ")")
            tissueCell.Formula = "=IFERROR(AVERAGEIF(" & pop.name & "_" & bType & "s[Tissue]," & cornerCell.offset(2 + t, 0).Address & "," & pop.name & "_" & bType & "s[" & tblRowCell.value & "]), """")"
            tissueCell.offset(0, 1).Formula = "=IFERROR((" & tissueCell.Address & "-" & ctrlValueStr & ")/" & ctrlValueStr & ", """")"
        Next p
    Next t
    cornerCell.offset(3, 0).Resize(maxTissues, 1).Font.Bold = True
    
    'Add formulas to the main table
    Dim formulaRng As Range, propRng As Range, summaryRng As Range
    For p = 0 To POPULATIONS.Count - 1
        Set formulaRng = tblRowCell.offset(0, 4 * p)
        Set propRng = cornerCell.offset(3, 2 * p).Resize(maxTissues, 1)
        formulaRng.offset(0, 1).Formula = "=IFERROR(AVERAGE(" & propRng.offset(0, 1).Address & "), """")"
        formulaRng.offset(0, 2).Formula = "=IFERROR(STDEV.S(" & propRng.offset(0, 1).Address & ")/SQRT(COUNT(" & propRng.offset(0, 1).Address & ")), """")"
        formulaRng.offset(0, 3).Formula = "=IFERROR(AVERAGE(" & propRng.offset(0, 2).Address & "), """")"
        formulaRng.offset(0, 4).Formula = "=IFERROR(STDEV.S(" & propRng.offset(0, 2).Address & ")/SQRT(COUNT(" & propRng.offset(0, 2).Address & ")), """")"
        Set summaryRng = tblRowCell.offset(0, 4 * POPULATIONS.Count + 3 * p)
        summaryRng.offset(0, 1).Formula = "=TEXT(" & formulaRng.offset(0, 1).Address & ",""0.000"")&"" ± ""&TEXT(" & formulaRng.offset(0, 2).Address & ",""0.000"")"
        summaryRng.offset(0, 2).Formula = "=TEXT(" & formulaRng.offset(0, 3).Address & ",""0.0"")&"" ± ""&TEXT(" & formulaRng.offset(0, 4).Address & ",""0.0"")"
        summaryRng.offset(0, 3).Formula = "=IFERROR(TEXT(T.TEST(" & cornerCell.offset(3, 2 * p + 1).Resize(maxTissues, 1).Address & "," & ctrlRng.Address & ",TTTails,TTType), ""0.000""), """")"
    Next p
    
    'Add the new bar chart object
    Dim chartShp As Shape
    Set chartShp = ActiveSheet.Shapes.AddChart(xlColumnClustered, chartRng.Left, chartRng.Top, chartRng.Width, chartRng.Height)
    With chartShp
        .name = Replace(tblRowCell.value, " ", "_") & "_Chart"
        .Line.Visible = False
    End With
    
    With chartShp.Chart
    
        'Clear its default Series
        Dim s As Integer
        For s = 1 To .SeriesCollection.Count
            .SeriesCollection(1).Delete
        Next s
        
        'Set the new population Series (showing future hidden cells too)
        .PlotVisibleOnly = False
        Dim errorRng As Range
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            .SeriesCollection.Add Source:=tblRowCell.offset(0, 4 * p + 1)
            .SeriesCollection(p + 1).name = pop.name
            Set errorRng = tblRowCell.offset(0, 4 * p + 2)
            .SeriesCollection(p + 1).ErrorBar Direction:=xlY, include:=xlErrorBarIncludeBoth, Type:=xlErrorBarTypeCustom, _
                Amount:="='" & tblRowCell.Worksheet.name & "'!" & errorRng.Address, _
                minusvalues:="='" & tblRowCell.Worksheet.name & "'!" & errorRng.Address
        Next p
    
        'Format the Chart
        .HasAxis(xlCategory) = True
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionNone
            .MajorTickMark = xlTickMarkNone
            .Format.Line.ForeColor.RGB = vbBlack
            .Format.Line.Weight = 3
            .HasTitle = False
        End With
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .HasMajorGridlines = False
            .TickLabels.Font.Color = vbBlack
            .TickLabels.Font.Bold = True
            .TickLabels.Font.Size = 10
            .Format.Line.ForeColor.RGB = vbBlack
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
        With .ChartTitle
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
        For p = 0 To POPULATIONS.Count - 1
            Set pop = POPULATIONS.items()(p)
            With .FullSeriesCollection(p + 1)
                .Format.Fill.ForeColor.RGB = pop.BackColor
                .Format.Line.ForeColor.RGB = vbBlack
                .Format.Line.Weight = 2
                .ErrorBars.Format.Line.ForeColor.RGB = vbBlack
                .ErrorBars.Format.Line.Weight = 2
            End With
        Next p
        
    End With

End Sub

Private Sub combineIntoWb(ByRef numGoodTissues As Integer)
    'For each tissue, open its workbooks, fetch their data, and re-close the workbooks
    Dim popV As Variant, pop As Population, t As Tissue
    For Each popV In POPULATIONS.items
        Set pop = popV
        For Each t In pop.Tissues
            Call fetchTissue(t, numGoodTissues)
        Next t
    Next popV
    
    'Pretty up the sheets now that data is present
    Call cleanSheets(combineWb, CONTENTS_SHT_NAME)
    Call cleanSheets(combineWb, "_STTC")
    Call cleanSheets(combineWb, "_Bursts")
    Call cleanSheets(combineWb, "_WABs")
    Call cleanSheets(combineWb, "_NonWABs")

    'Save/close the workbook if the user doesn't want to keep it open
    If Not keepOpen Then _
        Call combineWb.Close(True)

End Sub

Private Sub fetchTissue(ByRef Tissue As Tissue, ByRef numGoodTissues As Integer)
    'Make sure an ID was provided for this tissue
    Dim result As VbMsgBoxResult
    If Tissue.ID = 0 Then
        result = MsgBox("A tissue in population " & Tissue.Population.name & " was not given an ID." & vbCr & _
                        "Its data will not be loaded.")
        Exit Sub
    End If
    
    'If so, then Initialize some local variables
    Dim fs As New FileSystemObject
    Dim tissueWb As Workbook
    Dim numBurstTypes As Integer
    numBurstTypes = UBound(BURST_TYPES, 2)
    
    'For each type of data...
    Dim wbFound As Boolean, wbTypeName As Variant, wbPath As String
    For Each wbTypeName In BURST_TYPES
        'Check that a workbook was provided and exists (display error dialogs if not)
        wbFound = False
        wbPath = Tissue.WorkbookPaths(wbTypeName)
        If fs.FileExists(wbPath) Then
            wbFound = True
            numGoodTissues = numGoodTissues + 1
        ElseIf wbPath = "" Then
            result = MsgBox("No " & wbTypeName & " workbook provided for tissue " & Tissue.ID & " in population " & Tissue.Population.name & ".", vbOKOnly)
        Else
            result = MsgBox("Workbook """ & wbPath & """ could not be found." & vbCr & _
                            "Make sure you provided the correct path to the " & wbTypeName & " workbook.", vbOKOnly)
        End If
        
        'Load the tissue's data and store it in the appropriate sheets of the combination workbook
        If wbFound Then
            Dim popName As String
            Set tissueWb = Workbooks.Open(wbPath)
            popName = Tissue.Population.name
            Select Case wbTypeName
                Case "Processed WAB Workbook"
                    Call copyTissueData(tissueWb, STTC_SHT_NAME, popName & "_STTC", Tissue.ID)
                    Call copyTissueData(tissueWb, ALL_AVGS_SHT_NAME, popName & "_Bursts", Tissue.ID)
                    Call copyTissueData(tissueWb, BURST_AVGS_SHT_NAME, popName & "_WABs", Tissue.ID)
                    
                Case "Processed NonWAB Workbook"
                    Call copyTissueData(tissueWb, BURST_AVGS_SHT_NAME, popName & "_NonWABs", Tissue.ID)
            End Select
            tissueWb.Close
        End If
    Next wbTypeName

End Sub

Private Sub copyTissueData(ByRef tissueWb As Workbook, ByVal fetchName As String, ByVal outputName As String, ByVal tissueID As String)
    'Set the Range of data to be copied from the tissue workbook
    Dim fetchRng As Range
    Set fetchRng = tissueWb.Worksheets(fetchName).ListObjects(fetchName).DataBodyRange
    
    'Set the Range to be copied to in the summary workbook
    Dim outputTbl As ListObject, outputRng
    Set outputTbl = combineWb.Worksheets(outputName).ListObjects(outputName)
    outputTbl.ListRows.Add
    Set outputRng = outputTbl.ListRows(outputTbl.ListRows.Count).Range.Cells(1, 2)
    
    'Copy the data, and add the provided tissueID to each row
    Dim idRng As Range
    fetchRng.Copy Destination:=outputRng
    Set idRng = outputRng.offset(0, -1).Resize(fetchRng.Rows.Count, 1)
    idRng.value = tissueID

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

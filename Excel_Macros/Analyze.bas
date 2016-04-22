Attribute VB_Name = "Analyze"
Option Private Module
Option Explicit

'Processing variables for this module
Private maxBursts As Integer
Private unitNames As Variant

Public Sub processRetinalWorkbook(ByVal wbName As String, ByRef numRecordings As Integer)
    Dim rec As Integer, u As Integer
    
    'Make sure the property strings and constants are initialized
    Call setupOptimizations
    
    Call GetConfigVars
    
    'Get some additional config settings from the user
    Dim analyzeSht As Worksheet, allChecked As Boolean, wabsChecked As Boolean
    Set analyzeSht = Worksheets(ANALYZE_SHT_NAME)
    allChecked = (analyzeSht.Shapes("AllRadio").OLEFormat.Object.value = 1)
    wabsChecked = (analyzeSht.Shapes("WabsRadio").OLEFormat.Object.value = 1)
    BURSTS_TO_USE = IIf(allChecked, BurstUseType.All, IIf(wabsChecked, BurstUseType.WABs, BurstUseType.NonWABs))
    ASSOC_SAME_CHANNEL_UNITS = (analyzeSht.Shapes("SameChannelAssocChk").OLEFormat.Object.value = 1)
    
    'If there are no recordings in this workbook then just return
    Dim wb As Workbook
    Set wb = Workbooks.Open(wbName)
    Dim summaryTbl As ListObject
    Set summaryTbl = wb.Worksheets(CONTENTS_SHT_NAME).ListObjects(SUMMARY_TBL_NAME)
    numRecordings = summaryTbl.ListRows.Count
    If numRecordings = 0 Then _
        GoTo ExitSub
            
    'Get the names of all units on the first sheet (assumed to be same on all other recording sheets)
    Dim wksht As Worksheet
    Set wksht = wb.Worksheets(summaryTbl.ListRows(1).Range(1, 2).value)
    Dim numUnits As Integer
    numUnits = wksht.Cells(1, 1).End(xlToRight).Column / 3    'Since every unit is mentioned once for spikes, burst_start, and burst_end
    unitNames = Application.Transpose(wksht.Cells(1, 1).Resize(1, numUnits))
    
    'Add output sheets
    Call addAllAvgsSheet(wb.name, unitNames)
    Call addBurstAvgsSheet(wb.name, unitNames)
    Call addSttcSheet(wb.name, unitNames)
    
    'Process each recording for this retina (represented as separate sheets)
    Dim recRow As ListRow, recName As String
    Dim startT As Double, endT As Double
    For Each recRow In summaryTbl.ListRows
        recName = recRow.Range(1, 2)
        startT = recRow.Range(1, 3)
        endT = recRow.Range(1, 4)
        wb.Worksheets(recName).Activate
        Call processRecording(unitNames, startT, endT)
    Next recRow

    'Reduce result sums to averages and finalize
    Call reduceToAvgs(numRecordings, numUnits)
    Call finalize(numUnits)
    
ExitSub:
    Call tearDownOptimizations
End Sub

Private Sub addAllAvgsSheet(ByVal retinaName As String, ByRef unitNames As Variant)
    Dim zeroes() As Variant
    Dim u, p As Integer
    
    Dim numUnits As Integer
    numUnits = UBound(unitNames, 1)
    
    'Add the Averages sheet, write out row (unit) headers, and add formatting
    Dim avgsRng As Range
    Worksheets.Add After:=Sheets(CONTENTS_SHT_NAME)
    ActiveSheet.name = ALL_AVGS_SHT_NAME
    Set avgsRng = Worksheets(ALL_AVGS_SHT_NAME).Cells(2, 2)
    Cells(1, 1).value = CELL_STR
    avgsRng.offset(0, -1).Resize(numUnits, 1).value = unitNames
    avgsRng.offset(-1, 0).Resize(1, NUM_BKGRD_PROPERTIES).value = PROPERTIES
    avgsRng.offset(-1, -1).Resize(numUnits + 1).Font.Bold = True
    
    'Write out column (property) headers, indicating which properties are being skipped with a "*"
    Dim pStr As String
    For p = 1 To NUM_BKGRD_PROPERTIES
        pStr = PROPERTIES(p)
        avgsRng.offset(-1, p - 1).value = pStr
    Next p
    
    'Initialize all output cells to zero
    ReDim zeroes(1 To numUnits, 1 To NUM_BKGRD_PROPERTIES)
    For u = 1 To numUnits
        For p = 1 To NUM_BKGRD_PROPERTIES
            zeroes(u, p) = 0
        Next p
    Next u
    avgsRng.Resize(numUnits, NUM_BKGRD_PROPERTIES).value = zeroes
        
    'Make a table for all the average values
    Worksheets(ALL_AVGS_SHT_NAME).ListObjects.Add( _
        xlSrcRange, _
        avgsRng.CurrentRegion, , _
        xlYes) _
    .name = ALL_AVGS_SHT_NAME
    avgsRng.Resize(numUnits, NUM_BKGRD_PROPERTIES).NumberFormat = "0.000"
End Sub

Private Sub addBurstAvgsSheet(ByVal retinaName As String, ByRef unitNames As Variant)
    Dim zeroes() As Variant
    Dim u, p As Integer
    
    Dim numUnits As Integer
    numUnits = UBound(unitNames, 1)
    
    'Add the Averages sheet, write out row (unit) headers, and add formatting
    Dim avgsRng As Range
    Worksheets.Add After:=Sheets(ALL_AVGS_SHT_NAME)
    ActiveSheet.name = BURST_AVGS_SHT_NAME
    Set avgsRng = Worksheets(BURST_AVGS_SHT_NAME).Cells(2, 2)
    Cells(1, 1).value = CELL_STR
    avgsRng.offset(0, -1).Resize(numUnits, 1).value = unitNames
    avgsRng.offset(-1, -1).Resize(numUnits + 1).Font.Bold = True
    
    'Write out column (property) headers, indicating which properties are being skipped with a "*"
    Dim pStr As String
    For p = 1 To NUM_BURST_PROPERTIES
        pStr = PROPERTIES(NUM_BKGRD_PROPERTIES + p)
        avgsRng.offset(-1, p - 1).value = pStr
    Next p
    
    'Initialize all output cells to zero
    ReDim zeroes(1 To numUnits, 1 To NUM_BURST_PROPERTIES)
    For u = 1 To numUnits
        For p = 1 To NUM_BURST_PROPERTIES
            zeroes(u, p) = 0
        Next p
    Next u
    avgsRng.Resize(numUnits, NUM_BURST_PROPERTIES).value = zeroes
        
    'Make a table for all the average values
    Worksheets(BURST_AVGS_SHT_NAME).ListObjects.Add( _
        xlSrcRange, _
        avgsRng.CurrentRegion, , _
        xlYes) _
    .name = BURST_AVGS_SHT_NAME
    avgsRng.Resize(numUnits, NUM_BURST_PROPERTIES).NumberFormat = "0.000"
End Sub

Private Sub addSttcSheet(ByVal retinaName As String, ByRef unitNames As Variant)
    Dim initial() As Variant
    Dim u, p As Integer
    
    Dim numUnits As Long
    Dim numRows As Long
    numUnits = UBound(unitNames, 1)
    numRows = numUnits * (numUnits - 1) / 2
    
    'Add the STTC sheet, write out row/column (channel) headers, and add formatting
    Dim sttcRng As Range
    Worksheets.Add After:=Sheets(BURST_AVGS_SHT_NAME)
    ActiveSheet.name = STTC_SHT_NAME
    Set sttcRng = Worksheets(STTC_SHT_NAME).Cells(4, 1)
    With sttcRng
        .offset(-3, 0).value = STTC_HEADER_STR
        .offset(-3, 0).Font.Bold = True
        .offset(-3, 0).Font.Size = 16
        .offset(-1, 0).value = "Cell1"
        .offset(-1, 1).value = "Cell2"
        .offset(-1, 2).value = "Unit Distance"
        .offset(-1, 3).value = "STTC"
        .offset(-1, 0).EntireRow.Font.Bold = True
    End With
    
    'Initialize all output cells to zero
    ReDim initial(1 To numRows, 1 To 3)
    Dim u1, u2, ch1, ch2 As Integer
    Dim row As Long
    row = 1
    For u1 = 1 To numUnits
        For u2 = u1 + 1 To numUnits
            ch1 = channelIndex(unitNames(u1, 1))
            ch2 = channelIndex(unitNames(u2, 1))
            initial(row, 1) = unitNames(u1, 1)
            initial(row, 2) = unitNames(u2, 1)
            initial(row, 3) = interElectrodeDistance(ch1, ch2)
            row = row + 1
        Next u2
    Next u1
    sttcRng.Resize(numRows, 3).value = initial
        
    'Make a table for all the STTC values
    Worksheets(STTC_SHT_NAME).ListObjects.Add( _
        xlSrcRange, _
        sttcRng.CurrentRegion, , _
        xlYes) _
    .name = STTC_SHT_NAME
    sttcRng.offset(0, 2).Resize(numRows, 2).NumberFormat = "0.000"
    
End Sub

Private Sub processRecording(ByRef unitNames As Variant, ByVal startTime As Double, ByVal endtime As Double)
    Dim u As Integer
    Dim spikes As Variant, preBursts As Variant, postBursts As Variant
        
    'Keep the below For loops separated so that
    'spikes on one channel arent correlated then deleted before pairing with other channels and so that
    'bursts don't get associated with bursts that would later get deleted
    
    'Store STTC values using the entire spike trains of every possible pair of channels
    Dim numUnits As Integer
    numUnits = UBound(unitNames)
    Dim duration As Double
    duration = endtime - startTime
    Call storeSttcValues(duration, numUnits)
        
    'Remove bursts that start/end too late/early (lolwut?)
    'Adjust the start/end times of each unit, if necessary
    Dim startEndTimes() As Double
    ReDim startEndTimes(1 To numUnits, 1 To 2)
    For u = 1 To numUnits
        spikes = getSpikeTrain(u)
        preBursts = getBurstTrain(u, numUnits)
        Call deleteBadBurstsFrom(u, numUnits, spikes, preBursts, startTime, endtime, startEndTimes)
        Call deleteBadSpikesFrom(u, spikes, startTime, endtime, startEndTimes(u, 1), startEndTimes(u, 2))
    Next u
    ActiveSheet.UsedRange   'Refresh used range by getting the property
    
    'Do PRE ANALYSES on each unit
    'I.e., analyses BEFORE unused bursts are excluded (background firing metrics)
    Dim preBurstCounts() As Integer
    ReDim preBurstCounts(1 To numUnits)
    For u = 1 To numUnits
        spikes = getSpikeTrain(u)
        preBursts = getBurstTrain(u, numUnits)
        preBurstCounts(u) = UBound(preBursts, 1)
        duration = startEndTimes(u, 2) - startEndTimes(u, 1)
        Call storePreValues(u, spikes, preBursts, duration)
    Next u
    
    'Exclude unused bursts (WABs or non-WABs), if requested
    If BURSTS_TO_USE <> BurstUseType.All Then
        Dim wabsOnly As Boolean
        wabsOnly = (BURSTS_TO_USE = BurstUseType.WABs)
        Call deleteUnusedBursts(wabsOnly, ASSOC_SAME_CHANNEL_UNITS, unitNames)
    End If
        
    'Do POST ANALYSES on each unit
    'I.e., analyses AFTER unused bursts are excluded (wave- or non-wave-associated firing metrics)
    Dim wabRatio As Double
    Dim postBurstCounts() As Integer
    ReDim postBurstCounts(1 To numUnits)
    For u = 1 To numUnits
        spikes = getSpikeTrain(u)
        postBursts = getBurstTrain(u, numUnits)
        postBurstCounts(u) = UBound(postBursts, 1)
        duration = startEndTimes(u, 2) - startEndTimes(u, 1)
        wabRatio = postBurstCounts(u) / preBurstCounts(u)
        Call storePostValues(u, spikes, postBursts, duration, wabRatio)
    Next u
    
    'Reformat the datasheet so its a little easier to view
    Cells.VerticalAlignment = xlCenter
    Cells.HorizontalAlignment = xlCenter
    Cells.EntireColumn.AutoFit
End Sub

Private Sub storeSttcValues(ByVal duration As Double, ByVal numUnits As Long)
    Dim tValues() As Double
    Dim cellIndex1, cellIndex2 As Integer
    Dim u1, u2 As Integer
    Dim sttc As Double
    Dim outputRng As Range
    Dim oldResults, newResults() As Variant
    Dim spikes1, spikes2 As Variant
    
    Dim numRows As Long
    numRows = numUnits * (numUnits - 1) / 2
            
    'Loop over each unit's spike timestamps (arranged in columns) to get T-values
    ReDim tValues(1 To numUnits)
    ReDim newResults(1 To numRows, 1 To 1)
    For u1 = 1 To numUnits
        'For each unit, store the fraction of the recording's duration wherein its spikes are delta-t apart
        spikes1 = getSpikeTrain(u1)
        tValues(u1) = CorrelatedTimeProportion(spikes1, duration)
    Next u1
    
    'Store STTC values using the entire spike trains of every possible pair of units
    Dim row As Long
    row = 1
    For u1 = 1 To numUnits
        spikes1 = getSpikeTrain(u1)
        For u2 = u1 + 1 To numUnits
            spikes2 = getSpikeTrain(u2)
            sttc = SpikeTimeTilingCoefficient2(spikes1, spikes2, tValues(u1), tValues(u2))
            newResults(row, 1) = sttc
            row = row + 1
        Next u2
    Next u1
        
    'Output spike time tiling coefficients for all pairs of units with this unit, adding new results to the old ones
    Set outputRng = Worksheets(STTC_SHT_NAME).Cells(4, 4).Resize(numRows, 1)
    oldResults = outputRng.value
    outputRng.value = Application.Pmt(0, -1, oldResults, newResults)   'Pmt() acts like Sum() but can accept two variants (of equal dimension)
End Sub

Private Sub deleteBadBurstsFrom(ByVal u As Integer, ByVal numUnits As Integer, ByRef spikes As Variant, ByRef bursts As Variant, ByVal recStart As Double, ByVal recEnd As Double, ByRef startEndTimes() As Double)
    Dim b As Integer
    
    'If there are no bursts then just return
    If bursts(1, 1) = -1 Then
        startEndTimes(u, 1) = recStart
        startEndTimes(u, 2) = recEnd
        Exit Sub
    End If
    
    'Store the valid bursts in a new Variant
    Dim activeRow As Integer, offset As Double
    Dim newStart As Double, newEnd As Double, newBursts() As Variant
    Dim bStart As Double, bEnd As Double, burstDur As Double, tooEarly As Boolean, tooLate As Boolean, validDur As Boolean
    newStart = recStart
    newEnd = recEnd
    activeRow = 0
    offset = MAX_DURATION / 2
    ReDim newBursts(1 To UBound(bursts), 1 To 2)
    For b = 1 To UBound(bursts)
        'Check if this burst is valid (not cut off by the recording and of a valid duration)
        bStart = bursts(b, 1)
        bEnd = bursts(b, 2)
        tooEarly = (bStart < recStart + offset)
        tooLate = (recEnd - offset < bEnd)
        burstDur = bEnd - bStart
        validDur = (MIN_DURATION <= burstDur And burstDur <= MAX_DURATION)
        
        'If it was too late/early, then adjust the start/end timestamps for this unit
        If tooEarly Then _
            newStart = bursts(b, 2)
        If tooLate Then _
            newEnd = bursts(b, 1)
        
        'Otherwise, store this burst in the new Variant
        If Not tooEarly And Not tooLate And validDur Then
            activeRow = activeRow + 1
            newBursts(activeRow, 1) = bStart
            newBursts(activeRow, 2) = bEnd
        End If
    Next b
    
    'Replace the old burst train with only those bursts that are valid
    Dim burstRng As Range
    Set burstRng = Cells(2, burstColumn(u, numUnits))
    burstRng.Resize(UBound(bursts), 2).Clear
    burstRng.Resize(UBound(newBursts), 2).value = newBursts
    
    'Set the new start and end times for this unit
    startEndTimes(u, 1) = newStart
    startEndTimes(u, 2) = newEnd
End Sub

Private Sub deleteBadSpikesFrom(ByVal u As Integer, ByRef spikes As Variant, ByVal recStart As Double, ByVal recEnd As Double, ByRef unitStart As Double, ByRef unitEnd As Double)
    Dim s As Integer
    
    'If there are no spikes then just return
    If spikes(1, 1) = -1 Then _
        Exit Sub
    
    'Store the valid bursts in a new Variant
    Dim activeRow As Integer
    Dim newSpikes() As Variant
    Dim tooEarly As Boolean, tooLate As Boolean
    ReDim newSpikes(1 To UBound(spikes), 1 To 2)
    For s = 1 To UBound(spikes)
        'Check if this burst is valid (not cut off by the recording and of a valid duration)
        tooEarly = (spikes(s, 1) < unitStart)
        tooLate = (unitEnd < spikes(s, 1))
        
        'Otherwise, store this burst in the new Variant
        If Not tooEarly And Not tooLate Then
            activeRow = activeRow + 1
            newSpikes(activeRow, 1) = spikes(s, 1)
        End If
    Next s
    
    'Replace the old spike train with only those spikes that are valid
    'Remove the first/last spikes if they represent the end/start of cutoff bursts
    Dim spikeRng As Range
    Set spikeRng = Cells(2, u)
    spikeRng.Resize(UBound(spikes)).Clear
    spikeRng.Resize(UBound(newSpikes)).value = newSpikes
    If unitEnd <> recEnd Then _
        spikeRng.offset(activeRow - 1, 0).Delete shift:=xlUp
    If unitStart <> recStart Then _
        spikeRng.Delete shift:=xlUp
        
End Sub

Private Sub deleteUnusedBursts(ByVal wabsOnly As Boolean, ByVal assocSameChannelUnits As Boolean, ByRef unitNames As Variant)
    Dim burstRng As Range
    Dim numBursts As Integer
    Dim bursts, trimmedBursts As Variant
    Dim burstPos As Integer
    Dim isWAB As Boolean
    Dim firstU, lastU, nFirstU, nLastU As Integer
    Dim neighbor As Variant
    Dim validNeighbors As Collection
    Dim u, ch, nCh, b, chPos As Integer
    Dim chStr, nChStr As String
    Dim numAssocUnits As Integer
    
    'Get all the burst start/end timestamps
    Dim numUnits As Integer
    numUnits = UBound(unitNames)
    maxBursts = getMaxBursts(numUnits)
    ReDim trimmedBursts(1 To maxBursts, 1 To 2 * numUnits)
    Set burstRng = Cells(2, numUnits + 1).Resize(maxBursts, numUnits * 2)
    bursts = burstRng.value
    
    'For each unit...
    For u = 1 To numUnits
        ch = channelIndex(unitNames(u, 1))
        chStr = CHANNEL_PREFIX & channelString(ch)
        chPos = 2 * u - 1
        burstPos = 0
        
        'Find the first and last unit on this same channel
        firstU = u
        lastU = u
        If Not assocSameChannelUnits Then
            Do While firstU > 0
                If InStr(1, unitNames(firstU, 1), chStr) = 0 Then _
                    Exit Do
                firstU = firstU - 1
            Loop
            firstU = firstU + 1
            Do While lastU <= numUnits
                If InStr(1, unitNames(lastU, 1), chStr) = 0 Then _
                    Exit Do
                lastU = lastU + 1
            Loop
            lastU = lastU - 1
        End If
                
        'Create a list of its valid neighbors (not itself or a channel off the MEA)
        Set validNeighbors = New Collection
        For neighbor = 0 To 8
            If neighborValid(ch, neighbor) Then _
                validNeighbors.Add (neighbor)
        Next neighbor
        
        'For each of this unit's bursts...
        For b = 1 To maxBursts
            If bursts(b, chPos) = "" Then _
                Exit For
            isWAB = False
            
            'See if this burst has the minimum number of bins associated with
            'any bin of any burst on a same-channel unit (if requested)
            numAssocUnits = 0
            If assocSameChannelUnits Then
                Call burstAssociatedWithUnits(firstU, lastU, bursts, chPos, b, numAssocUnits)
                isWAB = (numAssocUnits >= MIN_ASSOC_UNITS)
                If isWAB = wabsOnly Then Exit For
            End If
        
            'If not, then loop over each of the unit's valid neighbor channels
            For Each neighbor In validNeighbors
                Call neighborUnitsAssociatedWithBurst(u, neighbor, b, bursts, numAssocUnits)
            Next neighbor
            
            'If the burst was associated with bursts on the minimum number of neighbor units,
            'then add its start/end timestamps to the "trimmed" array
            If isWAB Then
                burstPos = burstPos + 1
                trimmedBursts(burstPos, chPos) = bursts(b, chPos)
                trimmedBursts(burstPos, chPos + 1) = bursts(b, chPos + 1)
            End If
        Next b
        
    Next u
    
    'Overwrite the old burst timestamps with the "trimmed" ones
    burstRng.Clear
    burstRng.value = trimmedBursts
End Sub

Private Sub neighborUnitsAssociatedWithBurst(ByVal unit As Integer, ByVal neighbor As Variant, ByVal b As Integer, ByRef bursts As Variant, ByRef numAssocUnits As Integer)
    Dim trimmedBursts As Variant
    Dim nFirstU, nLastU As Integer
    Dim nCh As Integer, nChStr As String
    
    Dim numUnits As Integer
    numUnits = UBound(unitNames)
    Dim ch As Integer, chStr As String, chPos As String
    ch = channelIndex(unitNames(unit, 1))
    chStr = CHANNEL_PREFIX & channelString(ch)
    chPos = 2 * unit - 1
    
    'Find the first and last unit on the neighbor channel (if they are represented on the sheet)
    Dim neighborAfter, inChannel As Boolean
    Dim tempU, endForU, step As Integer
    nFirstU = -1
    nLastU = -1
    nCh = neighborChannel(ch, CInt(neighbor))
    nChStr = CHANNEL_PREFIX & channelString(nCh)
    neighborAfter = (nCh > ch)
    endForU = IIf(neighborAfter, numUnits, 1)
    step = IIf(neighborAfter, 1, -1)
    inChannel = False
    For tempU = unit To endForU Step step
        If InStr(1, unitNames(tempU, 1), nChStr) > 0 Then
            inChannel = True
            If neighborAfter Then
                If nFirstU = -1 Then nFirstU = tempU
                nLastU = tempU
            Else
                If nLastU = -1 Then nLastU = tempU
                nFirstU = tempU
            End If
        Else
            If inChannel Then Exit For
        End If
    Next tempU
    
    'If this neighbor channel had units on the sheet,
    'See if this burst has the minimum number of bins associated with any bin of any burst on one of those units
    If nFirstU <> -1 And nLastU <> -1 Then _
        Call burstAssociatedWithUnits(nFirstU, nLastU, bursts, chPos, b, numAssocUnits)
End Sub

Private Function getMaxBursts(ByVal numUnits As Integer) As Integer
    Dim numBursts As Integer, maxBursts As Integer
    Dim u As Integer
    
    'Return the max number of bursts in any unit
    maxBursts = 0
    For u = 1 To numUnits
        numBursts = WorksheetFunction.Count(Columns(burstColumn(u, numUnits)))
        maxBursts = WorksheetFunction.Max(maxBursts, numBursts)
    Next u

    getMaxBursts = maxBursts
End Function

Private Sub burstAssociatedWithUnits(ByVal firstUnit As Integer, ByVal lastUnit As Integer, ByRef bursts As Variant, ByVal channelPos As Integer, ByVal burst As Integer, ByRef numAssocUnits As Integer)
    'Get the duration of a bin in this burst
    Dim start As Double, finish As Double, binDuration As Double
    start = bursts(burst, channelPos)
    finish = bursts(burst, channelPos + 1)
    binDuration = (finish - start) / NUM_BINS
    
    'Loop over each unit between the two provided units (inclusive)
    'Don't just associate this unit with itself though, of course
    Dim nU, nB, nChPos, nStart, nFinish, nBinDuration As Double
    Dim associated As Boolean
    For nU = firstUnit To lastUnit
        associated = False
        nChPos = 2 * nU - 1
        If nChPos <> channelPos Then
        
            'See if the provided burst has the minimum number of bins associated with any bin of any burst on this neighbor unit
            For nB = 1 To maxBursts
                If bursts(nB, nChPos) = "" Then _
                    Exit For
                nStart = bursts(nB, nChPos)
                If nStart > finish Then _
                    Exit For
                nFinish = bursts(nB, nChPos + 1)
                nBinDuration = (nFinish - nStart) / NUM_BINS
                associated = burstsAssociated(start, binDuration, nStart, nBinDuration)
                If associated Then
                    numAssocUnits = numAssocUnits + 1
                    Exit For
                End If
            Next nB
            
            'If the min number of associated units has been achieved, then exit the loop and return
            If numAssocUnits >= MIN_ASSOC_UNITS Then _
                Exit For
        End If
    Next nU
    
End Sub

Private Sub storePreValues(ByVal cellIndex As Integer, ByRef spikes As Variant, ByRef bursts As Variant, ByVal recDuration As Double)
    'Store background spiking properties (these deal w/ spikes outside ALL bursts, not just wave-bursts)
    Dim newResults() As Variant
    ReDim newResults(1 To NUM_BKGRD_PROPERTIES)
    newResults(1) = backgroundFiringInUnit(spikes, bursts, recDuration)
    newResults(2) = backgroundISIInUnit(spikes, bursts, recDuration)
    newResults(3) = PercentBurstSpikesInUnit(spikes, bursts)
    newResults(4) = burstFreqInUnit(bursts, recDuration)
    newResults(5) = IBIInUnit(bursts, recDuration)
    
    'Output values to the All Averages sheet, adding new results to the old ones
    Dim preCell As Range
    Dim oldResults As Variant
    Set preCell = Worksheets(ALL_AVGS_SHT_NAME).Cells(cellIndex + 1, 1 + 1)
    oldResults = preCell.Resize(1, NUM_BKGRD_PROPERTIES).value
    preCell.Resize(1, NUM_BKGRD_PROPERTIES).value = Application.Pmt(0, -1, oldResults, newResults)   'Pmt() acts like Sum() but can accept two variants (of equal dimension)
End Sub

Private Sub storePostValues(ByVal cellIndex As Integer, ByRef spikes As Variant, ByRef bursts As Variant, ByVal recDuration As Double, ByVal wabRatio As Double)
    'Store wave-associated spiking properties (if this channel HAD wave-associated bursts)
    Dim newResults() As Variant
    ReDim newResults(1 To NUM_BURST_PROPERTIES)
    newResults(1) = BurstDurationInUnit(bursts)
    newResults(2) = BurstSpikeFreqInUnit(spikes, bursts)
    newResults(3) = BurstISIInUnit(spikes, bursts)
    newResults(4) = PercentBurstTimeAboveFreqInUnit(spikes, bursts, 10)
    newResults(5) = SpikesPerBurstInUnit(spikes, bursts)
    
    'Output values to the Bursts Averages sheet, adding new results to the old ones
    Dim postCell As Range
    Dim oldResults As Variant
    Set postCell = Worksheets(BURST_AVGS_SHT_NAME).Cells(cellIndex + 1, 1 + 1)
    oldResults = postCell.Resize(1, NUM_BURST_PROPERTIES).value
    postCell.Resize(1, NUM_BURST_PROPERTIES).value = Application.Pmt(0, -1, oldResults, newResults)   'Pmt() acts like Sum() but can accept two variants (of equal dimension)
    
    'Output the PercentBurstInWaves value to the All Averages sheet, adding it the old result
    Dim pbwCell As Range
    Dim oldPBW As Double
    Set pbwCell = Worksheets(ALL_AVGS_SHT_NAME).Cells(cellIndex + 1, NUM_BKGRD_PROPERTIES + 1)
    oldPBW = pbwCell.value
    pbwCell.value = oldPBW + wabRatio * 100
End Sub

Private Function getSpikeTrain(ByVal spikeCol As Integer) As Variant
    Dim numSpikes As Long
    Dim spikeTrain() As Variant
    
    'If there are no spikes, then return a sentinel value
    numSpikes = Cells(1, spikeCol).End(xlDown).row - 1
    numSpikes = IIf(numSpikes = MAX_EXCEL_ROWS - 1, 0, numSpikes)
    If numSpikes = 0 Then
        ReDim spikeTrain(1 To 1, 1 To 1)
        spikeTrain(1, 1) = -1
        
    'If there is only one spike, Excel wont return an array so we have to construct it
    ElseIf numSpikes = 1 Then
        ReDim spikeTrain(1 To 1, 1 To 1)
        spikeTrain(1, 1) = Cells(2, spikeCol).value
    
    'Otherwise, return the spike train as a 1-column 2D array
    Else
        spikeTrain = Cells(2, spikeCol).Resize(numSpikes, 1).value
    End If
    
    getSpikeTrain = spikeTrain
End Function

Private Function getBurstTrain(ByVal spikeCol As Integer, ByVal numUnits As Integer) As Variant
    Dim burstCol As Integer
    Dim numBursts As Long
    Dim burstTrain() As Variant
    
    'If there are no bursts, then return a sentinel value
    'Otherwise, return the burst timestamps as a 2-column 2D array
    'We don't have to check for one burst as in getSpikeTrain because we still ask for a 2D array from Excel
    burstCol = burstColumn(spikeCol, numUnits)
    numBursts = Cells(1, burstCol).End(xlDown).row - 1
    numBursts = IIf(numBursts = MAX_EXCEL_ROWS - 1, 0, numBursts)
    If numBursts = 0 Then
        ReDim burstTrain(1 To 1, 1 To 1)
        burstTrain(1, 1) = -1
    Else
        burstTrain = Cells(2, burstCol).Resize(numBursts, 2).value
    End If
    
    getBurstTrain = burstTrain
End Function

Private Sub reduceToAvgs(ByVal numRecordings As Integer, ByVal numUnits As Integer)
    Dim data As Range
    Dim averages, sttcs As Variant
    Dim r, c As Integer

    'Don't bother reducing if there were 0 or 1 recordings
    If numRecordings <= 1 Then _
        Exit Sub
    
    'Reduce all sums of WAB property-values to averages
    Set data = Worksheets(ALL_AVGS_SHT_NAME).Cells(2, 2).Resize(numUnits, NUM_PROPERTIES)
    averages = data.value
    For r = 1 To numUnits
        For c = 1 To NUM_PROPERTIES
            averages(r, c) = averages(r, c) / numRecordings
        Next c
    Next r
    data.value = averages
      
    'Reduce all sums of STTC-values to averages
    Set data = Worksheets(STTC_SHT_NAME).Cells(3, 2).Resize(numUnits, numUnits)
    sttcs = data.value
    For r = 1 To numUnits
        For c = 1 To numUnits
            sttcs(r, c) = sttcs(r, c) / numRecordings
        Next c
    Next r
    data.value = sttcs
End Sub

Public Sub finalize(ByVal numUnits As Integer)
    'Finalize the Averages sheet
    With Worksheets(ALL_AVGS_SHT_NAME)
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        .Rows(1).Cells.HorizontalAlignment = xlLeft
        .Cells.EntireRow.AutoFit
        .Cells.EntireColumn.AutoFit
    End With
    
    'Finalize the STTC sheet
    With Worksheets(STTC_SHT_NAME)
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        .Columns.EntireColumn.AutoFit
        .Columns(1).ColumnWidth = .Columns(2).ColumnWidth
        .Cells(1, 1).HorizontalAlignment = xlLeft
        .Cells.EntireRow.AutoFit
    End With
    
    Worksheets(ALL_AVGS_SHT_NAME).Activate
End Sub

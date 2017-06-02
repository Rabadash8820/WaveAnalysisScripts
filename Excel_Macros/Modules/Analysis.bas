Attribute VB_Name = "Analysis"
Option Private Module
Option Explicit

'Processing variables for this module
Private maxBursts As Integer
Private unitNames As Variant
Private sttcResults() As Double, bkgrdResults() As Double, burstResults() As Double
Public Sub AnalyzeTissueWorkbook(ByVal wbName As String, ByVal tissView As cTissueView, ByVal burstsToUse As BurstUseType)
    Dim u As Integer
    
    'If there are no recordings in this workbook then just return
    Dim wb As Workbook
    Set wb = Workbooks.Open(wbName)
    Dim contentsTbl As ListObject
    Set contentsTbl = wb.Worksheets(CONTENTS_NAME).ListObjects(CONTENTS_NAME)
    If contentsTbl.ListRows.Count = 0 Then _
        Exit Sub
            
    'Get the names of all units
    Dim wksht As Worksheet, numUnits As Long
    Set wksht = wb.Worksheets(contentsTbl.ListRows(1).Range(1, 2).Value)
    numUnits = wksht.Cells(1, 1).End(xlToRight).Column / 3    'Since every unit is mentioned once for spikes, burst_start, and burst_end
    unitNames = Application.Transpose(wksht.Cells(1, 1).Resize(1, numUnits))
    
    'If invalid units were provided, then delete their data columns and adjust the unitNames
    Dim recName As String
    If tissView.BadUnits.Count > 0 Then
        recName = contentsTbl.ListRows(1).Range(1, 2)
        Call DeleteUnits(wb.Worksheets(recName), tissView, unitNames)
        numUnits = wksht.Cells(1, 1).End(xlToRight).Column / 3
        unitNames = Application.Transpose(wksht.Cells(1, 1).Resize(1, numUnits))
    End If
    
    'Add output sheets (these lines must come after resetting unitNames)
    Call addAllAvgsSheet
    Call addBurstAvgsSheet
    Call addSttcSheet
    
    'Allocate the result arrays (will automatically be filled with zeroes...)
    Dim numSttcRows As Long
    numSttcRows = numUnits * (numUnits - 1) / 2
    ReDim bkgrdResults(1 To numUnits, 1 To NUM_BKGRD_PROPERTIES)
    ReDim burstResults(1 To numUnits, 1 To NUM_BURST_PROPERTIES)
    ReDim sttcResults(1 To numSttcRows, 1 To 1)
    
    'Process each recording for this tissue (represented as separate sheets)
    Dim startT As Double, endT As Double
    Dim recRng As Range
    Set recRng = contentsTbl.ListRows(1).Range(1, 1)
    recName = recRng.offset(0, 1)
    startT = recRng.offset(0, 2)
    endT = recRng.offset(0, 3)
    wb.Worksheets(recName).Activate
    Call analyzeRecording(unitNames, burstsToUse, startT, endT)
        
    'Store results to their respective Excel tables
    Dim allTbl As ListObject, burstTbl As ListObject, sttcTbl As ListObject
    Set allTbl = Worksheets(ALL_AVGS_NAME).ListObjects(ALL_AVGS_NAME)
    Set burstTbl = Worksheets(BURST_AVGS_NAME).ListObjects(BURST_AVGS_NAME)
    Set sttcTbl = Worksheets(STTC_NAME).ListObjects(STTC_NAME)
    allTbl.DataBodyRange.offset(0, 1).Resize(, allTbl.ListColumns.Count - 1).Value = bkgrdResults
    burstTbl.DataBodyRange.offset(0, 1).Resize(, burstTbl.ListColumns.Count - 1).Value = burstResults
    sttcTbl.DataBodyRange.offset(0, 3).Resize(, sttcTbl.ListColumns.Count - 3).Value = sttcResults

    'Reduce result sums to averages and finalize
    Call cleanSheets
    wb.Close (True)
End Sub

Private Sub addAllAvgsSheet()
    Dim zeroes() As Variant
    Dim u As Integer, p As Integer
    
    Dim numUnits As Integer
    numUnits = UBound(unitNames, 1)
    
    'Add the Averages sheet, write out row (unit) headers, and add formatting
    Dim avgsRng As Range
    Worksheets.Add After:=Sheets(CONTENTS_NAME)
    ActiveSheet.Name = ALL_AVGS_NAME
    Set avgsRng = Worksheets(ALL_AVGS_NAME).Cells(2, 2)
    Cells(1, 1).Value = CELL_STR
    avgsRng.offset(0, -1).Resize(numUnits, 1).Value = unitNames
    avgsRng.offset(-1, 0).Resize(1, NUM_BKGRD_PROPERTIES).Value = PROPERTIES
    avgsRng.offset(-1, -1).Resize(numUnits + 1).Font.Bold = True
    
    'Initialize all output cells to zero
    ReDim zeroes(1 To numUnits, 1 To NUM_BKGRD_PROPERTIES)
    For u = 1 To numUnits
        For p = 1 To NUM_BKGRD_PROPERTIES
            zeroes(u, p) = 0
        Next p
    Next u
    avgsRng.Resize(numUnits, NUM_BKGRD_PROPERTIES).Value = zeroes
        
    'Make a table for all the average values
    Worksheets(ALL_AVGS_NAME).ListObjects.Add( _
        xlSrcRange, _
        avgsRng.CurrentRegion, , _
        xlYes) _
    .Name = ALL_AVGS_NAME
    avgsRng.Resize(numUnits, NUM_BKGRD_PROPERTIES).NumberFormat = "0.000"
End Sub

Private Sub addBurstAvgsSheet()
    Dim zeroes() As Variant
    Dim u As Integer, p As Integer
    
    Dim numUnits As Integer
    numUnits = UBound(unitNames, 1)
    
    'Add the Averages sheet, write out row (unit) headers, and add formatting
    Dim avgsRng As Range
    Worksheets.Add After:=Sheets(ALL_AVGS_NAME)
    ActiveSheet.Name = BURST_AVGS_NAME
    Set avgsRng = Worksheets(BURST_AVGS_NAME).Cells(2, 2)
    Cells(1, 1).Value = CELL_STR
    avgsRng.offset(0, -1).Resize(numUnits, 1).Value = unitNames
    avgsRng.offset(-1, -1).Resize(numUnits + 1).Font.Bold = True
    
    'Write out column (property) headers, indicating which properties are being skipped with a "*"
    Dim pStr As String
    For p = 1 To NUM_BURST_PROPERTIES
        pStr = PROPERTIES(NUM_BKGRD_PROPERTIES + p)
        avgsRng.offset(-1, p - 1).Value = pStr
    Next p
    
    'Initialize all output cells to zero
    ReDim zeroes(1 To numUnits, 1 To NUM_BURST_PROPERTIES)
    For u = 1 To numUnits
        For p = 1 To NUM_BURST_PROPERTIES
            zeroes(u, p) = 0
        Next p
    Next u
    avgsRng.Resize(numUnits, NUM_BURST_PROPERTIES).Value = zeroes
        
    'Make a table for all the average values
    Worksheets(BURST_AVGS_NAME).ListObjects.Add( _
        xlSrcRange, _
        avgsRng.CurrentRegion, , _
        xlYes) _
    .Name = BURST_AVGS_NAME
    avgsRng.Resize(numUnits, NUM_BURST_PROPERTIES).NumberFormat = "0.000"
End Sub

Private Sub addSttcSheet()
    Dim initial() As Variant
    Dim u As Integer, p As Integer
    
    Dim numUnits As Long
    Dim numRows As Long
    numUnits = UBound(unitNames, 1)
    numRows = numUnits * (numUnits - 1) / 2
    
    'Add the STTC sheet, write out row/column (channel) headers, and add formatting
    Dim sttcRng As Range
    Worksheets.Add After:=Sheets(BURST_AVGS_NAME)
    ActiveSheet.Name = STTC_NAME
    Set sttcRng = Worksheets(STTC_NAME).Cells(4, 1)
    With sttcRng
        .offset(-3, 0).Value = STTC_HEADER_STR
        .offset(-3, 0).Font.Bold = True
        .offset(-3, 0).Font.Size = 16
        .offset(-1, 0).Value = "Cell1"
        .offset(-1, 1).Value = "Cell2"
        .offset(-1, 2).Value = "Unit Distance"
        .offset(-1, 3).Value = "STTC"
        .offset(-1, 0).EntireRow.Font.Bold = True
    End With
    
    'Initialize all output cells to zero
    ReDim initial(1 To numRows, 1 To 3)
    Dim u1 As Integer, u2 As Integer, ch1 As Integer, ch2 As Integer
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
    sttcRng.Resize(numRows, 3).Value = initial
        
    'Make a table for all the STTC values
    Dim sttcTbl As ListObject
    Set sttcTbl = Worksheets(STTC_NAME).ListObjects.Add(xlSrcRange, sttcRng.CurrentRegion, , xlYes)
    sttcTbl.Name = STTC_NAME
    sttcTbl.ListColumns(3).DataBodyRange.NumberFormat = "0.000"
    sttcTbl.ListColumns(4).DataBodyRange.NumberFormat = "0.000"
    
End Sub

Private Sub analyzeRecording(ByRef unitNames As Variant, ByVal burstsToUse As BurstUseType, ByVal startTime As Double, ByVal endTime As Double)
    Dim spikes As Variant, preBursts As Variant, postBursts As Variant
        
    'Keep the below For loops separated so that
    'spikes on one channel arent correlated then deleted before pairing with other channels and so that
    'bursts don't get associated with bursts that would later get deleted
    
    'Store STTC values using the entire spike trains of every possible pair of channels
    Dim numUnits As Integer, Duration As Double
    numUnits = UBound(unitNames)
    Duration = endTime - startTime
    Call storeSttcValues(Duration, numUnits)
    
    'Initialize the start/end times of each Unit
    Dim startEndTimes() As Double, u As Integer
    ReDim startEndTimes(1 To numUnits, 1 To 2)
    For u = 1 To numUnits
        startEndTimes(u, 1) = startTime
        startEndTimes(u, 2) = endTime
    Next u
        
    'Remove bursts/spikes that start/end too late/early (lolwut?)
    'Removing spikes may adjust the start/end times of each unit
    If DELETE_BAD_BURSTS Then
        For u = 1 To numUnits
            spikes = getSpikeTrain(u)
            preBursts = getBurstTrain(u, numUnits)
            Call deleteBadBurstsFrom(u, numUnits, spikes, preBursts, startTime, endTime, startEndTimes)
        Next u
    End If
    If DELETE_BAD_SPIKES Then
        For u = 1 To numUnits
            spikes = getSpikeTrain(u)
            Call deleteBadSpikesFrom(u, spikes, startTime, endTime, startEndTimes(u, 1), startEndTimes(u, 2))
        Next u
    End If
    ActiveSheet.UsedRange   'Refresh UsedRange by getting the property
    
    'Do PRE ANALYSES on each unit
    'I.e., analyses BEFORE unused bursts get deleted (background firing metrics)
    Dim preBurstCounts() As Integer
    ReDim preBurstCounts(1 To numUnits)
    For u = 1 To numUnits
        spikes = getSpikeTrain(u)
        preBursts = getBurstTrain(u, numUnits)
        preBurstCounts(u) = UBound(preBursts, 1)
        Duration = startEndTimes(u, 2) - startEndTimes(u, 1)
        Call storePreValues(u, spikes, preBursts, Duration)
    Next u
    
    'Delete unused bursts (WABs or non-WABs), if requested
    If burstsToUse <> BurstUseType.All Then
        Dim wabsOnly As Boolean
        wabsOnly = (burstsToUse = BurstUseType.WABs)
        Call deleteUnusedBursts(wabsOnly, unitNames)
    End If
        
    'Do POST ANALYSES on each unit
    'I.e., analyses AFTER unused bursts have been deleted (wave- or non-wave-associated firing metrics)
    Dim wabRatio As Double
    Dim postBurstCounts() As Integer
    ReDim postBurstCounts(1 To numUnits)
    For u = 1 To numUnits
        spikes = getSpikeTrain(u)
        postBursts = getBurstTrain(u, numUnits)
        postBurstCounts(u) = UBound(postBursts, 1)
        Duration = startEndTimes(u, 2) - startEndTimes(u, 1)
        wabRatio = postBurstCounts(u) / preBurstCounts(u)
        Call storePostValues(u, spikes, postBursts, Duration, wabRatio)
    Next u
    
    'Reformat the datasheet so its a little easier to view
    Cells.VerticalAlignment = xlCenter
    Cells.HorizontalAlignment = xlCenter
    Cells.EntireColumn.AutoFit
End Sub

Private Sub storeSttcValues(ByVal Duration As Double, ByVal numUnits As Long)
    Dim tValues() As Double
    Dim cellIndex1 As Integer, cellIndex2 As Integer
    Dim u1 As Integer, u2 As Integer
    Dim sttc As Double
    Dim outputRng As Range
    Dim oldResults As Variant
    Dim spikes1 As Variant, spikes2 As Variant
    
    Dim numRows As Long
    numRows = numUnits * (numUnits - 1) / 2
            
    'Loop over each unit's spike timestamps (arranged in columns) to get T-values
    ReDim tValues(1 To numUnits)
    For u1 = 1 To numUnits
        'For each unit, store the fraction of the recording's duration wherein its spikes are delta-t apart
        spikes1 = getSpikeTrain(u1)
        tValues(u1) = correlatedTimeProportion(spikes1, Duration, CORRELATION_DT)
    Next u1
    
    'Increment STTC values using the entire spike trains of every possible pair of units
    Dim row As Long
    row = 1
    For u1 = 1 To numUnits
        spikes1 = getSpikeTrain(u1)
        For u2 = u1 + 1 To numUnits
            spikes2 = getSpikeTrain(u2)
            sttc = spikeTimeTilingCoefficient2(spikes1, spikes2, tValues(u1), tValues(u2), CORRELATION_DT)
            sttcResults(row, 1) = sttc
            row = row + 1
        Next u2
    Next u1
    
End Sub

Private Sub deleteBadBurstsFrom(ByVal u As Integer, ByVal numUnits As Integer, ByRef spikes As Variant, ByRef bursts As Variant, ByVal recStart As Double, ByVal recEnd As Double, ByRef startEndTimes() As Double)
    'If there are no bursts then just return
    If bursts(1, 1) = -1 Then Exit Sub
    
    'Store the valid bursts in a new Variant
    Dim activeRow As Integer, offset As Double
    Dim newStart As Double, newEnd As Double, newBursts() As Variant
    Dim b As Integer, bStart As Double, bEnd As Double, burstDur As Double, tooEarly As Boolean, tooLate As Boolean, validDur As Boolean
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
    burstRng.Resize(UBound(newBursts), 2).Value = newBursts
    
    'Set the new start and end times for this unit
    startEndTimes(u, 1) = newStart
    startEndTimes(u, 2) = newEnd
End Sub

Private Sub deleteBadSpikesFrom(ByVal u As Integer, ByRef spikes As Variant, ByVal recStart As Double, ByVal recEnd As Double, ByVal unitStart As Double, ByVal unitEnd As Double)
    Dim s As Long
    
    'If there are no spikes then just return
    If spikes(1, 1) = -1 Then _
        Exit Sub
    
    'Store the valid bursts in a new Variant
    Dim activeRow As Long
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
    'Remove the first/last spikes if they represent the end/start of bursts that were cut off
    Dim spikeRng As Range
    Set spikeRng = Cells(2, u)
    spikeRng.Resize(UBound(spikes)).Clear
    spikeRng.Resize(UBound(newSpikes)).Value = newSpikes
    If unitEnd <> recEnd Then _
        spikeRng.offset(activeRow - 1, 0).Delete Shift:=xlUp
    If unitStart <> recStart Then _
        spikeRng.Delete Shift:=xlUp
        
End Sub

Private Sub deleteUnusedBursts(ByVal wabsOnly As Boolean, ByRef unitNames As Variant)
    Dim burstRng As Range
    Dim numBursts As Integer
    Dim bursts As Variant, trimmedBursts As Variant
    Dim burstPos As Integer
    Dim isWAB As Boolean
    Dim firstU As Integer, lastU As Integer, nFirstU As Integer, nLastU As Integer
    Dim neighbor As Variant
    Dim validNeighbors As Collection
    Dim u As Integer, ch As Integer, nCh As Integer, b As Integer, chPos As Integer
    Dim chStr As String, nChStr As String
    Dim numAssocUnits As Integer
    
    'Get all the burst start/end timestamps
    Dim numUnits As Integer
    numUnits = UBound(unitNames)
    maxBursts = getMaxBursts(numUnits)
    ReDim trimmedBursts(1 To maxBursts, 1 To 2 * numUnits)
    Set burstRng = Cells(2, numUnits + 1).Resize(maxBursts, numUnits * 2)
    bursts = burstRng.Value
    
    'For each unit...
    For u = 1 To numUnits
        ch = channelIndex(unitNames(u, 1))
        chStr = CHANNEL_PREFIX & channelString(ch)
        chPos = 2 * u - 1
        burstPos = 0
        
        'Find the first and last unit on this same channel
        firstU = u
        lastU = u
        If ASSOC_SAME_CHANNEL_UNITS Then
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
            If ASSOC_SAME_CHANNEL_UNITS Then
                Call burstAssociatedWithUnits(firstU, lastU, bursts, chPos, b, numAssocUnits)
                isWAB = (numAssocUnits >= MIN_ASSOC_UNITS)
                If isWAB Then Exit For
            End If
        
            'If not, then do the same check for each of the unit's valid neighbor channels
            For Each neighbor In validNeighbors
                Call neighborUnitsAssociatedWithBurst(u, neighbor, b, bursts, numAssocUnits)
                isWAB = (numAssocUnits >= MIN_ASSOC_UNITS)
                If isWAB Then Exit For
            Next neighbor
            
            'If the burst was or was not wave-associated (as requested by user),
            'then add its start/end timestamps to the "trimmed" array
            If isWAB = wabsOnly Then
                burstPos = burstPos + 1
                trimmedBursts(burstPos, chPos) = bursts(b, chPos)
                trimmedBursts(burstPos, chPos + 1) = bursts(b, chPos + 1)
            End If
        Next b
        
    Next u
    
    'Overwrite the old burst timestamps with the "trimmed" ones
    burstRng.Clear
    burstRng.Value = trimmedBursts
End Sub

Private Sub neighborUnitsAssociatedWithBurst(ByVal unit As Integer, ByVal neighbor As Variant, ByVal b As Integer, ByRef bursts As Variant, ByRef numAssocUnits As Integer)
    Dim trimmedBursts As Variant
    Dim nFirstU As Integer, nLastU As Integer
    Dim nCh As Integer, nChStr As String
    
    Dim numUnits As Integer
    numUnits = UBound(unitNames)
    Dim ch As Integer, chStr As String, chPos As String
    ch = channelIndex(unitNames(unit, 1))
    chStr = CHANNEL_PREFIX & channelString(ch)
    chPos = 2 * unit - 1
    
    'Find the first and last unit on the neighbor channel (if they are represented on the sheet)
    Dim neighborAfter, inChannel As Boolean
    Dim tempU As Integer, endForU As Integer, step As Integer
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
    Dim nU As Integer, nB As Integer, nChPos As Integer, nStart As Double, nFinish As Double, nBinDuration As Double
    Dim associated As Boolean, assocUnitsHere As Integer
    assocUnitsHere = 0
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
                    assocUnitsHere = assocUnitsHere + 1
                    Exit For
                End If
            Next nB
            
            'If the min number of associated units has been achieved, then exit the loop and return
            Dim oneUnitAssoc As Boolean, enoughUnitsAssoc As Boolean
            oneUnitAssoc = (assocUnitsHere = 1 And Not ASSOC_MULTIPLE_UNITS)
            enoughUnitsAssoc = (numAssocUnits + assocUnitsHere >= MIN_ASSOC_UNITS)
            If oneUnitAssoc Or enoughUnitsAssoc Then _
                Exit For
        End If
    Next nU
    
    numAssocUnits = numAssocUnits + assocUnitsHere
End Sub

Private Sub storePreValues(ByVal index As Integer, ByRef spikes As Variant, ByRef bursts As Variant, ByVal recDuration As Double)
    'Calculate some fundamental values
    Dim hasSpikes As Boolean, hasBursts As Boolean
    Dim numSpikes As Long, numBurstSpikes As Long, numBursts As Integer, burstTime As Double, croppedDur As Double
    Dim bkgrdFiringRate As Double, ibi As Double
    hasSpikes = (spikes(1, 1) <> -1)
    hasBursts = (bursts(1, 1) <> -1)
    If hasSpikes Then _
        numSpikes = UBound(spikes)
    If hasBursts Then
        numBursts = UBound(bursts)
        burstTime = burstTimeInUnit(bursts)
        croppedDur = bursts(numBursts, 2) - bursts(1, 1)
    End If
    If hasSpikes And hasBursts Then _
        numBurstSpikes = burstSpikesInUnit(spikes, bursts)
    
    'Calculate more complicated values
    bkgrdFiringRate = (numSpikes - numBurstSpikes) / (recDuration - burstTime)
    If numBursts <> 1 Then _
        ibi = (croppedDur - burstTime) / (numBursts - 1) 'numerator comes out to zero if no bursts
    
    'Store background spiking properties (these deal w/ spikes outside ALL bursts, not just wave-bursts)
    bkgrdResults(index, 1) = numSpikes                                              'Number of spikes
    bkgrdResults(index, 2) = bkgrdFiringRate * 60                                   'Firing rate outside all bursts
    If bkgrdFiringRate > 0 Then _
        bkgrdResults(index, 4) = 1 / bkgrdFiringRate                                'ISI outside all bursts
    If numSpikes > 0 Then _
        bkgrdResults(index, 6) = (numSpikes - numBurstSpikes) / numSpikes * 100     'Percent spikes outside all bursts
    bkgrdResults(index, 8) = numBursts / recDuration * 60                           'Burst frequency
    bkgrdResults(index, 9) = ibi                                                    'Interburst interval
End Sub

Private Sub storePostValues(ByVal index As Integer, ByRef spikes As Variant, ByRef bursts As Variant, ByVal recDuration As Double, ByVal wabRatio As Double)
    'Calculate some fundamental values
    Dim hasSpikes As Boolean, hasBursts As Boolean
    Dim numSpikes As Long, numBursts As Integer, burstTime As Double, numBurstSpikes As Long
    Dim avgBurstDur As Double, avgBurstFiringRate As Double, pctTimeAbove10Hz As Double, firingRateOutside As Double
    hasSpikes = (spikes(1, 1) <> -1)
    hasBursts = (bursts(1, 1) <> -1)
    If hasSpikes Then _
        numSpikes = UBound(spikes)
    If hasBursts Then
        numBursts = UBound(bursts)
        burstTime = burstTimeInUnit(bursts)
    End If
    If hasSpikes And hasBursts Then _
        numBurstSpikes = burstSpikesInUnit(spikes, bursts)
    
    'Calculate more complicated values
    If hasBursts Then
        avgBurstDur = burstDurationInUnit(bursts)
        avgBurstFiringRate = burstFiringRateInUnit(spikes, bursts)
    End If
    If hasSpikes And hasBursts Then _
        pctTimeAbove10Hz = percentBurstTimeAboveFreqInUnit(spikes, bursts, 10)
    firingRateOutside = (numSpikes - numBurstSpikes) / (recDuration - burstTime)

    'Store burst-specific spiking properties (if this channel HAD bursts of the correct type)
    burstResults(index, 1) = numBursts                          'Number of bursts
    burstResults(index, 2) = avgBurstDur                        'Burst duration
    burstResults(index, 3) = avgBurstFiringRate                 'Burst firing rate
    If avgBurstFiringRate > 0 Then _
        burstResults(index, 4) = 1 / avgBurstFiringRate         'Burst ISI
    burstResults(index, 5) = pctTimeAbove10Hz * 100             'Percent time in burst firing >10 Hz
    If numBursts > 0 Then _
        burstResults(index, 6) = numBurstSpikes / numBursts     'Spikes per burst
    
    'Store other "background" properties that had to wait until after removing unneeded bursts
    bkgrdResults(index, 3) = firingRateOutside * 60             'Firing rate outside this kind of burst
    If firingRateOutside > 0 Then _
        bkgrdResults(index, 5) = 1 / firingRateOutside          'ISI outside this kind of burst
    If numSpikes > 0 Then _
        bkgrdResults(index, 7) = (numSpikes - numBurstSpikes) / numSpikes * 100   'Percent spikes outside this kind of burst
    bkgrdResults(index, 10) = wabRatio * 100                   'Percent bursts that are this kind of burst
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
        spikeTrain(1, 1) = Cells(2, spikeCol).Value
    
    'Otherwise, return the spike train as a 1-column 2D array
    Else
        spikeTrain = Cells(2, spikeCol).Resize(numSpikes, 1).Value
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
        burstTrain = Cells(2, burstCol).Resize(numBursts, 2).Value
    End If
    
    getBurstTrain = burstTrain
End Function

Public Sub cleanSheets()
    'Finalize the Averages sheet
    With Worksheets(ALL_AVGS_NAME)
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        .Rows(1).Cells.HorizontalAlignment = xlLeft
        .Cells.EntireRow.AutoFit
        .Cells.EntireColumn.AutoFit
    End With
    
    'Finalize the STTC sheet
    With Worksheets(STTC_NAME)
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        .Columns.EntireColumn.AutoFit
        .Columns(1).ColumnWidth = .Columns(2).ColumnWidth
        .Cells(1, 1).HorizontalAlignment = xlLeft
        .Cells.EntireRow.AutoFit
    End With
    
    Worksheets(ALL_AVGS_NAME).Activate
End Sub

Attribute VB_Name = "FunctionModule"
Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UNIT/CHANNEL FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function channelIndex(ByVal unitName As String) As Integer
    Dim str As String
    Dim r, c As Integer
    
    'Get the 0-based row and column indices from the column header
    str = Mid(unitName, Len(CHANNEL_PREFIX) + 1, 2)
    r = CInt(Left(str, 1)) - 1
    c = CInt(Right(str, 1)) - 1
    
    'Return the 0-based channel index
    channelIndex = MEA_COLS * r + c
End Function
Public Function meanSpikeFreqInUnit(ByRef spikes As Variant, ByVal recDuration As Double) As Double
    'If there are no spikes then return 0
    If spikes(1, 1) = -1 Then
        meanSpikeFreqInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the number of spikes over the recording duration
    meanSpikeFreqInUnit = UBound(spikes) / recDuration
End Function
Public Function burstSpikesInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Long
    Dim b As Integer
    
    'If there are no spikes then return 0
    If spikes(1, 1) = -1 Then
        burstSpikesInUnit = 0
        Exit Function
    End If
    
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        burstSpikesInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the number of spikes outside all bursts
    burstSpikesInUnit = 0
    For b = 1 To UBound(bursts)
        burstSpikesInUnit = burstSpikesInUnit + spikesInBurst(spikes, bursts, b)
    Next b
End Function
Public Function burstTimeInUnit(ByRef bursts As Variant) As Double
    Dim b As Integer
    Dim totalTime As Double
        
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        burstTimeInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the total amount of time spent bursting
    totalTime = 0
    For b = 1 To UBound(bursts)
        totalTime = totalTime + (bursts(b, 2) - bursts(b, 1))
    Next b
    burstTimeInUnit = totalTime
End Function
Public Function neighborChannel(ByVal channel As Integer, neighbor As Integer) As Integer
    Dim r, c As Integer
    Dim rowOffset, colOffset As Integer
    
    'Channel and neighbor indices are 0-based
    r = Int(channel / MEA_ROWS)
    c = channel Mod MEA_ROWS
    rowOffset = Int(neighbor / 3) - 1
    colOffset = (neighbor Mod 3) - 1
    neighborChannel = MEA_COLS * (r + rowOffset) + (c + colOffset)
End Function
Public Function channelString(ByVal channel As Integer) As Integer
    Dim r, c As Integer
    Dim rowOffset, colOffset As Integer
    
    'Channel is 0-based
    r = Int(channel / MEA_ROWS)
    c = channel Mod MEA_ROWS
    channelString = CStr(r + 1) & CStr(c + 1)  'eg, returns "42" for channel in row 4, col 2
End Function
Public Function neighborValid(ByVal channel As Integer, ByVal neighbor As Integer) As Boolean
    Dim r, c, nRow, nCol As Integer
    Dim rowOffset, colOffset As Integer
    Dim nCh As Integer
    Dim onGrid, corner, ground As Boolean

    'Assumes that channel and neighbor are both 0-based, colIndex 1-based
    
    'If the neighbor index points to the channel itself, then return false
    If neighbor = 4 Then
        neighborValid = False
        Exit Function
    End If
    
    'Otherwise, calculate the neighbor's channel
    'Cannot use UDF above because we need intermediate values
    r = Int(channel / MEA_ROWS)
    c = channel Mod MEA_ROWS
    rowOffset = Int(neighbor / 3) - 1
    colOffset = (neighbor Mod 3) - 1
    nRow = r + rowOffset
    nCol = c + colOffset
    nCh = MEA_COLS * nRow + nCol
    
    'Neighbor channel must be on the MEA grid and not the ground channel
    onGrid = (0 <= nRow And nRow < MEA_ROWS) And (0 <= nCol And nCol < MEA_COLS)
    corner = (nCh = 0 Or nCh = MEA_COLS - 1 Or nCh = NUM_CHANNELS - MEA_COLS Or nCh = NUM_CHANNELS - 1)
    ground = (nCh = GROUND_CHANNEL)
    neighborValid = (onGrid And Not corner And Not ground)
End Function
Public Function interElectrodeDistance(ByVal channel1 As Integer, ByVal channel2 As Integer) As Double
    Dim unitDistance As Double
    Dim row1, row2, col1, col2 As Integer
    Dim rowDiff, colDiff As Integer
    
    'Get the row and column of each channel
    row1 = Int(channel1 / MEA_ROWS)
    row2 = Int(channel2 / MEA_ROWS)
    col1 = channel1 Mod MEA_ROWS
    col2 = channel2 Mod MEA_ROWS
    
    'Return the distance between them
    rowDiff = row2 - row1
    colDiff = col2 - col1
    unitDistance = WorksheetFunction.Power(rowDiff * rowDiff + colDiff * colDiff, 0.5)
    interElectrodeDistance = unitDistance
End Function

Private Function correlatedSpikeProportion(ByRef spikes1 As Variant, ByRef spikes2 As Variant) As Double
    Dim s1, start2, end2 As Long
    Dim correlatedSpikes As Long
    
    'If there are no spikes on one of the trains then return 0
    If spikes1(1, 1) = -1 Or spikes2(1, 1) = -1 Then
        correlatedSpikeProportion = 0
        Exit Function
    End If
    
    'Loop through each spike on the first train
    correlatedSpikes = 0
    start2 = 1
    end2 = 1
    For s1 = 1 To UBound(spikes1)
        'If train1's tile has moved past the end of train2 then just break the loop
        If spikes1(s1, 1) - CORRELATION_DT > spikes2(UBound(spikes2), 1) Then _
            Exit For
    
        'Find the first spike in the second train that is at or after the start of this spike's tile
        Do Until spikes2(start2, 1) >= spikes1(s1, 1) - CORRELATION_DT Or start2 = UBound(spikes2)
            start2 = start2 + 1
        Loop
        
        'If this spike is actually in the tile...
        If spikes1(s1, 1) - CORRELATION_DT <= spikes2(start2, 1) And spikes2(start2, 1) <= spikes1(s1, 1) + CORRELATION_DT Then
            'Find the last in the second train that is in this tile
            Do Until spikes2(end2, 1) > spikes1(s1, 1) + CORRELATION_DT Or end2 = UBound(spikes2)
                end2 = end2 + 1
            Loop
            If spikes2(end2, 1) > spikes1(s1, 1) + CORRELATION_DT Then _
                end2 = end2 - 1
            
            'And increment the number of correlated spikes
            correlatedSpikes = correlatedSpikes + 1
        Else
            end2 = start2
        End If
    Next s1
    
    'Return correlated spikes over total spikes
    correlatedSpikeProportion = correlatedSpikes / UBound(spikes1)
End Function
Public Function correlatedTimeProportion(ByRef spikes As Variant, ByVal recordingDuration As Double) As Double
    Dim s1, s2 As Long
    Dim start, finish As Double
    Dim correlatedTime As Double
    
    'If there are no spikes then return 0
    If spikes(1, 1) = -1 Then
        correlatedTimeProportion = 0
        Exit Function
    End If
    
    'Find the start spike of the next tile
    correlatedTime = 0
    s1 = 1
    Do Until s1 > UBound(spikes)
        'Find the last spike of this tile
        start = WorksheetFunction.Max(spikes(s1, 1) - CORRELATION_DT, 0)
        s2 = s1
        finish = -1
        Do While finish = -1
            If s2 = UBound(spikes) Then
                finish = spikes(s2, 1)
            Else
                If spikes(s2, 1) + CORRELATION_DT >= spikes(s2 + 1, 1) - CORRELATION_DT Then
                    s2 = s2 + 1
                Else
                    finish = spikes(s2, 1) + CORRELATION_DT
                End If
            End If
        Loop
        
        'Increment the amount of correlated time
        correlatedTime = correlatedTime + (finish - start)
        s1 = s2 + 1
    Loop
    
    'Return correlated time over total time (duration of recording)
    correlatedTimeProportion = correlatedTime / recordingDuration
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SPIKE FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BURST FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function burstColumn(ByVal spikeCol As Integer, ByVal numUnits) As Integer
    burstColumn = numUnits + spikeCol * 2 - 1
End Function

Public Function spikesInBurst(ByRef spikes As Variant, ByRef bursts As Variant, ByVal bIndex As Long) As Long
    Dim burstStart, burstEnd As Double
    Dim s, startIndex, endIndex, numSpikes As Integer
    
    'Burst index is 1-based
    startIndex = binarySearch(spikes, bursts(bIndex, 1))
    endIndex = binarySearch(spikes, bursts(bIndex, 2))
    spikesInBurst = endIndex - startIndex + 1
End Function
Public Function spikeFreqInBurst(ByRef spikes As Variant, ByRef bursts As Variant, ByVal bIndex As Long) As Double
    Dim Duration As Double
    
    'Burst index is 1-based
    Duration = bursts(bIndex, 2) - bursts(bIndex, 1)
    spikeFreqInBurst = spikesInBurst(spikes, bursts, bIndex) / Duration
End Function
Public Function peakFreqInBurst(ByRef spikes As Variant, ByRef bursts As Variant, ByVal bIndex As Long) As Double
    Dim first, last, s As Integer
    Dim minISI, isi As Double
    
    'Burst index is 1-based
    first = binarySearch(spikes, bursts(bIndex, 1))
    last = binarySearch(spikes, bursts(bIndex, 2))
    minISI = (bursts(bIndex, 2) - bursts(bIndex, 1))
    s = first
    Do While s < last
        isi = spikes(s + 1, 1) - spikes(s, 1)
        minISI = WorksheetFunction.Min(isi, minISI)
        s = s + 1
    Loop
    peakFreqInBurst = 1 / minISI
End Function
Public Function isiInBurst(ByRef spikes As Variant, ByRef bursts As Variant, ByVal bIndex As Long) As Double
    isiInBurst = 1 / spikeFreqInBurst(spikes, bursts, bIndex)
End Function
Public Function burstsAssociated(ByVal start As Double, ByVal binDuration As Double, ByVal nStart As Double, ByVal nBinDuration As Double) As Boolean
    Dim bin, nBin As Integer
    Dim assocBins As Integer
    Dim timeDiff As Double

    burstsAssociated = False
        
    'Loop through each bin of this burst
    For bin = 0 To NUM_BINS - 1
        assocBins = 0
    
        'See if a point in any of the neighboring burst's bins fall within the necessary time range of a point in this burst's bin
        nBin = 0
        Do While nBin <= NUM_BINS And assocBins < MIN_BINS
            timeDiff = Abs((start + bin * binDuration) - (nStart + nBin * nBinDuration))
            If timeDiff <= binDuration / 2 + nBinDuration / 2 Then _
                assocBins = assocBins + 1
            nBin = nBin + 1
        Loop
        
        'If the minimum number of bins fell within the necessary time range, then return true
        If assocBins >= MIN_BINS Then
            burstsAssociated = True
            Exit Function
        End If
    Next bin

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ANALYSIS FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function backgroundFiringInUnit(ByRef spikes As Variant, ByRef bursts As Variant, ByVal recDuration As Double) As Double 'Spikes per min
    Dim nonWaveSpikes As Integer
    Dim nonWaveTime As Double
    
    'If there are no spikes then return 0
    If spikes(1, 1) = -1 Then
        backgroundFiringInUnit = 0
        Exit Function
    End If
    
    nonWaveSpikes = UBound(spikes) - burstSpikesInUnit(spikes, bursts)
    nonWaveTime = recDuration - burstTimeInUnit(bursts)
    backgroundFiringInUnit = nonWaveSpikes / nonWaveTime * 60
End Function
Public Function backgroundISIInUnit(ByRef spikes As Variant, ByRef bursts As Variant, ByVal recDuration As Double) As Double    'Seconds
    Dim bkgrdFiring As Double
    
    'If there are no background spikes then return 0
    bkgrdFiring = backgroundFiringInUnit(spikes, bursts, recDuration)
    If bkgrdFiring = 0 Then
        backgroundISIInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the reciprocal of the background firing rate
    backgroundISIInUnit = 1 / bkgrdFiring * 60
End Function
Public Function burstFreqInUnit(ByRef bursts As Variant, ByVal recDuration As Double) As Double   'Bursts per min
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        burstFreqInUnit = 0
        Exit Function
    End If
    
    'Otherwise return the total number of bursts over the recording duration
    burstFreqInUnit = UBound(bursts) / recDuration * 60
End Function
Public Function ibiInUnit(ByRef bursts As Variant, ByVal recDuration As Double) As Double 'Seconds
    Dim nonWaveTime As Double
    
    'If there is only one burst, then we can't compute IBI so return 0
    Dim numBursts As Integer
    numBursts = UBound(bursts)
    If numBursts = 1 Then
        ibiInUnit = 0
        Exit Function
    End If
    
    'Otherwise return the total time outside bursts (excluding time before first burst and after last burst)
    'Divided by number of inter-burst intervals
    Dim croppedDur As Double
    croppedDur = bursts(numBursts, 2) - bursts(1, 1)
    nonWaveTime = croppedDur - burstTimeInUnit(bursts)
    ibiInUnit = nonWaveTime / (numBursts - 1)
End Function
Public Function burstDurationInUnit(ByRef bursts As Variant) As Double  'Seconds
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        burstDurationInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the total time spent bursting over the number of bursts
    burstDurationInUnit = burstTimeInUnit(bursts) / UBound(bursts)
End Function
Public Function percentBurstSpikesInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Double     'Returns percent like 90.57% not 0.9057
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        percentBurstSpikesInUnit = 0
        Exit Function
    End If
    
    'If there are no spikes then return 0
    If spikes(1, 1) = -1 Then
        percentBurstSpikesInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the number of burst-spikes over the total number of spikes
    percentBurstSpikesInUnit = burstSpikesInUnit(spikes, bursts) / UBound(spikes) * 100
End Function
Public Function burstSpikeFreqInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Double    'Spikes per second
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        burstSpikeFreqInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the average of the spike frequencies of all bursts
    Dim sumSpikeFreq As Double, b As Integer
    sumSpikeFreq = 0
    For b = 1 To UBound(bursts)
        sumSpikeFreq = sumSpikeFreq + spikeFreqInBurst(spikes, bursts, b)
    Next b
    burstSpikeFreqInUnit = sumSpikeFreq / UBound(bursts)
End Function
Public Function burstISIInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Double   'Seconds
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        burstISIInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the reciprocal of the mean in-burst spike frequency
    burstISIInUnit = 1 / burstSpikeFreqInUnit(spikes, bursts)
End Function
Public Function percentBurstTimeAboveFreqInUnit(ByRef spikes As Variant, ByRef bursts As Variant, ByVal freq As Double) As Double
    Dim b, s As Integer
    Dim start, finish As Double
    Dim maxISI, isi As Double
    Dim time As Double
    
    'If there are no spikes then return 0
    If spikes(1, 1) = -1 Then
        percentBurstTimeAboveFreqInUnit = 0
        Exit Function
    End If
    
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        percentBurstTimeAboveFreqInUnit = 0
        Exit Function
    End If
    
    'Otherwise, for each burst, find the percent time spent firing above the given frequency
    'Return the average of this value for all bursts
    Dim avgTime As Double
    avgTime = 0
    maxISI = 1 / freq
    For b = 1 To UBound(bursts)
        'Burst index is 1-based
        start = binarySearch(spikes, bursts(b, 1))
        finish = binarySearch(spikes, bursts(b, 2))
        
        time = 0
        For s = start + 1 To finish
            isi = spikes(s, 1) - spikes(s - 1, 1)
            If (isi < maxISI) Then time = time + isi
        Next s
        avgTime = avgTime + time / (bursts(b, 2) - bursts(b, 1))
    Next b
    percentBurstTimeAboveFreqInUnit = avgTime / UBound(bursts) * 100
End Function
Public Function peakBurstSpikeFreqInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Double
    Dim b As Integer
    Dim sumPeakSpikeFreq As Double
    
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        peakBurstSpikeFreqInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the average of the peak frequencies of all bursts
    sumPeakSpikeFreq = 0
    For b = 1 To UBound(bursts)
        sumPeakSpikeFreq = sumPeakSpikeFreq + peakFreqInBurst(spikes, bursts, b)
    Next b
    peakBurstSpikeFreqInUnit = sumPeakSpikeFreq / UBound(bursts)
End Function
Public Function spikesPerBurstInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Double
    'If there are no bursts then return 0
    If bursts(1, 1) = -1 Then
        spikesPerBurstInUnit = 0
        Exit Function
    End If
    
    'Otherwise, return the number of burst-spikes over the number of bursts
    spikesPerBurstInUnit = burstSpikesInUnit(spikes, bursts) / UBound(bursts)
End Function
Public Function spikeTimeTilingCoefficient1(ByRef spikes1 As Variant, ByRef spikes2 As Variant, ByVal recordingDuration As Double) As Double
    Dim P1, P2 As Double
    Dim T1, T2 As Double
    Dim sttc As Double
    
    '1 in the function name is just there b/c VBA doesnt support overloading...
    
    P1 = correlatedSpikeProportion(spikes1, spikes2)
    P2 = correlatedSpikeProportion(spikes2, spikes1)
    T1 = correlatedTimeProportion(spikes1, recordingDuration)
    T2 = correlatedTimeProportion(spikes2, recordingDuration)
    
    sttc = 0.5 * ((P1 - T2) / (1 - P1 * T2) + (P2 - T1) / (1 - P2 * T1))
    spikeTimeTilingCoefficient1 = sttc
End Function
Public Function spikeTimeTilingCoefficient2(ByRef spikes1 As Variant, ByRef spikes2 As Variant, ByVal T1 As Double, ByVal T2 As Double) As Double
    Dim P1, P2 As Double
    Dim sttc As Double
    
    '2 in the function name is just there b/c VBA doesnt support overloading...
    
    P1 = correlatedSpikeProportion(spikes1, spikes2)
    P2 = correlatedSpikeProportion(spikes2, spikes1)
    
    sttc = 0.5 * ((P1 - T2) / (1 - P1 * T2) + (P2 - T1) / (1 - P2 * T1))
    spikeTimeTilingCoefficient2 = sttc
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'HELPER FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function binarySearch(ByRef list As Variant, ByVal lookup As Double) As Integer
    Dim pos As Integer
    Dim lower, middle, upper As Integer
    
    pos = -1
    upper = UBound(list)
    lower = LBound(list)
    Do While lower <= upper And pos = -1
        middle = Int((upper - lower) / 2) + lower
'        Debug.Print (lower & " " & middle & " " & upper & ": " & list(middle, 1))
        If list(middle, 1) < lookup Then
            lower = middle + 1
        ElseIf list(middle, 1) > lookup Then
            upper = middle - 1
        Else
            pos = middle
        End If
    Loop
    
    binarySearch = pos
End Function

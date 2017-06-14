Attribute VB_Name = "Functions"
Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UNIT/CHANNEL FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function channelIndex(ByVal unitName As String) As Integer
    Dim str As String
    Dim r As Integer, c As Integer
    
    'Get the 0-based row and column indices from the column header
    str = Mid(unitName, Len(CHANNEL_PREFIX) + 1, 2)
    r = CInt(Left(str, 1)) - 1
    c = CInt(Right(str, 1)) - 1
    
    'Return the 0-based channel index
    channelIndex = MEA_COLS * r + c
End Function
Public Function burstSpikesInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Long
    Dim b As Integer
    
    'Return the number of spikes outside all bursts
    burstSpikesInUnit = 0
    For b = 1 To UBound(bursts)
        burstSpikesInUnit = burstSpikesInUnit + spikesInBurst(spikes, bursts, b)
    Next b
End Function
Public Function burstTimeInUnit(ByRef bursts As Variant) As Double
    Dim b As Integer, totalTime As Double
    
    'Otherwise, return the total amount of time spent bursting
    totalTime = 0
    For b = 1 To UBound(bursts)
        totalTime = totalTime + (bursts(b, 2) - bursts(b, 1))
    Next b
    burstTimeInUnit = totalTime
End Function
Public Function neighborChannel(ByVal channel As Integer, neighbor As Integer) As Integer
    Dim r As Integer, c As Integer
    Dim rowOffset As Integer, colOffset As Integer
    
    'Channel and neighbor indices are 0-based
    r = Int(channel / MEA_ROWS)
    c = channel Mod MEA_ROWS
    rowOffset = Int(neighbor / 3) - 1
    colOffset = (neighbor Mod 3) - 1
    neighborChannel = MEA_COLS * (r + rowOffset) + (c + colOffset)
End Function
Public Function channelString(ByVal channel As Integer) As Integer
    Dim r As Integer, c As Integer
    Dim rowOffset As Integer, colOffset As Integer
    
    'Channel is 0-based
    r = Int(channel / MEA_ROWS)
    c = channel Mod MEA_ROWS
    channelString = CStr(r + 1) & CStr(c + 1)  'eg, returns "42" for channel in row 4, col 2
End Function
Public Function neighborValid(ByVal channel As Integer, ByVal neighbor As Integer) As Boolean
    Dim r As Integer, c As Integer, nRow As Integer, nCol As Integer
    Dim rowOffset As Integer, colOffset As Integer
    Dim nCh As Integer
    Dim onGrid As Boolean, corner As Boolean, ground As Boolean

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
    Dim row1 As Integer, row2 As Integer, col1 As Integer, col2 As Integer
    Dim rowDiff As Integer, colDiff As Integer
    
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

Private Function correlatedSpikeProportion(ByRef spikes1 As Variant, ByRef spikes2 As Variant, ByVal dt As Double) As Double
    Dim s1 As Long, start2 As Long, end2 As Long
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
        If spikes1(s1, 1) - dt > spikes2(UBound(spikes2), 1) Then _
            Exit For
    
        'Find the first spike in the second train that is at or after the start of this spike's tile
        Do Until spikes2(start2, 1) >= spikes1(s1, 1) - dt Or start2 = UBound(spikes2)
            start2 = start2 + 1
        Loop
        
        'If this spike is actually in the tile, then
        'find the last spike in the second train that is in this tile
        'and increment the number of correlated spikes
        If Abs(spikes1(s1, 1) - spikes2(start2, 1)) <= dt Then
            Do Until spikes2(end2, 1) > spikes1(s1, 1) + dt Or end2 = UBound(spikes2)
                end2 = end2 + 1
            Loop
            If spikes2(end2, 1) > spikes1(s1, 1) + dt Then _
                end2 = end2 - 1
            correlatedSpikes = correlatedSpikes + 1
        Else
            end2 = start2
        End If
    Next s1
    
    'Return correlated spikes over total spikes
    correlatedSpikeProportion = correlatedSpikes / UBound(spikes1)
End Function
Public Function correlatedTimeProportion(ByRef spikes As Variant, ByVal recordingDuration As Double, ByVal dt As Double) As Double
    Dim s1 As Long, s2 As Long
    Dim start As Double, finish As Double
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
        start = WorksheetFunction.Max(spikes(s1, 1) - dt, 0)
        s2 = s1
        finish = -1
        Do While finish = -1
            If s2 = UBound(spikes) Then
                finish = spikes(s2, 1)
            Else
                If spikes(s2, 1) + dt >= spikes(s2 + 1, 1) - dt Then
                    s2 = s2 + 1
                Else
                    finish = spikes(s2, 1) + dt
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
    Dim burstStart As Double, burstEnd As Double
    Dim s As Integer, startIndex As Long, endIndex As Long, numSpikes As Integer
    
    'Burst index is 1-based
    startIndex = binarySearch(spikes, bursts(bIndex, 1))
    endIndex = binarySearch(spikes, bursts(bIndex, 2))
    spikesInBurst = endIndex - startIndex + 1
End Function
Public Function firingRateInBurst(ByRef spikes As Variant, ByRef bursts As Variant, ByVal bIndex As Long) As Double
    Dim Duration As Double, numSpikes As Long
    
    'Burst index is 1-based
    Duration = bursts(bIndex, 2) - bursts(bIndex, 1)
    numSpikes = spikesInBurst(spikes, bursts, bIndex)
    firingRateInBurst = numSpikes / Duration
End Function
Public Function burstsAssociated(ByVal start As Double, ByVal binDuration As Double, ByVal nStart As Double, ByVal nBinDuration As Double) As Boolean
    Dim bin As Integer, nBin As Integer
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
Public Function burstDurationInUnit(ByRef bursts As Variant) As Double  'Seconds
    Dim sumDuration As Double, b As Integer
    
    'Average burst duration
    sumDuration = 0
    For b = 1 To UBound(bursts)
        sumDuration = sumDuration + bursts(b, 2) - bursts(b, 1)
    Next b
    burstDurationInUnit = sumDuration / UBound(bursts)
End Function
Public Function burstFiringRateInUnit(ByRef spikes As Variant, ByRef bursts As Variant) As Double    'Spikes per second
    'Otherwise, return the average of the spike frequencies of all bursts
    Dim sumSpikeFreq As Double, b As Integer
    sumSpikeFreq = 0
    For b = 1 To UBound(bursts)
        sumSpikeFreq = sumSpikeFreq + firingRateInBurst(spikes, bursts, b)
    Next b
    burstFiringRateInUnit = sumSpikeFreq / UBound(bursts)
End Function
Public Function percentBurstTimeAboveFreqInUnit(ByRef spikes As Variant, ByRef bursts As Variant, ByVal freq As Double) As Double
    Dim b As Integer, s As Long
    Dim start As Long, finish As Long
    Dim maxISI As Double, isi As Double
    Dim time As Double
    
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
    percentBurstTimeAboveFreqInUnit = avgTime / UBound(bursts)
End Function
Public Function spikeTimeTilingCoefficient1(ByRef spikes1 As Variant, ByRef spikes2 As Variant, ByVal recordingDuration As Double, ByVal dt As Double) As Double
    Dim P1 As Double, P2 As Double
    Dim T1 As Double, T2 As Double
    Dim sttc As Double
    
    '1 in the function name is just there b/c VBA doesnt support overloading...
    
    P1 = correlatedSpikeProportion(spikes1, spikes2, dt)
    P2 = correlatedSpikeProportion(spikes2, spikes1, dt)
    T1 = correlatedTimeProportion(spikes1, recordingDuration, dt)
    T2 = correlatedTimeProportion(spikes2, recordingDuration, dt)
    
    sttc = 0.5 * ((P1 - T2) / (1 - P1 * T2) + (P2 - T1) / (1 - P2 * T1))
    spikeTimeTilingCoefficient1 = sttc
End Function
Public Function spikeTimeTilingCoefficient2(ByRef spikes1 As Variant, ByRef spikes2 As Variant, ByVal T1 As Double, ByVal T2 As Double, ByVal dt As Double) As Double
    Dim P1 As Double, P2 As Double
    Dim sttc As Double
    
    '2 in the function name is just there b/c VBA doesnt support overloading...
    
    P1 = correlatedSpikeProportion(spikes1, spikes2, dt)
    P2 = correlatedSpikeProportion(spikes2, spikes1, dt)
    
    sttc = 0.5 * ((P1 - T2) / (1 - P1 * T2) + (P2 - T1) / (1 - P2 * T1))
    spikeTimeTilingCoefficient2 = sttc
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'HELPER FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function binarySearch(ByRef list As Variant, ByVal lookup As Double) As Long
    Dim pos As Long
    Dim lower As Long, middle As Long, upper As Long
    
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

Public Function numDimensions(ByRef arr As Variant)
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

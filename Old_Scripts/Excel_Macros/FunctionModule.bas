Attribute VB_Name = "FunctionModule"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHANNEL FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function channelStart(ByVal channel As Integer) As Double
    channelStart = spikeStamp(channel, 1)
End Function
Public Function channelEnd(ByVal channel As Integer) As Double
    channelEnd = spikeStamp(channel, spikesInChannel(channel))
End Function
Public Function channelDuration(ByVal channel As Integer) As Double
    channelDuration = channelEnd(channel) - channelStart(channel)
End Function
Public Function spikeTrainOfChannel(ByVal channel As Integer) As Variant
    'Assumes that channel is 0-based
    'Excel won't count the column header
    
    'If there are no spikes, thenreturn a sentinel value
    Dim numSpikesInChannel As Long
    Dim spikeTrain() As Variant
    numSpikesInChannel = Cells(2, channel + 1).End(xlDown).row - 1
    If numSpikesInChannel = 1048576 - 1 Then
        ReDim spikeTrain(1 To 1, 1 To 1)
        spikeTrain(1, 1) = -1
    Else
        spikeTrain = Cells(2, channel + 1).Resize(numSpikesInChannel, 1).Value
    End If
    spikeTrainOfChannel = spikeTrain
End Function
Public Function burstTrainOfChannel(ByVal channel As Integer) As Variant
    'Assumes that channel is 0-based
    'Excel won't count the column header
    
    'If there are no bursts, then return a sentinel value
    Dim burstCol, numBurstsInChannel As Long
    Dim burstTrain() As Variant
    burstCol = burstColumn(channel)
    numBurstsInChannel = Cells(2, burstCol).End(xlDown).row - 1
    If numBurstsInChannel = 1048576 - 1 Then
        ReDim burstTrain(1 To 1, 1 To 1)
        burstTrain(1, 1) = -1
    Else
        burstTrain = Cells(2, burstCol).Resize(numBurstsInChannel, 2).Value
    End If
    burstTrainOfChannel = burstTrain
End Function
Public Function spikesInChannel(ByVal channel As Integer) As Long
    'Excel won't count the column header
    spikesInChannel = WorksheetFunction.Count(Columns(channel + 1))
End Function
Public Function MeanSpikeFreqOnChannel(ByVal channel As Integer) As Double
    spikeFreqInChannel = spikesInChannel(channel) / (TIME_END - TIME_START)
End Function
Public Function BurstSpikesInChannel(ByVal channel As Integer) As Double
    Dim burst As Integer
    
    BurstSpikesInChannel = 0
    For burst = 1 To burstsInChannel(channel)
        BurstSpikesInChannel = BurstSpikesInChannel + spikesInBurst(channel, burst)
    Next burst
End Function
Public Function BurstTimeInChannel(ByVal channel As Integer) As Double
    Dim burst As Integer
    
    BurstTimeInChannel = 0
    For burst = 1 To burstsInChannel(channel)
        BurstTimeInChannel = BurstTimeInChannel + burstDuration(channel, burst)
    Next burst
End Function
Public Function burstsInChannel(ByVal channel As Integer) As Integer
    'Assumes that channel is 0-based
    'Excel won't count the column header
    burstsInChannel = WorksheetFunction.Count(Columns(burstColumn(channel)))
End Function
Public Function neighborChannel(ByVal channel As Integer, neighbor As Integer) As Integer
    Dim r, c As Integer
    Dim rowOffset, colOffset As Integer

    r = Int(channel / MEA_ROWS)
    c = channel Mod MEA_ROWS
    rowOffset = WorksheetFunction.Floor(neighbor / 3, 1) - 1
    colOffset = (neighbor Mod 3) - 1
    neighborChannel = MEA_ROWS * (r + rowOffset) + (c + colOffset)
End Function
Public Function neighborValid(ByVal channel As Integer, ByVal neighbor As Integer) As Boolean
    Dim r, c As Integer
    Dim rowOffset, colOffset As Integer
    Dim tempChannel As Integer

    'Assumes that channel and neighbor are both 0-based
    neighborValid = False
    If (neighbor <> 4) Then
        r = Int(channel / MEA_ROWS)
        c = channel Mod MEA_ROWS
        rowOffset = WorksheetFunction.Floor(neighbor / 3, 1) - 1
        colOffset = (neighbor Mod 3) - 1
        tempChannel = MEA_ROWS * (r + rowOffset) + (c + colOffset)
        
        'Neighbor channel must be on the MEA grid and not one of the corners
        If (0 <= r + rowOffset And r + rowOffset < MEA_ROWS) And (0 <= c + colOffset And c + colOffset < MEA_COLS) And _
            tempChannel <> 0 And tempChannel <> MEA_COLS And tempChannel <> (MEA_ROWS - 1) * MEA_COLS And tempChannel <> NUM_CHANNELS Then
            neighborValid = True
        End If
    End If
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
    interElectrodeDistance = INTER_ELECTRODE_DISTANCE * unitDistance
End Function
Public Function CorrelatedSpikeProportion(ByVal channel1 As Integer, ByVal channel2 As Integer) As Double
    Dim p As Double
    p = 2
    CorrelatedSpikeProportion = p
End Function
Public Function CorrelatedTimeProportion(ByVal channel As Integer) As Double
    Dim t As Double
    t = 2
    CorrelatedTimeProportion = t
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SPIKE FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function spikeStamp(ByVal channel As Integer, ByVal spike As Long) As Double
    'Assumes that channel is 0-based and spike is 1-based
    spikeStamp = Cells(spike + 1, channel + 1).Value
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BURST FUNCTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function burstColumn(ByVal channel As Integer) As Integer
    'Assumes that channel is 0-based
    burstColumn = NUM_CHANNELS + (2 * channel) + 1
End Function
Public Function burstStart(ByVal channel As Integer, ByVal burst As Long) As Double
    burstStart = 0 + Cells(burst + 1, burstColumn(channel)).Value
End Function
Public Function burstEnd(ByVal channel As Integer, ByVal burst As Long) As Double
    burstEnd = 0 + Cells(burst + 1, burstColumn(channel) + 1).Value
End Function
Public Function burstDuration(ByVal channel As Integer, ByVal burst As Long) As Double
    'Assumes that channel and neighbor are 0-based, burst is 1-based,
    'the start and finish timestamp cells are in order horizontally, and that there are row headers
    burstDuration = burstEnd(channel, burst) - burstStart(channel, burst)
End Function
Public Function firstSpikeRowInBurst(ByVal channel As Integer, ByVal burst As Integer) As Integer
    Dim r As Integer
    Dim bStart As Double

    r = 2
    bStart = burstStart(channel, burst)
    Do Until Cells(r, channel + 1).Value >= bStart
        r = r + 1
    Loop
    firstSpikeRowInBurst = r
End Function
Public Function lastSpikeRowInBurst(ByVal channel As Integer, ByVal burst As Integer) As Integer
    Dim r As Integer
    Dim bEnd As Double

    r = Cells(2, channel + 1).End(xlDown).row
    bEnd = burstEnd(channel, burst)
    Do Until Cells(r, channel + 1).Value <= bEnd
        r = r - 1
    Loop
    lastSpikeRowInBurst = r
End Function
Public Function spikesInBurst(ByVal channel As Integer, ByVal burst As Integer) As Integer
    'Assumes that channel is 0-based and burst is 1-based
    spikesInBurst = WorksheetFunction.CountIfs(Columns(channel + 1), _
        ">=" & burstStart(channel, burst), Columns(channel + 1), "<=" & burstEnd(channel, burst))
End Function
Public Function spikeFreqInBurst(ByVal channel As Integer, ByVal burst As Integer) As Double
    spikeFreqInBurst = spikesInBurst(channel, burst) / burstDuration(channel, burst)
End Function
Public Function peakFreqInBurst(ByVal channel As Integer, ByVal burst As Integer) As Double
    Dim r As Integer
    Dim first, last As Integer
    Dim minISI As Double
    Dim timestamp, nextTimestamp As Double
    
    'Assumes that channel is 0-based and that there are column headers
    first = firstSpikeRowInBurst(channel, burst)
    last = lastSpikeRowInBurst(channel, burst)
    minISI = burstDuration(channel, burst)
    r = first
    Do While r < last
        timestamp = 0 + Cells(r, channel + 1).Value 'Additional 0 needed for Excel to treat as a number in some cases
        nextTimestamp = 0 + Cells(r + 1, channel + 1).Value
        minISI = WorksheetFunction.Min(nextTimestamp - timestamp, minISI)
        r = r + 1
    Loop
    peakFreqInBurst = 1 / minISI
End Function
Public Function ISIInBurst(ByVal channel As Integer, ByVal burst As Integer) As Double
    ISIInBurst = (burstEnd(channel, burst) - burstStart(channel, burst)) / (spikesInBurst(channel, burst) - 1)
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
Public Function backgroundFiringOnChannel(ByVal channel As Integer) As Double
    backgroundFiringOnChannel = (spikesInChannel(channel) - BurstSpikesInChannel(channel)) / _
                                ((TIME_END - TIME_START) - BurstTimeInChannel(channel)) * 60    'Spikes per min
End Function
Public Function backgroundISIOnChannel(ByVal channel As Integer) As Double
    backgroundISIOnChannel = 1 / backgroundFiringOnChannel(channel) * 60  'Seconds
End Function
Public Function burstFreqOnChannel(ByVal channel As Integer) As Double
    burstFreqOnChannel = burstsInChannel(channel) / (TIME_END - TIME_START) * 60    'Bursts per min
End Function
Public Function IBIOnChannel(ByVal channel As Integer) As Double
    IBIOnChannel = ((TIME_END - TIME_START) - BurstTimeInChannel(channel)) / (burstsInChannel(channel) - 1)     'Seconds
End Function
Public Function BurstDurationOnChannel(ByVal channel As Integer) As Double
    BurstDurationOnChannel = BurstTimeInChannel(channel) / burstsInChannel(channel) 'Seconds
End Function
Public Function PercentBurstSpikesOnChannel(ByVal channel As Integer) As Double
    'Returns a percent like 90.57% not 0.9057
    PercentBurstSpikesOnChannel = BurstSpikesInChannel(channel) / spikesInChannel(channel) * 100
End Function
Public Function BurstSpikeFreqOnChannel(ByVal channel As Integer) As Double
    BurstSpikeFreqOnChannel = BurstSpikesInChannel(channel) / BurstTimeInChannel(channel)   'Spikes per second
End Function
Public Function PeakBurstSpikeFreqOnChannel(ByVal channel As Integer) As Double
    Dim burst As Integer
    Dim sumPeakSpikeFreq As Double
    
    sumPeakSpikeFreq = 0
    For burst = 1 To burstsInChannel(channel)
        sumPeakSpikeFreq = sumPeakSpikeFreq + peakFreqInBurst(channel, burst)
    Next burst
    PeakBurstSpikeFreqOnChannel = sumPeakSpikeFreq / burstsInChannel(channel)
End Function
Public Function BurstISIOnChannel(ByVal channel As Integer) As Double
    BurstISIOnChannel = BurstTimeInChannel(channel) / (BurstSpikesInChannel(channel) - burstsInChannel(channel))    'Seconds
End Function
Public Function SpikesPerBurstOnChannel(ByVal channel As Integer) As Double
    SpikesPerBurstOnChannel = BurstSpikesInChannel(channel) / burstsInChannel(channel)
End Function
Public Function SpikeTimeTilingCoefficient(ByVal channel1 As Integer, ByVal channel2 As Integer) As Double
    Dim P1, P2, T1, T2 As Double
    Dim sttc As Double
    
    P1 = CorrelatedSpikeProportion(channel1, channel2)
    P2 = CorrelatedSpikeProportion(channel2, channel1)
    T1 = CorrelatedTimeProportion(channel1)
    T2 = CorrelatedTimeProportion(channel2)
    
    sttc = 0.5 * ((P1 - T2) / (1 - P1 * T2) + (P2 - T1) / (1 - P2 * T1))
    SpikeTimeTilingCoefficient = sttc
End Function

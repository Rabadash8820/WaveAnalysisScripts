Attribute VB_Name = "SpikestampModule"
Option Explicit

Public Sub GetWaveAssociatedSpikeTimestamps()
    Dim channel, burst As Integer
    Dim numWABs As Integer
    Dim afterRow As Integer
    
    Sheets("InputSpikestamps").Activate
    
    'Loop over each wave-associated burst (WAB) of each channel
    For channel = 1 To NUM_CHANNELS - 2  '0-based but ignoring first and last
        afterRow = 1
        numWABs = NumBurstsOnChannel(channel)
        For burst = 1 To numWABs
            
            'Remove all spikes before the burst but after "afterRow"
            Call clearSpikesBeforeBurst(burst, channel, afterRow)
            afterRow = lastSpikeRowInBurst(channel, burst)
            
        Next burst
        
        'Remove any spikes after the last burst
        Call clearSpikesAfterRow(afterRow, channel)
        
    Next channel
    
    'Delete corner channel and burst start/end timestamp columns
    Call deleteColumns

End Sub
Private Sub clearSpikesBeforeBurst(ByVal burst As Integer, ByVal channel As Integer, ByVal afterRow As Double)
    Dim spikeCol As Integer
    Dim first As Integer
    
    'Don't bother deleting rows if there aren't any between afterRow and the burst
    first = firstSpikeRowInBurst(channel, burst)
    If first = afterRow + 1 Then Exit Sub
    
    spikeCol = channel + 1
'    Range(Cells(afterRow + 1, spikeCol), Cells(first - 1, spikeCol)).Select
    Range(Cells(afterRow + 1, spikeCol), Cells(first - 1, spikeCol)).Delete (xlShiftUp)
End Sub
Private Sub clearSpikesAfterRow(ByVal afterRow As Double, ByVal channel As Integer)
    Dim spikeCol As Integer
    Dim last As Integer
    
    'Don't bother trying to delete rows if there aren't any after afterRow
    spikeCol = channel + 1
    If Cells(afterRow + 1, spikeCol).Value = "" Then Exit Sub
    
    last = Cells(1, channel + 1).End(xlDown).row
'    Range(Cells(afterRow + 1, spikeCol), Cells(last, spikeCol)).Select
    Range(Cells(afterRow + 1, spikeCol), Cells(last, spikeCol)).Delete (xlShiftUp)
End Sub

Private Sub deleteColumns()
    'Delete corner channels (increasing subtracted numbers are necessary because columns are being deleted...)
    Columns(1 - 0).Delete
    Columns(MEA_COLS - 1).Delete
    Columns((MEA_ROWS - 1) * MEA_COLS + 1 - 2).Delete
    Columns(NUM_CHANNELS - 3).Delete
    
    'Delete all the burst start/end timestamp columns
    Dim firstBurst, lastBurst As Integer
    firstBurst = NUM_CHANNELS - 4 + 1
    lastBurst = Columns(firstBurst).End(xlToRight).Column
    Columns(firstBurst).Resize(, lastBurst - firstBurst + 1).Delete
End Sub

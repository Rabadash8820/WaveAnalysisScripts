Attribute VB_Name = "Optimization"
Option Explicit
Option Private Module

'Global variables
Public programStart As Date

Public Sub setupOptimizations()
    'Optimize application while macro runs
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Get the start time of the calling program
    programStart = Now
End Sub

Public Sub tearDownOptimizations()
    'Restore application state
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Public Function ProgramDuration() As Date
    ProgramDuration = Now - programStart
End Function

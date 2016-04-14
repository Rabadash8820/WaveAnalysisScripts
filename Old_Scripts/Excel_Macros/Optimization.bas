Attribute VB_Name = "Optimization"
Option Explicit
Option Private Module

'Local optimization variables
Dim screenUpdateState As Boolean
Dim statusBarState As Boolean
Dim eventState As Boolean
Dim pageBreakState As Boolean
Dim calcState As XlCalculation

'Global optimization variables
Public programStart As Date

Public Sub setupOptimizations()
    'Store application state
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventState = Application.EnableEvents
    
    'Optimize application while macro runs
    Application.ScreenUpdating = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Get the start time of the calling program
    programStart = Now
End Sub

Public Sub tearDownOptimizations()
    'Restore application state
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventState
End Sub

Public Function ProgramDuration() As Date
    ProgramDuration = Now - programStart
End Function

Attribute VB_Name = "Util"
Option Explicit
Option Private Module

'Global variables for this module
Public programStart As Date

'Global constants for this module
Private Const TIME_TAKEN_STR = "Time taken (hh:mm:ss)> "

Public Sub setupOptimizations()
    'Optimize application while macro runs
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    
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

Public Function PickWorkbook(ByVal pickMsg As String) As File
    Dim wbName As Workbook
    
    'Create the file-selection dialog box
    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.Title = pickMsg
    dialog.AllowMultiSelect = False
    
    'If the user didn't select anything, then return an empty string
    If dialog.Show = False Then
        Set PickWorkbook = Nothing
        Exit Function
    End If
    
    'Make sure that the selected file was actually an Excel workbook
    Dim fileSystem As New FileSystemObject
    Dim wbFile As File
    Set wbFile = fileSystem.GetFile(dialog.SelectedItems(1))
    Dim correctType As Boolean
    correctType = wbFile.Type = "Microsoft Excel Worksheet" Or wbFile.Type = "Microsoft Excel Macro-Enabled Worksheet"
    If Not correctType Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Selected file is not an Excel workbook.", vbOKOnly, "Routine complete")
        Exit Function
    End If
    
    'If it was then return the File
    Set PickWorkbook = wbFile
End Function

Public Sub showLog(ByRef outputStrs As Collection)
    'Add time taken, and provided output strings, to the MainListBox
    Dim log As New LogForm, str As Variant
    With log.MainListBox
        .AddItem TIME_TAKEN_STR & Format(Now - programStart, "hh:mm:ss")
        .AddItem ""
        For Each str In outputStrs
            .AddItem str
        Next str
    End With
    
    log.Show
End Sub

# WaveAnalysisScripts
All the old wave analysis scripts from my research with Dr. Jordan Renna ("forked" from MEA Cruncher)

# Module Manager
The Module Manager handles automatic importing and exporting of Excel class, form, and VB modules.  Basically, the Manager imports all modules (*.bas, *.frm, *.frx, and *.cls files) stored in the same directory as your workbook when the workbook is opened, re-exports them every time it is saved, and removes them when the workbook is closed.

## Benefits
* Storing code as text allows macros to be edited in the IDE of your choice, like Visual Studio or Notepad++.  All changes will be imported the next time you open the workbook!
* Storing code as text (rather inside a binary workbook file) allows change tracking with VCS software like Git or SVN.
* Storing code in files separate from the main workbook eases collaboration on multiple macros within the same workbook, and allows for more logically distinct commits.
* By removing modules when the workbook is closed, the Manager offers the above benefits *without* duplicating code between the text files and the workbook itself.

## Setup
ModuleManager works with Excel 2007 and later (not tested on 2003 or earlier).

1. __Import the ModuleManager module__ file into your workbook(s).  Within the VB Editor (VBE), in the Project Explorer view, right click anywhere under the name of your workbook and select "Import file...".  Select the ModuleManager.bas file that you just downloaded and click "Open".  (Note, normal module management does not apply to the ModuleManager itself, i.e. it will always be present in the workbook and will not be re-exported or removed).

2. __Add necessary references.__  Within the VBE, select "Tools > References".  In the dialog box, make sure that the following references are selected (if any references are already selected, you should probably leave them checked):
 * Visual Basic For Applications
 * Microsoft Excel x.x Object Library
 * OLE Automation
 * Microsoft Office x.x Object Library
 * Microsoft Scripting Runtime
 * Microsoft Visual Basic for Applications Extensibility x.x

3. __Enable developer macro settings.__  In Excel, click "File > Options > Trust Center > Trust Center Settings...".  In the dialog box, select "Macro Settings", then check "Enable all macros" (or "Disable all macros except digitally signed macros" if you know what you're doing), __and__ "Trust access to the VBA project object model".

4. __Paste the following code__ into the ThisWorkbook module of your workbook.  This is the code that actually handles the Workbook Open, Save, and Close events.  Without it, ModuleManager would just take up space!
```
Option Explicit

Dim alreadySaved As Boolean

Private Sub Workbook_Open()
    Call importMacros
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    Call exportMacros
End Sub
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    'Prevent a save event loop
    If alreadySaved Then
        alreadySaved = False
        Exit Sub
    End If
    
    'Remove all modules and save (so that modules are never saved with this workbook)
    alreadySaved = True
    Call removeMacros
    ThisWorkbook.Save
End Sub
```

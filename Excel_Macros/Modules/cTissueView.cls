VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTissueView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Tissue As cTissue
Public Population As cPopulation
Public RecordingViews As New Collection
Public BadUnits As New Collection
Public WorkbookPaths As New Dictionary

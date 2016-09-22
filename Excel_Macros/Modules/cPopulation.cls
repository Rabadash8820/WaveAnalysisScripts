VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPopulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Abbreviation As String
Public IsControl As Boolean
Public ForeColor As Long
Public BackColor As Long
Public TissueViews As New Collection
Public SheetTissueIDs As New Dictionary

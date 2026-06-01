VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assumed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Utils"
'@ModuleDescription("Provides safe helpers for assumptions about the active Excel context.")
Option Explicit

'@Description("Returns the active worksheet when available.")
Public Property Get ActiveWorksheet() As Worksheet
  On Error Resume Next
  Set ActiveWorksheet = ActiveSheet
  On Error GoTo 0
End Property

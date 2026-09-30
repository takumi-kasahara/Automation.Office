VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Exception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Utils"
'@ModuleDescription("Displays VBA runtime errors to the user.")
Option Explicit

'@Description("Shows the current Err details in a message box.")
Public Sub Show()
  Debug.Assert Err.Number <> 0
  MsgBox Err.Description & "(" & Err.Number & ")", vbExclamation
End Sub

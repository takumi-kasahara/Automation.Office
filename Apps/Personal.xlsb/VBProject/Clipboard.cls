VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Utils")
'@ModuleDescription("Wraps clipboard read and write operations.")
Option Explicit

'@Description("Copies text to the system clipboard and returns success state.")
Public Function Copy(ByVal text As String) As Boolean
  On Error GoTo Catch

  With New DataObject
    .SetText text
    .PutInClipboard
  End With

  Copy = True
  Exit Function
Catch:
  Copy = False
  MsgBox "Failed to copy to clipboard."
End Function

'@Description("Gets text content from the system clipboard.")
Public Property Get value() As String
  On Error GoTo Catch

  With New DataObject
    .GetFromClipboard
    value = .GetText
  End With

  Exit Property
Catch:
  MsgBox "Failed to get clipboard value."
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Utils"
'@ModuleDescription("Executes shell commands and returns standard output.")
Option Explicit

'@Description("Runs a command and returns stdout as text.")
Public Function Exec(ByVal cmd As String) As String
  With New WshShell
    With .Exec(cmd)
      Do While .Status
        DoEvents
      Loop

      Exec = .StdOut.ReadAll
      Debug.Print .StdErr.ReadAll
    End With
  End With
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SaveFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utils.Dialogs"
'@ModuleDescription("Wraps Save As file selection with overwrite confirmation.")
Option Explicit

'@VariableDescription("Initial file name shown in the save dialog.")
Public InitialFileName As Variant
'@VariableDescription("Filter pattern string used by the save dialog.")
Public FileFilter As Variant
'@VariableDescription("Selected filter index used by the save dialog.")
Public FilterIndex As Variant
'@VariableDescription("Dialog title text.")
Public Title As Variant
'@VariableDescription("Custom button text.")
Public ButtonText As Variant

'@Description("Shows the save dialog and returns the selected file path.")
Public Function Show() As String
  Dim path As Variant
  path = Application.GetSaveAsFilename(InitialFileName, FileFilter, FilterIndex, Title, ButtonText)
  If VarType(path) = vbBoolean Then Exit Function

  With New FileSystemObject
    If .FileExists(path) Then
      Dim msg As String
      msg = path & vbNewLine & "already exists. Do you want to replace it ?"
      If MsgBox(msg, vbOKCancel) <> vbOK Then
        Show = vbNullString
      End If
    End If
    Show = path
  End With
End Function

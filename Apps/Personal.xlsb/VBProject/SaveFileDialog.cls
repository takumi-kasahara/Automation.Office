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
Attribute VB_Description = "Wraps Save As file selection with overwrite confirmation."
Option Explicit

'@VariableDescription("Initial file name shown in the save dialog.")
Public InitialFileName As Variant
Attribute InitialFileName.VB_VarDescription = "Initial file name shown in the save dialog."
'@VariableDescription("Filter pattern string used by the save dialog.")
Public FileFilter As Variant
Attribute FileFilter.VB_VarDescription = "Filter pattern string used by the save dialog."
'@VariableDescription("Selected filter index used by the save dialog.")
Public FilterIndex As Variant
Attribute FilterIndex.VB_VarDescription = "Selected filter index used by the save dialog."
'@VariableDescription("Dialog title text.")
Public Title As Variant
Attribute Title.VB_VarDescription = "Dialog title text."
'@VariableDescription("Custom button text.")
Public ButtonText As Variant
Attribute ButtonText.VB_VarDescription = "Custom button text."

'@Description("Shows the save dialog and returns the selected file path.")
Public Function Show() As String
Attribute Show.VB_Description = "Shows the save dialog and returns the selected file path."
  Dim path As Variant
  path = Application.GetSaveAsFilename(InitialFileName, FileFilter, FilterIndex, Title, ButtonText)
  If VarType(path) = vbBoolean Then Exit Function

  With New FileSystemObject
    If .FileExists(path) Then
      Dim msg As String
      msg = path & vbNewLine & "already exists. Do you want to overwrite it?"
      If MsgBox(msg, vbOKCancel) <> vbOK Then
        Show = vbNullString
        Exit Function
      End If
    End If
    Show = path
  End With
End Function

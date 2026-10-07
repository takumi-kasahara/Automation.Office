Attribute VB_Name = "Shape"
'@Folder("Macros")
'@ModuleDescription("Provides shape manipulation macros.")
Attribute VB_Description = "Provides shape manipulation macros."
Option Explicit

'@Description("Opens the resize picture dialog for the selected shape range.")
'@ExcelHotkey "Z"
Public Sub ResizePicture()
Attribute ResizePicture.VB_ProcData.VB_Invoke_Func = "Z\n14"
Attribute ResizePicture.VB_Description = "Opens the resize picture dialog for the selected shape range."
  If Not WorkbookExtensions.Backup Then Exit Sub

  With New ResizePictureDialog
    .Show
  End With
End Sub

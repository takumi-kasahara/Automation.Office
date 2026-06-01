'@Folder "Sheets"
'@ModuleDescription("Workbook document module for startup and save behaviors.")
Option Explicit

'@VariableDescription("Workbook-level event bridge used to handle backup events.")
Private BackupEvent As BackupEvent

'@Description("Initializes workbook event handlers when Personal.xlsb opens.")
'@Ignore ParameterNotUsed
Private Sub Workbook_Open()
  Set BackupEvent = New BackupEvent
  Set BackupEvent.App = Application
End Sub

'@Description("Toggles visibility of the Personal.xlsb workbook window.")
Public Sub ToggleHidden()
  With Application.Windows.Item(ThisWorkbook.name)
    .Visible = Not .Visible
  End With
End Sub

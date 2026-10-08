'@Folder "Sheets"
'@ModuleDescription("Workbook document module for startup and save behaviors.")
Attribute VB_Description = "Workbook document module for startup and save behaviors."
Option Explicit

'@VariableDescription("Workbook-level event bridge used to handle backup events.")
Private BackupEvent As BackupEvent
Attribute BackupEvent.VB_VarDescription = "Workbook-level event bridge used to handle backup events."

'@Description("Initializes workbook event handlers when Personal.xlsb opens.")
'@Ignore ParameterNotUsed
Private Sub Workbook_Open()
Attribute Workbook_Open.VB_Description = "Initializes workbook event handlers when Personal.xlsb opens."
  Set BackupEvent = New BackupEvent
  Set BackupEvent.App = Application
End Sub

'@Description("Toggles visibility of the Personal.xlsb workbook window.")
Public Sub ToggleHidden()
Attribute ToggleHidden.VB_Description = "Toggles visibility of the Personal.xlsb workbook window."
  With Application.Windows.Item(ThisWorkbook.name)
    .Visible = Not .Visible
  End With
End Sub

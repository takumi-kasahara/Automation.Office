VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BackupEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Utils.Events")
'@ModuleDescription("Handles workbook events used to trigger backup operations.")
Option Explicit

'@VariableDescription("Application event source bound during workbook startup.")
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

'@Description("Backs up visible workbooks before save and can cancel the save on failure.")
Private Sub App_WorkbookBeforeSave(ByVal this As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
  If Application.Windows.Item(this.name).Visible Then Cancel = Not WorkbookExtensions.Backup(this)
End Sub

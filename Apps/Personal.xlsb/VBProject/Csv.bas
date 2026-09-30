Attribute VB_Name = "Csv"
'@Folder("Macros")
'@ModuleDescription("Provides CSV import and export macros.")
Option Explicit

'@Description("Imports CSV data into the active worksheet.")
'@ExcelHotkey "I"
Public Sub ImportCsv()
Attribute ImportCsv.VB_ProcData.VB_Invoke_Func = "I\n14"
  With New CsvLoader
    .Import
  End With
End Sub

'@Description("Exports the selected range as CSV.")
'@ExcelHotkey "E"
Public Sub ExportCsv()
Attribute ExportCsv.VB_ProcData.VB_Invoke_Func = "E\n14"
  Dim rng As Range
  Set rng = WorksheetExtensions.GetSelectedRange
  If rng Is Nothing Then Exit Sub
  With New CsvCreator
    .Export rng
  End With
End Sub

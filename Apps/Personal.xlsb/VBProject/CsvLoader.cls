VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Utils.Classes")
'@ModuleDescription("Imports CSV files into the active worksheet as text columns.")
Option Explicit

'@Description("Loads a CSV file into the active worksheet and auto fits imported cells.")
Public Sub Import()
  Dim path As Variant
  path = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
  If VarType(path) = vbBoolean Then Exit Sub
  If Dir(path) = vbNullString Then Exit Sub
  Dim ws As Worksheet
  Set ws = Assumed.ActiveWorksheet
  If ws Is Nothing Then Exit Sub

  Dim fileColumnDataTypes() As Long
  With ws
    ReDim fileColumnDataTypes(0 To .Columns.Count - 1)
    Dim i As Long
    For i = 0 To .Columns.Count - 1
      fileColumnDataTypes(i) = XlColumnDataType.xlTextFormat
    Next
    Dim qt As QueryTable
    Set qt = .QueryTables.Add("TEXT;" & path, Application.ActiveCell)
    With qt
      .TextFilePlatform = 65001 ' UTF-8
      .TextFileParseType = xlDelimited
      .TextFileCommaDelimiter = True
      .TextFileColumnDataTypes = fileColumnDataTypes
      .Refresh
      .Delete
    End With
  End With

  Dim rng As Range
  Set rng = WorksheetExtensions.GetUsedRange
  If rng Is Nothing Then Exit Sub
  With rng
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
  End With
End Sub

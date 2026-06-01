VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Utils.Extensions")
'@PredeclaredId
'@ModuleDescription("Provides helpers for visible and selected worksheet ranges.")
Option Explicit

'@Description("Returns the visible used range of the active worksheet.")
Public Function GetUsedRange() As Range
  Dim ws As Worksheet
  Set ws = Assumed.ActiveWorksheet
  Set GetUsedRange = GetUsedRangeInWorksheet(ws)
End Function

'@Description("Returns the visible used range for the specified worksheet.")
Public Function GetUsedRangeInWorksheet(ByVal ws As Worksheet, Optional ByVal app As Object = Nothing) As Range
  If ws Is Nothing Then Exit Function

  On Error Resume Next
  Set GetUsedRangeInWorksheet = GetVisibleIn(ws.UsedRange, app)
  On Error GoTo 0
End Function

'@Description("Returns the visible selected range in the active worksheet.")
Public Function GetSelectedRange() As Range
  Dim ws As Worksheet
  Set ws = Assumed.ActiveWorksheet
  Set GetSelectedRange = GetSelectedRangeInWorksheet(ws, Application.Selection)
End Function

'@Description("Returns the visible selected range for the specified worksheet and selection.")
Public Function GetSelectedRangeInWorksheet(ByVal ws As Worksheet, ByVal selection As Object, Optional ByVal app As Object = Nothing) As Range
  If ws Is Nothing Then Exit Function
  If selection Is Nothing Then Exit Function
  If TypeName(selection) <> "Range" Then Exit Function

  If app Is Nothing Then Set app = Application

  On Error Resume Next
  Dim target As Range
  Set target = app.Intersect(ws.UsedRange, selection)
  Set GetSelectedRangeInWorksheet = GetVisibleIn(target, app)
  On Error GoTo 0
End Function

'@Description("Intersects the target range with visible cells only.")
Public Function GetVisibleIn(ByVal target As Range, Optional ByVal app As Object = Nothing) As Range
  If target Is Nothing Then Exit Function
  If app Is Nothing Then Set app = Application

  On Error Resume Next
  Set GetVisibleIn = app.Intersect(target, target.SpecialCells(xlCellTypeVisible))
  On Error GoTo 0
End Function

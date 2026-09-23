Attribute VB_Name = "WorkbookMacros"
'@Folder("Macros")
'@ModuleDescription("Provides workbook-level utility macros.")
Option Explicit

'@Description("Copies worksheet names in the active workbook to the clipboard.")
Public Sub CopySheetNames()
  Dim sheetNames As Collection
  Set sheetNames = New Collection
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    sheetNames.Add ws.name
  Next

  If Clipboard.Copy(CollectionExtensions.JoinByLine(sheetNames)) Then MsgBox "Copied to clipboard."
End Sub

'@Description("Resets worksheet and chart views and refreshes pivot tables.")
Public Sub ResetView()
  If Not WorkbookExtensions.Backup Then Exit Sub

  With New Optimizer
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
      ' Fold all groups.
      With ws
        If Not .Visible Then GoTo Continue
        .Outline.ShowLevels 1, 1

        Dim pt As PivotTable
        For Each pt In .PivotTables
          pt.PivotCache.Refresh

          Dim pf As PivotField
          For Each pf In pt.PivotFields
            PivotFieldExtensions.ForceHideDetails pf
          Next
        Next
Continue:
      End With
    Next

    Dim cht As Chart
    For Each cht In ActiveWorkbook.Charts
      With cht
        .PlotArea.Select
      End With
    Next
  End With
End Sub

'@Description("Moves each visible worksheet to a default zoom and cursor position.")
'@ExcelHotkey "Q"
Public Sub MoveToDefaultPosition()
Attribute MoveToDefaultPosition.VB_ProcData.VB_Invoke_Func = "Q\n14"
  If Not WorkbookExtensions.Backup Then Exit Sub

  With New Optimizer
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
      With ws
        If Not .Visible Then GoTo Continue
        .Activate
        ActiveWindow.Zoom = 100
        Application.GoTo .Cells.Item(1, 1), True
Continue:
      End With
    Next

    WorkbookExtensions.MoveToFirstSheet
  End With
End Sub

'@ExcelHotkey "K"
'@Description("Toggles worksheet outline summary orientation.")
Public Sub ToggleOutlineStyle()
Attribute ToggleOutlineStyle.VB_ProcData.VB_Invoke_Func = "K\n14"
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    With ws.Outline
      If .SummaryColumn = xlSummaryOnRight And .SummaryRow = xlSummaryBelow Then
        .SummaryColumn = xlSummaryOnLeft
        .SummaryRow = xlSummaryAbove
      Else
        .SummaryColumn = xlSummaryOnRight
        .SummaryRow = xlSummaryBelow
      End If
    End With
  Next
End Sub

'@Description("Reveals the active workbook location in Explorer or browser.")
'@ExcelHotkey "R"
Public Sub RevealInExplorer()
Attribute RevealInExplorer.VB_ProcData.VB_Invoke_Func = "R\n14"
  With ActiveWorkbook
    If .FullName Like "http://*" Or .FullName Like "https://*" Then
      Dim url As String
      url = ActiveWorkbook.FullName
      Dim pos As Long
      pos = InStrRev(url, "/")
      With New WshShell
        If pos > 0 Then
          .Run Left(url, InStrRev(url, "/") - 1)
        Else
          .Run url
        End If
      End With
    Else
      Shell "explorer.exe /select,""" & .FullName & """", vbNormalFocus
    End If
  End With
End Sub

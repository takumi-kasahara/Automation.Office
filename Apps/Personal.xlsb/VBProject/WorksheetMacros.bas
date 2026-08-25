Attribute VB_Name = "WorksheetMacros"
'@Folder("Macros")
'@ModuleDescription("Provides worksheet-level utility macros and hotkeys.")

Option Explicit

'@Description("Auto-fits used columns and rows on the active worksheet.")
'@ExcelHotkey "J"
Public Sub AdjustCells()
Attribute AdjustCells.VB_ProcData.VB_Invoke_Func = "J\n14"
  If Not WorkbookExtensions.Backup Then Exit Sub
  Dim rng As Range
  Set rng = WorksheetExtensions.GetUsedRange
  If rng Is Nothing Then Exit Sub

  With New Optimizer
    With rng
      Dim target As Range
      For Each target In .Areas
        target.EntireColumn.AutoFit
        target.EntireRow.AutoFit
      Next
    End With
  End With
End Sub

'@Description("Appends clipboard lines to cells from the active selection downward.")
'@ExcelHotkey "j"
Public Sub AppendLine()
Attribute AppendLine.VB_ProcData.VB_Invoke_Func = "j\n14"
  If Not WorkbookExtensions.Backup Then Exit Sub
  Dim ws As Worksheet
  Set ws = Assumed.ActiveWorksheet
  If ws Is Nothing Then Exit Sub

  Dim text As String
  text = Replace(Trim$(Clipboard.value), vbCr, vbNullString)
  If text = vbNullString Then Exit Sub
  Dim lines As Variant
  lines = Split(text, vbLf)

  Dim cell As Range
  Dim i As Long
  For i = 0 To UBound(lines)
    Set cell = ws.Cells(Application.ActiveCell.row + i, Application.ActiveCell.Column)
    If cell.HasFormula Then GoTo Continue
    If cell.Formula = vbNullString Then
      cell.value = lines(i)
    Else
      cell.value = cell.value & vbNewLine & lines(i)
    End If
Continue:
  Next
End Sub

'@Description("Copies the selected range addresses to the clipboard.")
'@ExcelHotkey "A"
Public Sub CopyAddress()
Attribute CopyAddress.VB_ProcData.VB_Invoke_Func = "A\n14"
  Dim rng As Range
  Set rng = WorksheetExtensions.GetSelectedRange
  If rng Is Nothing Then Exit Sub
  With rng
    Dim values As Collection
    Set values = New Collection

    Dim area As Range
    For Each area In .Areas
      values.Add area.Address(True, True, Application.ReferenceStyle)
    Next

    Clipboard.Copy CollectionExtensions.JoinByComma(values)
  End With
End Sub

'@Description("Copies formulas from the selected range to the clipboard.")
'@ExcelHotkey "F"
Public Sub CopyFormula()
Attribute CopyFormula.VB_ProcData.VB_Invoke_Func = "F\n14"
  Dim rng As Range
  Set rng = WorksheetExtensions.GetSelectedRange
  If rng Is Nothing Then Exit Sub
  With rng
    Dim formulas As Collection
    Set formulas = New Collection

    Dim area As Range
    For Each area In .Areas
      Dim cell As Range
      For Each cell In area
        If Trim$(cell.Formula) <> vbNullString Then
          With Application
            Select Case .ReferenceStyle
              Case XlReferenceStyle.xlA1
                formulas.Add Trim$(cell.Formula)
              Case XlReferenceStyle.xlR1C1
                formulas.Add Trim$(cell.FormulaR1C1)
            End Select
          End With
        End If
      Next

      If formulas.Count > 0 Then Clipboard.Copy CollectionExtensions.JoinByLine(formulas)
    Next
  End With
End Sub

'@Description("Copies print area dimensions and row or column sizes to the clipboard.")
Public Sub CopyPrintArea()
  Dim ws As Worksheet
  Set ws = Assumed.ActiveWorksheet
  If ws Is Nothing Then Exit Sub
  If ws.PageSetup.PrintArea = vbNullString Then
    MsgBox "PrintArea undefined."
    Exit Sub
  End If

  With ws.Range(ws.PageSetup.PrintArea)
    Dim text As String
    text = text & "Height = " & .Height & " pt (" & .rows.Count & " rows)" & vbNewLine
    Dim c As Range
    For Each c In .rows
      text = text & "R" & c.row & " = " & c.Height & " pt" & vbNewLine
    Next
    text = text & "Width = " & .Width & " pt (" & .Columns.Count & " cells)" & vbNewLine
    For Each c In .Columns
      text = text & "C" & c.Column & " = " & c.Width & " pt" & vbNewLine
    Next
    If Clipboard.Copy(text) Then MsgBox "PrintArea Infomation Copied."
  End With
End Sub

'@Description("Copies text from the current selection to the clipboard.")
'@ExcelHotkey "X"
Public Sub CopyText()
Attribute CopyText.VB_ProcData.VB_Invoke_Func = "X\n14"
  Debug.Print "Copy:" & vbTab & TypeName(Application.Selection)
  With Application.Selection
    Select Case TypeName(Application.Selection)
      Case "Range"
        With WorksheetExtensions.GetSelectedRange
          Dim texts As Collection
          Set texts = New Collection
          Dim area As Range
          For Each area In .Areas
            Dim cell As Range
            For Each cell In area
              If Trim$(cell.text) <> vbNullString Then
                texts.Add Trim$(cell.text)
              End If
            Next
            If texts.Count > 0 Then Clipboard.Copy CollectionExtensions.JoinByLine(texts)
          Next
        End With
      Case "Picture"
        .Copy
      Case "Nothing"
        Exit Sub
      Case Else
        Dim text As String
        text = SelectionExtensions.GetText
        If Trim$(text) <> vbNullString Then Clipboard.Copy Trim$(text)
    End Select
  End With
End Sub

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

' https://forest.watch.impress.co.jp/docs/serial/exceltips/1193072.html
'@Description("Fills blank cells in the selected range with the value from the previous row.")
'@ExcelHotkey "D"
Public Sub FillBlankCells()
Attribute FillBlankCells.VB_ProcData.VB_Invoke_Func = "D\n14"
  If Not WorkbookExtensions.Backup Then Exit Sub
  Dim rng As Range
  Set rng = WorksheetExtensions.GetSelectedRange
  If rng Is Nothing Then Exit Sub
  If MsgBox("Fill in the blank cells. Continue?", vbYesNo) = vbNo Then Exit Sub

  rng.SpecialCells(xlCellTypeBlanks).Select
  Application.CutCopyMode = False
  rng.FormulaR1C1 = "=R[-1]C"
End Sub

'@Description("Prompts and opens hyperlinks found in the used range.")
'@ExcelHotkey "H"
Public Sub FollowHyperlinks()
Attribute FollowHyperlinks.VB_ProcData.VB_Invoke_Func = "H\n14"
  Dim rng As Range
  Dim link As Hyperlink
  For Each rng In WorksheetExtensions.GetUsedRange
    For Each link In rng.Hyperlinks
      Select Case MsgBox("Open " & link.Address & " at " & rng.Address & " ?", vbYesNoCancel)
        Case vbYes
          link.Follow
        Case vbNo
          GoTo Continue
        Case Else
          Exit Sub
      End Select
Continue:
    Next
  Next
End Sub

'@Description("Opens the resize picture dialog for the selected shape range.")
'@ExcelHotkey "Z"
Public Sub ResizePicture()
Attribute ResizePicture.VB_ProcData.VB_Invoke_Func = "Z\n14"
  If Not WorkbookExtensions.Backup Then Exit Sub

  With New ResizePictureDialog
    .Show
  End With
End Sub

' https://www.relief.jp/docs/excel-vba-unmerge-and-fill.html
'@Description("Unmerges selected merged cells and preserves their values.")
'@ExcelHotkey "M"
Public Sub UnMergeCells()
Attribute UnMergeCells.VB_ProcData.VB_Invoke_Func = "M\n14"
  If Not WorkbookExtensions.Backup Then Exit Sub

  With New Optimizer
    Dim rng As Range
    For Each rng In WorksheetExtensions.GetSelectedRange
      If rng.MergeCells Then
        Dim value As Variant
        value = rng.value
        With rng.MergeArea
          .UnMerge
          .value = value
        End With
      End If
    Next
  End With
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Utils.Classes")
'@ModuleDescription("Builds and exports CSV data from selected worksheet ranges.")
Option Explicit

'@Description("Exports the target range as CSV rows.")
Public Sub Export(ByVal target As Range)
  Dim columnIndices As Collection
  Set columnIndices = New Collection
  With target
    Dim i As Long
    For i = 1 To .Columns.Count
      If Trim$(.Cells.Item(1, i)) = vbNullString Then GoTo Continue
      columnIndices.Add i
Continue:
    Next
  End With

  With target
    Dim rs As Collection
    Set rs = New Collection
    Dim area As Range
    For Each area In .Areas
      Dim row As Range
      For Each row In area.rows
        Dim cs As Collection
        Set cs = New Collection
        Dim j As Variant
        For Each j In columnIndices
          cs.Add """" & Trim$(row.Cells.Item(1, j)) & """"
        Next
        rs.Add CollectionExtensions.JoinByComma(cs)
      Next
    Next
  End With

  Save rs, target.Worksheet.name
End Sub

'@Description("Prompts for destination path and saves CSV rows to a file.")
Private Sub Save(ByVal rs As Collection, ByVal InitialFileName As String)
  On Error GoTo Finally

  Dim path As String
  With New SaveFileDialog
    .FileFilter = "CSV File (*.csv), *.csv"
    .InitialFileName = InitialFileName
    path = .Show
    If path = vbNullString Then Exit Sub
  End With

  With New FileSystemObject
    Dim ts As TextStream
    Set ts = .CreateTextFile(path)
    Dim r As Variant
    For Each r In rs
      ' nbsp (U+00A0) to half-width space.
      ts.WriteLine Replace(r, ChrW$(&HA0), " ")
    Next
Finally:
    If Not ts Is Nothing Then ts.Close
    If Err.Number <> 0 Then Exception.Show
  End With
End Sub

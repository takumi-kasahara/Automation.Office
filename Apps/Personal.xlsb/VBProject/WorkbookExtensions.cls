VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Utils.Extensions")
'@ModuleDescription("Provides workbook helper functions used across macros.")
Option Explicit

'@Description("Creates a timestamped backup copy of a workbook.")
Public Function Backup(Optional ByVal wb As Workbook) As Boolean
  On Error GoTo Catch

  Backup = False
  If wb Is Nothing Then Set wb = Application.ActiveWorkbook

  With New FileSystemObject
    Dim fileName As String
    fileName = .GetBaseName(wb.name)
    Dim ext As String
    ext = .GetExtensionName(wb.name)
  End With

  If ext = vbNullString Then
    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename
    If VarType(filePath) = vbBoolean Then Exit Function
    wb.SaveAs filePath
    Backup = True
    Exit Function
  End If

  Dim bakPath As String
  With New FileSystemObject
    bakPath = .BuildPath(Environ("TEMP"), fileName & "." & Format$(Now, "yyyymmddhhmmss") & "." & ext)
  End With
  Debug.Print "Saved:" & vbTab & bakPath
  wb.SaveCopyAs bakPath
  Backup = True
  Exit Function
Catch:
  Exception.Show
End Function

'@Description("Activates the first visible worksheet in the active workbook.")
Public Sub MoveToFirstSheet()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    If ws.Visible Then
      ws.Activate
      Exit Sub
    End If
  Next
End Sub

'@Description("Finds a worksheet by name in the active workbook.")
Public Function FindWorksheet(ByVal name As String) As Worksheet
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    If ws.name = name Then
      Set FindWorksheet = ws
      Exit Function
    End If
  Next
End Function

'@Description("Finds a table by name across all worksheets in the active workbook.")
Public Function FindListObject(ByVal name As String) As ListObject
  Dim ws As Worksheet
  Dim tbl As ListObject
  For Each ws In ActiveWorkbook.Worksheets
    For Each tbl In ws.ListObjects
      If tbl.name = name Then
        Set FindListObject = tbl
        Exit Function
      End If
    Next
  Next
End Function

'@Description("Returns the base name part from a qualified workbook identifier.")
Public Function GetBaseName(ByVal name As String) As String
  Dim pos As Long
  pos = InStr(1, name, "!")
  If pos > 0 Then
    GetBaseName = Mid$(name, pos + 1)
  Else
    GetBaseName = name
  End If
End Function

Attribute VB_Name = "SelectionExtensionsTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
  ' this method runs once per module.
  Set Assert = New Rubberduck.AssertClass
  Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
  ' this method runs once per module.
  Set Assert = Nothing
  Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
  ' this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
  ' this method runs after every test in the module.
End Sub

'@TestMethod("TextFrom")
Private Sub TextFromReturnsTextFromMockedSelection()
  On Error GoTo TestFail

  ' Arrange:
  Dim wb As Workbook
  Set wb = Application.Workbooks.Add

  Dim selection As Object
  Set selection = wb.Worksheets(1).Range("A1")
  selection.Value = "MockedText"

  ' Act:
  Dim actual As String
  actual = SelectionExtensions.TextFrom(selection)

  ' Assert:
  Assert.AreEqual "MockedText", actual, "TextFrom should return text provided by the selection object."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  If Not wb Is Nothing Then wb.Close SaveChanges:=False
  Set selection = Nothing
  Set wb = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("ShapeRangeFrom")
Private Sub ShapeRangeFromReturnsShapeRangeFromMockedSelection()
  On Error GoTo TestFail

  ' Arrange:
  Dim wb As Workbook
  Set wb = Application.Workbooks.Add
  Dim ws As Worksheet
  Set ws = wb.Worksheets(1)

  Dim selection As Object
  Set selection = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 40, 20)

  ' Act:
  Dim actual As ShapeRange
  Set actual = SelectionExtensions.ShapeRangeFrom(selection)

  ' Assert:
  Assert.IsFalse actual Is Nothing, "ShapeRangeFrom should return shape range provided by selection object."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  If Not wb Is Nothing Then wb.Close SaveChanges:=False
  Set actual = Nothing
  Set selection = Nothing
  Set ws = Nothing
  Set wb = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("TextFrom")
Private Sub TextFromReturnsEmptyStringWhenSelectionIsNothing()
  On Error GoTo TestFail

  ' Arrange:
  Dim selection As Object
  Set selection = Nothing

  ' Act:
  Dim actual As String
  actual = SelectionExtensions.TextFrom(selection)

  ' Assert:
  Assert.AreEqual vbNullString, actual, "TextFrom should return empty string when selection is Nothing."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set selection = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

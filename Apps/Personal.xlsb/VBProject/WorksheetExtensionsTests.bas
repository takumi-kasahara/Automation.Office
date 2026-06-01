Attribute VB_Name = "WorksheetExtensionsTests"
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

'@TestMethod("GetVisibleIn")
Private Sub GetVisibleInReturnsIntersectedVisibleRange()
  On Error GoTo TestFail

  ' Arrange:
  Dim wb As Workbook
  Set wb = Application.Workbooks.Add
  Dim ws As Worksheet
  Set ws = wb.Worksheets(1)
  ws.Range("A1").Value = "Hello"
  ws.Range("B1").Value = "World"

  Dim target As Range
  Set target = ws.Range("A1:B1")

  ' Act:
  Dim actual As Range
  Set actual = WorksheetExtensions.GetVisibleIn(target)

  ' Assert:
  Assert.IsFalse actual Is Nothing, "GetVisibleIn should return the intersected visible range."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set actual = Nothing
  Set target = Nothing
  Set ws = Nothing
  wb.Close SaveChanges:=False
  Set wb = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetSelectedRangeInWorksheet")
Private Sub GetSelectedRangeInWorksheetReturnsVisibleSelectedRange()
  On Error GoTo TestFail

  ' Arrange:
  Dim wb As Workbook
  Set wb = Application.Workbooks.Add
  Dim ws As Worksheet
  Set ws = wb.Worksheets(1)
  ws.Range("A1").Value = "Hello"
  ws.Range("B1").Value = "World"

  Dim selection As Range
  Set selection = ws.Range("A1:B1")

  ' Act:
  Dim actual As Range
  Set actual = WorksheetExtensions.GetSelectedRangeInWorksheet(ws, selection)

  ' Assert:
  Assert.IsFalse actual Is Nothing, "GetSelectedRangeInWorksheet should return visible cells from selected used range intersection."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set actual = Nothing
  Set selection = Nothing
  Set ws = Nothing
  wb.Close SaveChanges:=False
  Set wb = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetSelectedRangeInWorksheet")
Private Sub GetSelectedRangeInWorksheetReturnsNothingWhenSelectionIsNotRange()
  On Error GoTo TestFail

  ' Arrange:
  Dim wb As Workbook
  Set wb = Application.Workbooks.Add
  Dim ws As Worksheet
  Set ws = wb.Worksheets(1)

  Dim notRangeSelection As Object
  Set notRangeSelection = CreateObject("Scripting.Dictionary")

  ' Act:
  Dim actual As Range
  Set actual = WorksheetExtensions.GetSelectedRangeInWorksheet(ws, notRangeSelection)

  ' Assert:
  Assert.IsTrue actual Is Nothing, "GetSelectedRangeInWorksheet should return Nothing when selection is not a range."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set actual = Nothing
  Set notRangeSelection = Nothing
  Set ws = Nothing
  wb.Close SaveChanges:=False
  Set wb = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

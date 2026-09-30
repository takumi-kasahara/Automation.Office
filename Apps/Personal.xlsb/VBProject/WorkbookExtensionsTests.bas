Attribute VB_Name = "WorkbookExtensionsTests"
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

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsNameAfterExclamationMark()
  On Error GoTo TestFail

  Assert.AreEqual "Sheet1", WorkbookExtensions.GetBaseName("Personal.xlsb!Sheet1"), _
    "GetBaseName should return the part after the exclamation mark."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsInputAsIsWhenNoExclamationMark()
  On Error GoTo TestFail

  Assert.AreEqual "Sheet1", WorkbookExtensions.GetBaseName("Sheet1"), _
    "GetBaseName should return the input unchanged when there is no exclamation mark."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsEmptyStringWhenInputIsEmpty()
  On Error GoTo TestFail

  Assert.AreEqual vbNullString, WorkbookExtensions.GetBaseName(vbNullString), _
    "GetBaseName should return an empty string for empty input."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsEmptyStringWhenExclamationMarkIsAtEnd()
  On Error GoTo TestFail

  Assert.AreEqual vbNullString, WorkbookExtensions.GetBaseName("Personal.xlsb!"), _
    "GetBaseName should return an empty string when exclamation mark is at the end."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsPartAfterFirstExclamationMarkOnly()
  On Error GoTo TestFail

  Assert.AreEqual "Sheet1!Extra", WorkbookExtensions.GetBaseName("Personal.xlsb!Sheet1!Extra"), _
    "GetBaseName should split only on the first exclamation mark."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsSuffixWhenNameStartsWithExclamationMark()
  On Error GoTo TestFail

  Assert.AreEqual "Sheet1", WorkbookExtensions.GetBaseName("!Sheet1"), _
    "GetBaseName should return text after the first exclamation mark even when it is at the beginning."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsOriginalWhenNameContainsFullWidthExclamationMark()
  On Error GoTo TestFail

  Assert.AreEqual "Personal.xlsb！Sheet1", WorkbookExtensions.GetBaseName("Personal.xlsb！Sheet1"), _
    "GetBaseName should not split when a full-width exclamation mark is used."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("GetBaseName")
Private Sub GetBaseNameReturnsWhitespaceSegmentUnchanged()
  On Error GoTo TestFail

  Assert.AreEqual " Sheet 1 ", WorkbookExtensions.GetBaseName("Personal.xlsb! Sheet 1 "), _
    "GetBaseName should preserve whitespace in the suffix segment."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

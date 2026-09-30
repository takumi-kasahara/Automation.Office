Attribute VB_Name = "PivotFieldExtensionsTests"
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

'@TestMethod("ForceClearAllFilters")
Private Sub ForceClearAllFiltersDoesNotRaiseErrorWhenPivotFieldIsNothing()
  On Error GoTo TestFail

  Dim target As PivotField
  Set target = Nothing

  PivotFieldExtensions.ForceClearAllFilters target

  Assert.Succeed

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("ForceHideDetails")
Private Sub ForceHideDetailsDoesNotRaiseErrorWhenPivotFieldIsNothing()
  On Error GoTo TestFail

  Dim target As PivotField
  Set target = Nothing

  PivotFieldExtensions.ForceHideDetails target

  Assert.Succeed

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

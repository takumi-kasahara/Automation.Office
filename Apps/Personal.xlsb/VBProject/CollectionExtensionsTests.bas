Attribute VB_Name = "CollectionExtensionsTests"
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
  ' this method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
  ' this method runs after every test in the module.
End Sub

'@TestMethod("Contains")
Private Sub ContainsReturnsTrueWhenKeyExists()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add Item:="A", key:="A"

  Assert.IsTrue CollectionExtensions.Contains(target, "A"), "Contains should return True for an existing key."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("Contains")
Private Sub ContainsReturnsFalseWhenKeyDoesNotExist()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add Item:="A", key:="A"

  Assert.IsFalse CollectionExtensions.Contains(target, "B"), "Contains should return False for a missing key."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("Contains")
Private Sub ContainsReturnsFalseWhenTargetDoesNotSupportItemLookup()
  On Error GoTo TestFail

  Dim target As Object
  Set target = CreateObject("Scripting.Dictionary")
  target.Add "A", "A"

  Assert.IsFalse CollectionExtensions.Contains(target, "A"), _
    "Contains should return False when target does not support default Item call semantics."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub ToArrayReturnsNothingForEmptyCollection()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection

  Dim values() As Variant
  values = CollectionExtensions.ToArray(target)

  Assert.IsNothing values, "ToArray should returns nothing for an empty collection."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Erase values
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub ToArrayReturnsSingleItemArray()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "Only"

  Dim values() As Variant
  values = CollectionExtensions.ToArray(target)

  Assert.AreEqual 1&, UBound(values) - LBound(values) + 1, "ToArray should return exactly one element for a single-item collection."
  Assert.AreEqual "Only", CStr(values(0)), "Single value should be preserved."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Erase values
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub ToArrayReturnsItemsInOrder()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "A"
  target.Add "B"
  target.Add "C"

  Dim values() As Variant
  values = CollectionExtensions.ToArray(target)

  Assert.AreEqual 3&, UBound(values) - LBound(values) + 1, "ToArray should return an array with the same item count."
  Assert.AreEqual "A", CStr(values(0)), "First item should match the source collection."
  Assert.AreEqual "B", CStr(values(1)), "Second item should match the source collection."
  Assert.AreEqual "C", CStr(values(2)), "Third item should match the source collection."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Erase values
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("JoinBy")
Private Sub JoinByJoinsWithSpecifiedDelimiter()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "A"
  target.Add "B"
  target.Add "C"

  Assert.AreEqual "A|B|C", CollectionExtensions.JoinBy(target, "|"), "JoinBy should join values with the given delimiter."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("JoinBy")
Private Sub JoinByReturnsSingleValueWithoutDelimiter()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "Single"

  Assert.AreEqual "Single", CollectionExtensions.JoinBy(target, "|"), _
    "JoinBy should return the single element as-is."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("JoinByComma")
Private Sub JoinByCommaUsesCommaDelimiter()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "A"
  target.Add "B"
  target.Add "C"

  Assert.AreEqual "A,B,C", CollectionExtensions.JoinByComma(target), "JoinByComma should use a comma delimiter."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("JoinByTab")
Private Sub JoinByTabUsesTabDelimiter()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "A"
  target.Add "B"
  target.Add "C"

  Assert.AreEqual "A" & vbTab & "B" & vbTab & "C", CollectionExtensions.JoinByTab(target), "JoinByTab should use vbTab as delimiter."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

'@TestMethod("JoinByLine")
Private Sub JoinByLineUsesLineBreakDelimiter()
  On Error GoTo TestFail

  Dim target As Collection
  Set target = New Collection
  target.Add "A"
  target.Add "B"
  target.Add "C"

  Assert.AreEqual "A" & vbNewLine & "B" & vbNewLine & "C", CollectionExtensions.JoinByLine(target), "JoinByLine should use vbNewLine as delimiter."

TestExit:
  '@Ignore UnhandledOnErrorResumeNext
  On Error Resume Next
  Set target = Nothing
  Exit Sub
TestFail:
  Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
  Resume TestExit
End Sub

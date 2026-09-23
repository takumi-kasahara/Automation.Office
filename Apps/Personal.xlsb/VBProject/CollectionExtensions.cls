VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Utils.Extensions")
'@ModuleDescription("Provides helper functions for collection-like objects.")
Option Explicit

'@Description("Returns True when the specified key exists in the collection-like object.")
Public Function Contains(ByVal this As Object, ByVal key As Variant) As Boolean
  On Error GoTo Catch

  ' Force Error
  this.Item key
  Contains = True
  Exit Function
Catch:
  Contains = False
End Function

' https://stackoverflow.com/questions/48153398/converting-vba-collection-to-array
'@Description("Converts a collection-like object to a Variant array.")
Public Function ToArray(ByVal this As Object) As Variant()
  Dim values() As Variant
  If this.Count = 0 Then Exit Function
  ReDim values(0 To this.Count - 1)
  Dim index As Long
  For index = 1 To this.Count
    values(index - 1) = this.Item(index)
  Next
  ToArray = values
End Function

'@Description("Joins all items using the specified delimiter.")
Public Function JoinBy(ByVal this As Object, ByVal delimiter As String) As String
  JoinBy = Join(Me.ToArray(this), delimiter)
End Function

'@Description("Joins all items with commas.")
Public Function JoinByComma(ByVal this As Object) As String
  JoinByComma = Me.JoinBy(this, ",")
End Function

''@Description("Joins all items with tab characters.")
Public Function JoinByTab(ByVal this As Object) As String
  JoinByTab = Me.JoinBy(this, vbTab)
End Function

'@Description("Joins all items with line breaks.")
Public Function JoinByLine(ByVal this As Object) As String
  JoinByLine = Me.JoinBy(this, vbNewLine)
End Function

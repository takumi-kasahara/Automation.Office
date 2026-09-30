VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Utils.Extensions")
'@PredeclaredId
'@ModuleDescription("Provides safe helpers for reading selection state.")
Option Explicit

'@Description("Returns ShapeRange from the current selection when available.")
Public Property Get GetShapeRange() As ShapeRange
  Set GetShapeRange = ShapeRangeFrom(Application.Selection)
End Property

'@Description("Returns ShapeRange from the specified selection object when available.")
Public Property Get ShapeRangeFrom(ByVal selection As Object) As ShapeRange
  If selection Is Nothing Then Exit Property

  On Error Resume Next
  Set ShapeRangeFrom = selection.ShapeRange
  On Error GoTo 0
End Property

'@Description("Returns text from the current selection when available.")
Public Property Get GetText() As String
  GetText = TextFrom(Application.Selection)
End Property

'@Description("Returns text from the specified selection object when available.")
Public Property Get TextFrom(ByVal selection As Object) As String
  If selection Is Nothing Then Exit Property

  On Error Resume Next
  TextFrom = selection.text
  On Error GoTo 0
End Property

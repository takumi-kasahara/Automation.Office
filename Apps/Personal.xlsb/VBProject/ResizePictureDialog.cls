VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResizePictureDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utils.Dialogs"
'@ModuleDescription("Shows a dialog workflow to scale selected shapes.")
Option Explicit

'@Description("Prompts for a scale percentage and resizes selected shapes.")
Public Sub Show()
  Dim sr As ShapeRange
  Set sr = SelectionExtensions.GetShapeRange
  If sr Is Nothing Then Exit Sub
  With sr
    Do
      Dim percent As Variant
      percent = Application.InputBox("Enter scale (%).", Default:=100, Type:=1)
      If VarType(percent) = vbBoolean Then Exit Sub
    Loop Until percent > 0

    .ScaleHeight (percent / 100), msoTrue, msoScaleFromTopLeft
    .ScaleWidth (percent / 100), msoTrue, msoScaleFromTopLeft
    With .line
      .Transparency = 0
      .Visible = msoTrue
      .Weight = 1
      With .ForeColor
        .Brightness = 0
        .ObjectThemeColor = msoThemeColorAccent1
        .TintAndShade = 0
      End With
    End With
  End With
End Sub

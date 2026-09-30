Attribute VB_Name = "ApplicationMacros"
'@Folder("Macros")
'@ModuleDescription("Provides application-level Excel utility macros.")
Option Explicit

'@Description("Toggles full screen mode for the Excel application window.")
Public Sub ToggleFullScreen()
  With Application
    .DisplayFullScreen = Not .DisplayFullScreen
  End With
End Sub

' https://amacoda.net/blog/2018/09/toggle_reference_style_a1_to_from_r1c1/
'@Description("Toggles cell reference style between A1 and R1C1.")
'@ExcelHotkey "T"
Public Sub ToggleReferenceStyle()
Attribute ToggleReferenceStyle.VB_ProcData.VB_Invoke_Func = "T\n14"
  With Application
    Select Case .ReferenceStyle
      Case XlReferenceStyle.xlA1
        .ReferenceStyle = xlR1C1
      Case XlReferenceStyle.xlR1C1
        .ReferenceStyle = xlA1
    End Select
  End With
End Sub

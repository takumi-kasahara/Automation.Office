Attribute VB_Name = "WindowMacros"
'@Folder("Macros")
'@ModuleDescription("Provides window-level view switching macros.")
Attribute VB_Description = "Provides window-level view switching macros."
Option Explicit

'@Description("Cycles the active window through normal, page break, and page layout views.")
'@ExcelHotkey "P"
Public Sub ToggleView()
Attribute ToggleView.VB_ProcData.VB_Invoke_Func = "P\n14"
Attribute ToggleView.VB_Description = "Cycles the active window through normal, page break, and page layout views."
  Select Case ActiveWindow.View
    Case XlWindowView.xlNormalView
      ActiveWindow.View = xlPageBreakPreview
    Case XlWindowView.xlPageBreakPreview
      ActiveWindow.View = xlPageLayoutView
    Case XlWindowView.xlPageLayoutView
      ActiveWindow.View = xlNormalView
  End Select
End Sub

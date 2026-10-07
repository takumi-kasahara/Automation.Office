VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotFieldExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Utils.Extensions")
'@ModuleDescription("Provides safe helper operations for PivotField objects.")
Attribute VB_Description = "Provides safe helper operations for PivotField objects."
Option Explicit

'@Description("Clears all pivot filters without raising an error.")
Public Sub ForceClearAllFilters(ByVal this As PivotField)
Attribute ForceClearAllFilters.VB_Description = "Clears all pivot filters without raising an error."
  On Error Resume Next
  this.ClearAllFilters
End Sub

'@Description("Forces a pivot field to hide detail rows without raising an error.")
Public Sub ForceHideDetails(ByVal this As PivotField)
Attribute ForceHideDetails.VB_Description = "Forces a pivot field to hide detail rows without raising an error."
  On Error Resume Next
  this.ShowDetail = False
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Optimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Utils.Classes"
'@ModuleDescription("Temporarily applies performance-friendly Excel application settings.")
Option Explicit

Private Optimizing As Boolean

Private Sub Class_Initialize()
  If Optimizing Then Exit Sub

  With Application
    .calculation = xlCalculationManual
    .Cursor = xlWait
    .DisplayAlerts = False
    .EnableEvents = False
    .ScreenUpdating = False
  End With

  Optimizing = True
End Sub

Private Sub Class_Terminate()
  If Not Optimizing Then Exit Sub

  With Application
    .calculation = xlCalculationAutomatic
    .Cursor = xlDefault
    .DisplayAlerts = True
    .EnableEvents = True
    .ScreenUpdating = True
  End With

  Optimizing = False
End Sub

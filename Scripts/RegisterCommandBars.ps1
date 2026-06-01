using namespace Microsoft.Office.Core
using namespace System.Runtime.InteropServices

# https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.msobuttonstyle?view=office-pia
[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$app = New-Object -ComObject Excel.Application
try {
  $app.VBE.CommandBars |
  ForEach-Object {
    $_.Reset()
    $_.Controls |
    ForEach-Object {
      try {
        $_.Reset()
      }
      catch [COMException] {
        Write-Warning -Message $_.Exception.Message
      }
      $control = $_
      switch ($_.Id) {
        192 {
          $control.Caption = '(&/)'
          $control.Style = [MsoButtonStyle]::msoButtonIconAndCaption
        }
        2552 {
          $control.Caption = '(&\)'
          $control.Style = [MsoButtonStyle]::msoButtonIconAndCaption
        }
        2525 {
          $control.Caption = '(&;)'
          $control.Style = [MsoButtonStyle]::msoButtonIconAndCaption
        }
        2526 {
          $control.Caption = '(&.)'
          $control.Style = [MsoButtonStyle]::msoButtonIconAndCaption
        }
        2527 {
          $control.Caption = '(&,)'
          $control.Style = [MsoButtonStyle]::msoButtonIconAndCaption
        }
      }
    }
  }
}
finally {
  try {
    if ($app) {
      $app.Quit()
    }
  }
  finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

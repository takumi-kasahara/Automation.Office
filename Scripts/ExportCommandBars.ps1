<#
.SYNOPSIS
  Exports CommandBar configurations from Office applications to CSV files.

.DESCRIPTION
  This script exports CommandBar configurations from Access, Excel, PowerPoint, and Word applications to CSV files.
  It creates COM objects for each application, retrieves the CommandBars and their controls, and exports the data to CSV files in the `./Out` directory.
  For Access, it also exports VBA CommandBars if available.

.NOTES
  The script creates the `./Out` directory if it doesn't exist.
#>
[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$out = './Out'
$apps = [ordered]@{
  Access     = @{
    Id  = 'Access.Application'
    VBE = $true
  }
  Excel      = @{
    Id  = 'Excel.Application'
    VBE = $false
  }
  PowerPoint = @{
    Id  = 'PowerPoint.Application'
    VBE = $false
  }
  Word       = @{
    Id  = 'Word.Application'
    VBE = $false
  }
}
$apps.GetEnumerator() |
ForEach-Object {
  Write-Progress -Activity "Exporting CommandBars for $($_.Key)"
  $app = New-Object -ComObject $_.Value.Id
  try {
    $csv = $out | Join-Path -ChildPath "CommandBars.$($_.Key).csv"
    if (Test-Path -LiteralPath $csv) {
      Remove-Item -LiteralPath $csv
    }
    $commandBars = $app.CommandBars |
    ForEach-Object {
      $commandBar = $_
      $_.Reset()
      $_.Controls |
      ForEach-Object {
        [PSCustomObject]@{
          CommandBarId   = $commandBar.Id
          CommandBarName = $commandBar.Name
          ControlId      = $_.Id
          Caption        = $_.Caption
          TooltipText    = $_.TooltipText
        }
      }
    }
    $commandBars |
    Sort-Object -Property CommandBarId, ControlId |
    Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

    if (-not $_.Value.VBE) {
      return
    }
    $csv = $out | Join-Path -ChildPath 'CommandBars.csv'
    if (Test-Path -LiteralPath $csv) {
      Remove-Item -LiteralPath $csv
    }
    $commandBars = $app.VBE.CommandBars |
    ForEach-Object {
      $commandBar = $_
      $_.Reset()
      $_.Controls |
      ForEach-Object {
        [PSCustomObject]@{
          CommandBarId   = $commandBar.Id
          CommandBarName = $commandBar.Name
          ControlId      = $_.Id
          Caption        = $_.Caption
          TooltipText    = $_.TooltipText
        }
      }
    }
    $commandBars |
    Sort-Object -Property CommandBarId, ControlId |
    Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
  } finally {
    try {
      if ($app) {
        $app.Quit()
      }
    } finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
      Write-Progress -Activity "Exporting CommandBars for $($_.Key)" -Completed
    }
  }
}

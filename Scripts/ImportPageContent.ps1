<#
.SYNOPSIS
  Imports OneNote page content from XML files.

.DESCRIPTION
  This script imports OneNote page content from XML files into the OneNote hierarchy.

.NOTES
  The script expects a Config.psd1 file with an Output property specifying the source directory.
#>
using module ..\Modules\Automation.Office.psd1

[CmdletBinding()]
param ()
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = Import-PowerShellDataFile -LiteralPath 'Config.psd1'
if ($null -eq $config.Output) {
  throw [InvalidOperationException]::new('Output is required.')
}
$output = $config.Output | Join-Path -ChildPath 'Page'
if (-not (Test-Path -LiteralPath $output)) {
  throw [DirectoryNotFoundException]::new("$output does not exist.")
}

Compress-Archive -Path "$output/*" -Destination "$output.$(Get-Date -Format 'yyyyMMddHHmmss').zip"

$activity = 'Importing PageContents'
try {
  Get-ChildItem -LiteralPath $output -Filter '*.xml' |
  ForEach-Object {
    Write-Progress -Activity $activity -Status "Importing: $($_.Name)"
    Import-OneNotePageContent -LiteralPath $_.FullName
    "Imported: $($_.Name)" | Out-Host
  }
} finally {
  Write-Progress -Activity $activity -Completed
}

& .\ExportPageContent.ps1

<#
.SYNOPSIS
  Exports binary objects from OneNote pages to files.

.DESCRIPTION
  This script exports binary objects (attachments, images, etc.) from OneNote pages to a specified output directory.
  It uses the Get-OneNoteHierarchy cmdlet to retrieve the page hierarchy and then exports binary objects from each page.

.NOTES
  The script expects a Config.psd1 file with Output and StartNodeId properties.
#>
using module .\Modules\Automation.Office.psd1
using namespace System.Xml

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$configPath = 'Config.psd1'
if (-not (Test-Path -LiteralPath $configPath)) {
  Copy-Item -LiteralPath 'Config.tmpl.psd1' -Destination $configPath
}
$config = Import-PowerShellDataFile -LiteralPath $configPath
if ($null -eq $config.Output) {
  throw [InvalidOperationException]::new('Output is required.')
}
if (-not (Test-Path -LiteralPath $config.Output)) {
  New-Item -Path $config.Output -ItemType Directory | Out-Null
}
$output = $config.Output | Join-Path -ChildPath 'Bin'
if (Test-Path -LiteralPath $output) {
  Get-ChildItem -LiteralPath $output | Remove-Item
} else {
  New-Item -Path $output -ItemType Directory | Out-Null
}

$hierarchies = Get-OneNoteHierarchy -Id $config.StartNodeId -HierarchyScope hsPages
$hierarchies.SelectNodes('//*') |
Where-Object {
  $_ -is [XmlElement] -and
  $_.HasAttribute('ID') -and
  (-not $_.HasAttribute('isInRecycleBin') -or -not [bool]::Parse($_.GetAttribute('isInRecycleBin'))) -and
  $_.LocalName -eq 'Page'
} |
ForEach-Object {
  $destination = Resolve-Path -LiteralPath $output
  Export-OneNoteBinaryObject -Id $_.GetAttribute('ID') -Destination $destination
}
if (@(Get-Item -Path "$output/*").Count -gt 0) {
  Compress-Archive -Path "$output/*" -Destination "$output.$(Get-Date -Format 'yyyyMMddHHmmss').zip"
}

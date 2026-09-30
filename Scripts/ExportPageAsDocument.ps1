<#
.SYNOPSIS
  Exports OneNote pages as documents in various formats.

.DESCRIPTION
  This script exports OneNote pages as documents in multiple formats (OneNote, OneNote Package, MHT, PDF, XPS, DOCX, etc.).

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
$output = $config.Output | Join-Path -ChildPath 'Publish'
if (-not (Test-Path -LiteralPath $output)) {
  New-Item -Path $output -ItemType Directory | Out-Null
}
Get-ChildItem -LiteralPath $output -Filter '*.one' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.onepkg' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.mht' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.pdf' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.xps' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.docx' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.wbk' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.emf' | Remove-Item
Get-ChildItem -LiteralPath $output -Filter '*.html' | Remove-Item

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
  Export-OneNotePageAsDocument -Id $_.GetAttribute('ID') -Destination $destination -PublishFormat pfOneNote
}
if (@(Get-Item -Path "$output/*").Count -gt 0) {
  Compress-Archive -Path "$output/*" -Destination "$output.$(Get-Date -Format 'yyyyMMddHHmmss').zip"
}

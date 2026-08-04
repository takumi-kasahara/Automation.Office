using namespace System.IO

[CmdletBinding()]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Modules\Automation.Office.psd1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = Import-PowerShellDataFile -LiteralPath 'Config.psd1'
if ($null -eq $config.Output) {
  throw [InvalidOperationException]::new('Output is required.')
}
$source = $config.Output | Join-Path -ChildPath 'Hierarchy'
if (-not (Test-Path -LiteralPath $source)) {
  throw [DirectoryNotFoundException]::new("$source does not exist.")
}

Compress-Archive -Path "$source/*" -Destination "$source.$(Get-Date -Format 'yyyyMMddHHmmss').zip"
Get-ChildItem -LiteralPath $source -Filter '*.xml' |
ForEach-Object {
  "Import:`t$($_.FullName)" | Out-Host
  Import-OneNoteHierarchy -LiteralPath $_.FullName
}

& .\ExportHierarchy.ps1

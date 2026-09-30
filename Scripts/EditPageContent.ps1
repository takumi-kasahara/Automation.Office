<#
.SYNOPSIS
  Edits page content in OneNote documents by removing language attributes.

.DESCRIPTION
  This script processes OneNote XML files to remove language attributes from Page elements and OE elements.
  It extracts the XML files from a compressed archive, modifies them to remove language attributes, and then recreates the archive with the updated files.

.NOTES
  The script expects a Config.psd1 file in the same directory with an Output property specifying the base directory for processing.
#>
using module .\Modules\Automation.Office.psd1
using namespace System.IO
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
$source = $config.Output | Join-Path -ChildPath 'Page'
if (-not (Test-Path -LiteralPath $source)) {
  throw [DirectoryNotFoundException]::new("$source does not exist.")
}

Compress-Archive -Path "$source/*" -Destination "$source.$(Get-Date -Format 'yyyyMMddHHmmss').zip"
Get-ChildItem -LiteralPath $source -Filter '*.xml' |
ForEach-Object {
  $path = $_.FullName
  $xml = [XmlDocument]::new()
  $xml.PreserveWhitespace = $true
  $xml.Load($path)
  $updated = $false
  $uri = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
  $documentElement = $xml.DocumentElement
  if (
    $null -ne $documentElement -and
    $documentElement.LocalName -eq 'Page' -and
    $documentElement.NamespaceURI -eq $uri -and
    $documentElement.HasAttribute('lang')
  ) {
    $documentElement.RemoveAttribute('lang')
    $updated = $true
  }
  $ns = [XmlNamespaceManager]::new($xml.NameTable)
  $ns.AddNamespace('one', $uri)
  $nodes = $xml.SelectNodes('//one:OE[@lang]', $ns)
  if (-not $updated -and $nodes.Count -gt 0) {
    $updated = $true
  }
  $nodes |
  ForEach-Object { $_.RemoveAttribute('lang') }
  if ($updated) {
    $xml.Save($path)
    "Updated:`t$path" | Out-Host
  }
}
if (@(Get-Item -Path "$source/*").Count -gt 0) {
  Compress-Archive -Path "$source/*" -Destination "$source.$(Get-Date -Format 'yyyyMMddHHmmss').zip"
}

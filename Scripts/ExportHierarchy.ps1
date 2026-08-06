using namespace System.Xml

[CmdletBinding()]
param ()

Import-Module -Name ($PSScriptRoot | Join-Path -ChildPath '..\Modules\Automation.Office.psd1') -Force
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

if (Test-Path -LiteralPath $config.Output) {
  Get-ChildItem -LiteralPath $config.Output -Filter '*.csv' | Remove-Item
} else {
  New-Item -Path $config.Output -ItemType Directory | Out-Null
}
$csv = Resolve-Path -LiteralPath $config.Output | Join-Path -ChildPath 'Hierarchy.csv'
$xml = Get-OneNoteHierarchy -Id $config.StartNodeId -HierarchyScope hsPages
$hierarchies = $xml.SelectNodes('//*') |
Where-Object {
  $_ -is [XmlElement] -and
  $_.HasAttribute('ID') -and
  (-not $_.HasAttribute('isInRecycleBin') -or -not [bool]::Parse($_.GetAttribute('isInRecycleBin')))
}

$hierarchies |
ForEach-Object {
  $nodeId = $_.GetAttribute('ID')
  $nodeName = if ($_.HasAttribute('name')) {
    $_.GetAttribute('name')
  } else {
    [string]::Empty
  }
  $parentId = try {
    if ($_.ParentNode -is [XmlElement] -and $_.ParentNode.HasAttribute('ID')) {
      $_.ParentNode.GetAttribute('ID')
    } else {
      [string]::Empty
    }
  } catch {
    [string]::Empty
  }
  $parentName = try {
    if ($_.ParentNode -is [XmlElement] -and $_.ParentNode.HasAttribute('name')) {
      $_.ParentNode.GetAttribute('name')
    } else {
      [string]::Empty
    }
  } catch {
    [string]::Empty
  }
  [PSCustomObject]@{
    LocalName  = $_.LocalName
    Id         = $nodeId
    Name       = $nodeName
    ParentId   = $parentId
    ParentName = $parentName
  }
} |
Sort-Object -Property LocalName |
Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

$output = $config.Output | Join-Path -ChildPath 'Hierarchy'
if (Test-Path -LiteralPath $output) {
  Get-ChildItem -LiteralPath $output -Filter '*.xml' | Remove-Item
} else {
  New-Item -Path $output -ItemType Directory | Out-Null
}

$hierarchies |
ForEach-Object {
  $destination = Resolve-Path -LiteralPath $output
  Export-OneNoteHierarchy -Id $_.GetAttribute('ID') -Destination $destination
}
if (@(Get-Item -Path "$output/*").Count -gt 0) {
  Compress-Archive -Path "$output/*" -Destination "$output.$(Get-Date -Format 'yyyyMMddHHmmss').zip"
}

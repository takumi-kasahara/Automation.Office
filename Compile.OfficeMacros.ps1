using module .\Modules\Automation.Office.psd1

[CmdletBinding()]
param (
  [switch]
  $Release
)
if ($PSEdition -ne 'Desktop') {
  return
}
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = if ($Release) {
  'Release'
} else {
  'Debug'
}
$bin = '.\Bin'
if (-not (Test-Path -LiteralPath $bin)) {
  New-Item -Path $bin -ItemType Directory | Out-Null
}

try {
  $target = $bin | Join-Path -ChildPath "Personal.$config.xlsb"
  New-ExcelFile -Path $target -FileFormat xlExcel12 -Force
  $source = Resolve-Path -LiteralPath ".\Apps\Personal.xlsb\VBProject.$config.json"
  Import-ExcelVBProject -Path $target -Source $source -Hidden
} catch {
  Get-Process |
  Where-Object -Property Name -EQ 'EXCEL' |
  Stop-Process
  Remove-Item -LiteralPath $target -Force
  Write-Error -ErrorRecord $_
}

try {
  $target = $bin | Join-Path -ChildPath "Normal.$config.dotm"
  New-WordFile -Path $target -FileFormat wdFormatXMLTemplateMacroEnabled -Force
  $source = Resolve-Path -LiteralPath ".\Apps\Normal.dotm\VBProject.$config.json"
  Import-WordVBProject -Path $target -Source $source
} catch {
  Get-Process |
  Where-Object -Property Name -EQ 'WINWORD' |
  Stop-Process
  Remove-Item -LiteralPath $target -Force
  Write-Error -ErrorRecord $_
}

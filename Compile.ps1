[CmdletBinding()]
param (
  [switch]
  $Release
)
if ($PSEdition -ne 'Desktop') {
  return
}
Import-Module -Name ($PSScriptRoot | Join-Path -ChildPath '..\Modules\Automation.Office.psd1') -Force
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$root = Resolve-Path -LiteralPath '.\Bin'
$config = if ($Release) {
  'Release'
} else {
  'Debug'
}

try {
  $bin = $root | Join-Path -ChildPath "Personal.$config.xlsb"
  New-ExcelFile -Path $bin -FileFormat xlExcel12 -Force
  $source = Resolve-Path -LiteralPath ".\Apps\Personal.xlsb\VBProject.$config.json"
  Import-ExcelVBProject -Path $bin -Source $source -Hidden
} catch {
  Get-Process |
  Where-Object -Property Name -EQ 'EXCEL' |
  Stop-Process
  Remove-Item -LiteralPath $bin -Force
  Write-Error -ErrorRecord $_
}

try {
  $bin = $root | Join-Path -ChildPath "Normal.$config.dotm"
  New-WordFile -Path $bin -FileFormat wdFormatXMLTemplateMacroEnabled -Force
  $source = Resolve-Path -LiteralPath ".\Apps\Normal.dotm\VBProject.$config.json"
  Import-WordVBProject -Path $bin -Source $source
} catch {
  Get-Process |
  Where-Object -Property Name -EQ 'WINWORD' |
  Stop-Process
  Remove-Item -LiteralPath $bin -Force
  Write-Error -ErrorRecord $_
}

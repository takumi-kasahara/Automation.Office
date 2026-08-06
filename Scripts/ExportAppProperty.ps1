using module .\Modules\Automation.Office.psd1

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$out = './Out'
try {
  Write-Progress -Activity 'Exporting Application Properties'
  $json = $out | Join-Path -ChildPath 'Excel.Application.json'
  Get-ExcelAppProperty | ConvertTo-Json | Out-File -LiteralPath $json -Encoding UTF8
} finally {
  Write-Progress -Activity 'Done' -Completed
}

[CmdletBinding()]
param ()

Import-Module -Name ($PSScriptRoot | Join-Path -ChildPath '..\Modules\Automation.Office.psd1') -Force
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = './Config'
$json = $config | Join-Path -ChildPath 'Excel.Application.json'
Set-ExcelAppProperty -Properties (Get-Content -LiteralPath $json -Raw | ConvertFrom-Json)

[CmdletBinding()]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Modules\Automation.Office.psm1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = './Config'
$json = $config | Join-Path -ChildPath 'Excel.Application.json'
Set-ExcelAppProperty -Properties (Get-Content -LiteralPath $json -Raw | ConvertFrom-Json)

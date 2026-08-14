<#
.SYNOPSIS
  Imports Excel application properties from a JSON file.

.DESCRIPTION
  This script imports Excel application properties from a JSON file into the Excel Application object.

.NOTES
  The script expects a Config.psd1 file with an Output property containing the Config directory.
#>
using module .\Modules\Automation.Office.psd1

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = './Config'
$json = $config | Join-Path -ChildPath 'Excel.Application.json'
Set-ExcelAppProperty -Properties (Get-Content -LiteralPath $json -Raw | ConvertFrom-Json)

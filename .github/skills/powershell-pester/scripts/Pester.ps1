[CmdletBinding()]
param (
  [Parameter(Mandatory)]
  [ValidateScript({ Test-Path -LiteralPath $_ })]
  [string]
  $Path,
  [int]
  $LineNumber = 0
)

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = New-PesterConfiguration
if ($LineNumber -gt 0) {
  $config.Filter.Line = "$(Resolve-Path -LiteralPath $Path):$LineNumber"
}
$config.Run.Exit = $true
$config.Run.Path = (Resolve-Path -LiteralPath $Path).Path
Invoke-Pester -Configuration $config

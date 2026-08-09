using namespace System.IO

[CmdletBinding()]
param (
  [ValidateScript({ Test-Path -LiteralPath $_ })]
  [string]
  $Path,
  [int]
  $LineNumber = 0,
  [switch]
  $Parallel
)
if ($PSEdition -ne 'Desktop') {
  return
}
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$ext = [Path]::GetExtension($Path)
if (Test-Path -LiteralPath $Path -PathType Container) {
  $module = Get-ChildItem -Path "$([WildcardPattern]::Escape($Path))\*" -File -Include '*.ps1', '*.psm1' | Where-Object { $_.Name -notlike '*.Tests.ps1' }
  $test = Get-ChildItem -LiteralPath $Path -File -Filter '*.Tests.ps1'
} elseif ($ext -eq '.ps1') {
  $parent = [WildcardPattern]::Escape($Path) | Split-Path -Parent
  $base = [Path]::GetFileNameWithoutExtension($Path) -replace '\.Tests$', [string]::Empty
  $module = $parent | Join-Path -ChildPath "$base.ps1"
  $test = $Path
} elseif ($ext -eq '.psm1') {
  $parent = [WildcardPattern]::Escape($Path) | Split-Path -Parent
  $base = [Path]::GetFileNameWithoutExtension($Path)
  $module = $Path
  $test = $parent | Join-Path -ChildPath "$base.Tests.ps1"
} else {
  throw "Unsupported file type: $ext"
}
if (-not (Test-Path -LiteralPath $module)) {
  throw "Module not found: $module"
}
if (-not (Test-Path -LiteralPath $test)) {
  throw "Test not found: $test"
}
$config = New-PesterConfiguration
if ($LineNumber -gt 0 -and @($test).Count -eq 1) {
  $config.Filter.Line = "$((Resolve-Path -LiteralPath $test).Path):$($LineNumber)"
}
$config.Run.Parallel = $Parallel.IsPresent
$config.Run.Path = (Resolve-Path -LiteralPath $test).Path
$config.TestResult.OutputFormat = 'NUnitXml'
if (-not $Parallel.IsPresent) {
  $config.CodeCoverage.OutputFormat = 'JaCoCo'
  $config.CodeCoverage.Path = (Resolve-Path -LiteralPath $module).Path
}
Invoke-Pester -Configuration $config

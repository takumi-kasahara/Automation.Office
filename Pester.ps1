using namespace System.IO

[CmdletBinding()]
param (
  [ValidateScript({ Test-Path -LiteralPath $_ })]
  [string]
  $Path = ($PSScriptRoot | Join-Path -ChildPath 'Modules\Sources')
)
if ($PSEdition -ne 'Desktop') {
  return
}
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$ext = [Path]::GetExtension($Path)
if (Test-Path -LiteralPath $Path -PathType Container) {
  $module = Get-ChildItem -LiteralPath $Path -File -Filter '*.psm1' | Where-Object { $_.BaseName -notlike '.*' }
  $test = Get-ChildItem -LiteralPath $Path -File -Filter '*.Tests.ps1'
}
elseif ($ext -eq '.ps1') {
  $parent = [WildcardPattern]::Escape($Path) | Split-Path -Parent
  $base = [Path]::GetFileNameWithoutExtension($Path) -replace '\.Tests$', [string]::Empty
  $module = $parent | Join-Path -ChildPath "$base.psm1"
  $test = $Path
}
elseif ($ext -eq '.psm1') {
  $parent = [WildcardPattern]::Escape($Path) | Split-Path -Parent
  $base = [Path]::GetFileNameWithoutExtension($Path)
  $module = $Path
  $test = $parent | Join-Path -ChildPath "$base.Tests.ps1"
}
else {
  throw "Unsupported file type: $ext"
}
if (-not (Test-Path -LiteralPath $module)) {
  throw "Module not found: $module"
}
if (-not (Test-Path -LiteralPath $test)) {
  throw "Test not found: $test"
}
$config = New-PesterConfiguration
$config.Run.Path = (Resolve-Path -LiteralPath $test).Path
$config.TestResult.Enabled = $true
$config.TestResult.OutputFormat = 'NUnitXml'
$config.CodeCoverage.Enabled = $true
$config.CodeCoverage.OutputFormat = 'CoverageGutters'
$config.CodeCoverage.Path = (Resolve-Path -LiteralPath $module).Path
Invoke-Pester -Configuration $config

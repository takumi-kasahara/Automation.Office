using namespace System.Management.Automation.Language

[CmdletBinding()]
param (
  [Parameter(ValueFromRemainingArguments)]
  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
  [string[]]
  $LiteralPath
)

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$items = Get-Item -LiteralPath $LiteralPath -Force
$results = $items |
ForEach-Object {
  $errors = $null
  # https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.language.parser.parsefile?view=powershellsdk-7.4.0
  [Parser]::ParseFile($_.FullName, [ref]$null, [ref]$errors) | Out-Null
  if ($errors -and $errors.Count -gt 0) {
    return $errors
  }
}
$results | ConvertTo-Json

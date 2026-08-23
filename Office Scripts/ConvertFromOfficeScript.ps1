using namespace System.IO
using namespace System.Text

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

Get-ChildItem -Filter '*.osts' |
ForEach-Object {
  $destination = [Path]::ChangeExtension($_.FullName, '.ts')
  $content = Get-Content -Path $_.FullName | ConvertFrom-Json
  $normalized = $content.body -replace "`r`n", "`n"
  [File]::WriteAllText($destination, $normalized, [UTF8Encoding]::new($false))
}

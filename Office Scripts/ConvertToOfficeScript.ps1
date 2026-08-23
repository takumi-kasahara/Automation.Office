using namespace System.IO
using namespace System.Text
[CmdletBinding()]
param ()
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

Get-ChildItem -Filter '*.osts' | ForEach-Object {
  $source = [Path]::ChangeExtension($_.FullName, '.ts')
  $json = Get-Content -Path $_.FullName -Raw | ConvertFrom-Json
  $json.body = (Get-Content -Path $source -Raw).ToString()
  $content = $json | ConvertTo-Json -Compress
  [File]::WriteAllText($_.FullName, $content, [UTF8Encoding]::new($false))
}

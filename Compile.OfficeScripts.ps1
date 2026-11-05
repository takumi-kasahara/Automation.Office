using namespace System.IO
using namespace System.Management.Automation
using namespace System.Text

[CmdletBinding()]
param ()
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$root = Resolve-Path -LiteralPath '.\Office Scripts'
Get-ChildItem -LiteralPath $root -Filter '*.ts' |
Where-Object { $_.Name -notlike '*.d.ts' } |
ForEach-Object {
  $source = [Path]::ChangeExtension($_.FullName, '.osts')
  $json = Get-Content -LiteralPath $source -Encoding UTF8 -Raw | ConvertFrom-Json
  $json.body = (Get-Content -LiteralPath $_.FullName -Encoding UTF8 -Raw).ToString()
  $content = $json | ConvertTo-Json -Compress
  [File]::WriteAllText($source, $content, [UTF8Encoding]::new($false))
  Get-Item -LiteralPath $source
}

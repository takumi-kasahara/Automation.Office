using namespace System.IO
using namespace System.Management.Automation
using namespace System.Text

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$root = Resolve-Path -LiteralPath '.\Office Scripts'
Get-ChildItem -LiteralPath $root -Filter '*.ts' | ForEach-Object {
  $source = [Path]::ChangeExtension($_.FullName, '.osts')
  $json = Get-Content -Path $source -Raw | ConvertFrom-Json
  $json.body = (Get-Content -Path $_.FullName -Raw).ToString()
  $content = $json | ConvertTo-Json -Compress
  [File]::WriteAllText($source, $content, [UTF8Encoding]::new($false))
  Get-Item -Path $source
}

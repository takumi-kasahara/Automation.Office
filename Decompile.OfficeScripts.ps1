using namespace System.IO
using namespace System.Management.Automation
using namespace System.Text

[CmdletBinding()]
param ()
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$root = Resolve-Path -LiteralPath '.\Office Scripts'
$target = (Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').Personal | Join-Path -ChildPath 'Office Scripts'
if (-not (Test-Path -LiteralPath $target)) {
  throw [ItemNotFoundException]::new("$target not found.")
}
Get-ChildItem -LiteralPath $target -Filter '*.osts' |
ForEach-Object {
  $destination = $root | Join-Path -ChildPath "$([Path]::GetFileNameWithoutExtension($_.FullName)).ts"
  $content = Get-Content -LiteralPath $_.FullName -Encoding UTF8 -Raw | ConvertFrom-Json
  $normalized = $content.body -replace "`r`n", "`n"
  [File]::WriteAllText($destination, $normalized, [UTF8Encoding]::new($false))
  Get-Item -Path $destination
}

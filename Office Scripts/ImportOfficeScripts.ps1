using namespace System.Management.Automation

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$target = (Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').Personal | Join-Path -ChildPath 'Office Scripts'
if (-not (Test-Path -LiteralPath $target)) {
  throw [ItemNotFoundException]::new("$target not found.")
}
Copy-Item -Path ([WildcardPattern]::Escape($PSScriptRoot) | Join-Path -ChildPath '*.osts') -Destination $target -PassThru

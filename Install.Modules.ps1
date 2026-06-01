[CmdletBinding()]
param ()
if ($PSEdition -ne 'Desktop') {
  return
}
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

Get-ChildItem -LiteralPath 'Modules' -File -Filter '*.psd1' |
ForEach-Object {
  $destination = [WildcardPattern]::Escape($PROFILE.CurrentUserAllHosts) |
  Split-Path -Parent |
  Join-Path -ChildPath 'Modules' |
  Join-Path -ChildPath $_.BaseName

  "Create:`t$destination" | Out-Host
  if (-not (Test-Path -LiteralPath $destination)) {
    New-Item -Path $destination -ItemType Directory >$null
  }
  Copy-Item -Path 'Modules\*' -Destination $destination -Exclude '*.Tests.ps1' -Recurse -Force -PassThru
}

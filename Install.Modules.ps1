[CmdletBinding()]
param ()
if ($PSEdition -ne 'Desktop') {
  return
}
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

# Word > Options > Save > Create new files in the cloud automatically: false
# https://support.microsoft.com/en-us/office/collab-files/what-administrators-need-to-know-about-the-cloud-focused-save-experience-in-office
@(
  'HKCU:\Software\Microsoft\Office\16.0\Common\General'
) |
Where-Object { Test-Path -LiteralPath $_ } |
ForEach-Object { Set-ItemProperty -LiteralPath $_ -Name PreferCloudSaveLocations -Value 0 -Type DWord }

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

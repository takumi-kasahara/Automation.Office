[CmdletBinding()]
param ()
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot
if ($PSEdition -ne 'Desktop') {
  return
}

$bin = '.\Bin'
if (-not (Test-Path -LiteralPath $bin)) {
  New-Item -Path $bin -ItemType Directory | Out-Null
}

Get-ChildItem -Path '.\Modules\*.psd1' -File |
ForEach-Object {
  try {
    Test-ModuleManifest -Path $_ -ErrorAction Stop
  } catch {
    Write-Warning -Message $_.Exception.Message
    exit 1
  }
  Compress-PSResource -Path '.\Modules' -DestinationPath '.\Bin' -PassThru
}

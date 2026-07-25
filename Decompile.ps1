[CmdletBinding()]
param (
  [switch]
  $Release
)
if ($PSEdition -ne 'Desktop') {
  return
}
$modulePath = $PSScriptRoot | Join-Path -ChildPath '.\Modules\Automation.Office.psm1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$config = if ($Release) {
  'Release'
} else {
  'Debug'
}

try {
  $source = if ($Release) {
    $env:APPDATA | Join-Path -ChildPath 'Microsoft\Excel\XLSTART\Personal.xlsb'
  } else {
    Resolve-Path -LiteralPath '.\Bin\Personal.Debug.xlsb'
  }
  if ((Test-Path -LiteralPath $source)) {
    $destination = Resolve-Path -LiteralPath '.\Apps\Personal.xlsb'
    $componentRoot = $destination | Join-Path -ChildPath 'VBProject'
    Export-ExcelVBProject -Path $source -Destination ($destination | Join-Path -ChildPath "VBProject.$config.json") -ComponentRoot $componentRoot
  } else {
    Write-Warning -Message "$source not found."
  }
} catch {
  Get-Process |
  Where-Object -Property Name -EQ 'EXCEL' |
  Stop-Process
  Write-Error -ErrorRecord $_
}

try {
  $source = if ($Release) {
    $env:APPDATA | Join-Path -ChildPath 'Microsoft\Templates\Normal.dotm'
  } else {
    Resolve-Path -LiteralPath '.\Bin\Normal.Debug.dotm'
  }
  if ((Test-Path -LiteralPath $source)) {
    $destination = Resolve-Path -LiteralPath '.\Apps\Normal.dotm'
    $componentRoot = $destination | Join-Path -ChildPath 'VBProject'
    Export-WordVBProject -Path $source -Destination ($destination | Join-Path -ChildPath "VBProject.$config.json") -ComponentRoot $componentRoot
  } else {
    Write-Warning -Message "$source not found."
  }
} catch {
  Get-Process |
  Where-Object -Property Name -EQ 'WINWORD' |
  Stop-Process
  Write-Error -ErrorRecord $_
}

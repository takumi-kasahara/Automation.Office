[CmdletBinding()]
param ()
if ($PSEdition -ne 'Desktop') {
  return
}
Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

# Trust Center > Macro Settings > Trust access to the VBA project object model
@(
  'HKCU:\Software\Microsoft\Office\16.0\Access\Security'
  'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
  'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security'
  'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
) |
Where-Object { Test-Path -LiteralPath $_ } |
ForEach-Object { Set-ItemProperty -LiteralPath $_ -Name AccessVBOM -Value 1 -Type DWord }

# Tools > Options > General > Error Trapping > Break in Class Module
$vba = 'HKCU:\Software\Microsoft\VBA\7.1\Common'
if (Test-Path -LiteralPath $vba) {
  Set-ItemProperty -LiteralPath $vba -Name BreakOnAllErrors -Value 0 -Type DWord
  Set-ItemProperty -LiteralPath $vba -Name BreakOnServerErrors -Value 1 -Type DWord
}

@(
  @{
    Source      = Resolve-Path -LiteralPath 'Bin' | Join-Path -ChildPath 'Personal.Release.xlsb'
    Destination = $env:APPDATA | Join-Path -ChildPath 'Microsoft\Excel\XLSTART\Personal.xlsb'
  }
  @{
    Source      = Resolve-Path -LiteralPath 'Bin' | Join-Path -ChildPath 'Normal.Release.dotm'
    Destination = $env:APPDATA | Join-Path -ChildPath 'Microsoft\Templates\Normal.dotm'
  }
) |
ForEach-Object {
  if (Test-Path -LiteralPath $_.Source) {
    if (-not (Test-Path -LiteralPath $_.Destination)) {
      New-Item -ItemType Directory -Path $_.Destination >$null
    }
    Copy-Item -LiteralPath $_.Source -Destination $_.Destination -PassThru
  }
  else {
    Write-Warning -Message "$($_.Source) not found."
  }
}

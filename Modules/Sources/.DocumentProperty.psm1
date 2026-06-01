using namespace Microsoft.Office.Core
using namespace System.Diagnostics.CodeAnalysis
using namespace System.Reflection
using namespace System.Runtime.InteropServices

Set-StrictMode -Version Latest

function Get-MsoDocProperty {
  [CmdletBinding()]
  # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.msodocproperties?view=office-pia
  [OutputType([Microsoft.Office.Core.MsoDocProperties])]
  param (
    [Parameter(Mandatory)]
    [ValidateNotNull()]
    [type]
    $Type
  )
  if ($Type -in [byte], [sbyte], [int16], [uint16], [int], [uint32], [long], [uint64]) {
    return [MsoDocProperties]::msoPropertyTypeNumber
  }
  if ($Type -in [bool]) {
    return [MsoDocProperties]::msoPropertyTypeBoolean
  }
  if ($Type -in [datetime]) {
    return [MsoDocProperties]::msoPropertyTypeDate
  }
  if ($Type -in [double], [float]) {
    return [MsoDocProperties]::msoPropertyTypeFloat
  }
  if ($Type -in [string]) {
    return [MsoDocProperties]::msoPropertyTypeString
  }
  throw [ArgumentException]::new("Unsupported:$Type")
}
function Get-DocumentProperty {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [string[]]
    $Name,
    [switch]
    $Custom
  )
  $values = [PSCustomObject]@{}
  $properties = if ($Custom) {
    $InputObject.CustomDocumentProperties
  }
  else {
    $InputObject.BuiltInDocumentProperties
  }
  $properties |
  ForEach-Object {
    try {
      $key = [__ComObject].InvokeMember('Name', [BindingFlags]::GetProperty, $null, $_, $null)
      if ($null -ne $Name -and $Name.Count -gt 0 -and $key -notin $Name) {
        return
      }
      $value = [__ComObject].InvokeMember('Value', [BindingFlags]::GetProperty, $null, $_, $null)
      $values | Add-Member -MemberType NoteProperty -Name $key -Value $value
    }
    catch [COMException] {
      Write-Warning -Message "$($_.Exception.Message) on $key"
    }
  }
  return $values
}
function Set-DocumentProperty {
  [CmdletBinding()]
  [OutputType([bool])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function updates document properties, which is the intended purpose of the utility')]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [Parameter(Mandatory)]
    [PSObject]
    $Properties,
    [switch]
    $Custom
  )
  try {
    if ($Custom) {
      return Set-CustomDocumentProperty -InputObject $InputObject -Properties $Properties
    }
    else {
      return Set-BuiltInDocumentProperty -InputObject $InputObject -Properties $Properties
    }
    return $true
  }
  finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Set-BuiltInDocumentProperty {
  [CmdletBinding()]
  [OutputType([bool])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function updates built-in document properties, which is the intended purpose of the utility')]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [Parameter(Mandatory)]
    [PSObject]
    $Properties
  )
  try {
    $result = $Properties.PSObject.Properties |
    ForEach-Object {
      try {
        # https://learn.microsoft.com/en-us/office/vba/api/office.documentproperties.item
        $property = $InputObject.BuiltInDocumentProperties($_.Name)
        # https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.value
        [__ComObject].InvokeMember('Value', [BindingFlags]::SetProperty, $null, $property, $_.Value)
        return $true
      }
      catch [COMException] {
        Write-Warning -Message "$($_.Exception.Message) on $($_.Name)"
        return $false
      }
    }
    return $result -contains $true
  }
  finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Set-CustomDocumentProperty {
  [CmdletBinding()]
  [OutputType([bool])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function updates custom document properties, which is the intended purpose of the utility')]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [Parameter(Mandatory)]
    [PSObject]
    $Properties
  )
  try {
    $result = $Properties.PSObject.Properties |
    ForEach-Object {
      try {
        $property = try {
          # https://learn.microsoft.com/en-us/office/vba/api/office.documentproperties.item
          $InputObject.CustomDocumentProperties($_.Name)
        }
        catch [ArgumentException] {
          $null
        }
        if ($property) {
          # https://learn.microsoft.com/en-us/office/vba/api/office.documentproperty.delete
          [__ComObject].InvokeMember('Delete', [BindingFlags]::InvokeMethod, $null, $property, $null)
        }
        # https://learn.microsoft.com/en-us/office/vba/api/office.documentproperties.add
        $arguments = @($_.Name, $false, (Get-MsoDocProperty -Type ($_.Value.GetType())), $_.Value)
        [__ComObject].InvokeMember('Add', [BindingFlags]::InvokeMethod, $null, $InputObject.CustomDocumentProperties, $arguments) | Out-Null
        return $true
      }
      catch [COMException] {
        Write-Warning -Message "$($_.Exception.Message) on $($_.Name)"
        return $false
      }
    }
    return $result -contains $true
  }
  finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

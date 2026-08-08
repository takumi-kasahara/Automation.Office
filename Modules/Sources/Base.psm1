using module .\DialogSuppressor.psm1
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Reflection

Set-StrictMode -Version Latest

function Get-ObjectProperty {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject
  )
  try {
    $properties = [PSCustomObject]@{}
    $InputObject |
    Get-Member -MemberType Properties |
    Where-Object {
      $propertyInfo = $InputObject.GetType().GetProperty($_.Name)
      if ($propertyInfo) {
        if (-not $propertyInfo.CanRead) {
          return $false
        }
        $type = $propertyInfo.PropertyType
      } else {
        $property = $InputObject.PSObject.Properties | Where-Object -Property Name -Value $_.Name -EQ | Select-Object -First 1
        if (-not $property.IsGettable) {
          return $false
        }
        $type = [type]$property.TypeNameOfValue
      }
      return $type.IsPrimitive -or $type.IsEnum -or $type -eq [datetime]
    } |
    ForEach-Object {
      $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $InputObject.$($_.Name)
    }
    return $properties
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Set-ObjectProperty {
  [CmdletBinding()]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function sets properties on a COM object, which is the intended purpose of the utility')]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [Parameter(Mandatory)]
    [PSCustomObject]
    $Properties
  )
  try {
    $Properties.PSObject.Properties |
    ForEach-Object {
      $propertyInfo = $InputObject.GetType().GetProperty($_.Name)
      if ($propertyInfo) {
        if (-not $propertyInfo.CanWrite) {
          Write-Warning -Message "$($_.Name) is read-only."
          return
        }
        $type = $propertyInfo.PropertyType
      } else {
        $property = $InputObject.PSObject.Properties | Where-Object -Property Name -Value $_.Name -EQ | Select-Object -First 1
        if (-not $property) {
          throw [InvalidOperationException] "Property $($_.Name) does not exist."
          return
        }
        if (-not $property.IsSettable) {
          Write-Warning -Message "$($_.Name) is read-only."
          return
        }
        $type = [type]$property.TypeNameOfValue
      }
      if ($type.IsPrimitive -or $type.IsEnum -or $type -eq [datetime]) {
        if ($InputObject.$($_.Name) -eq $_.Value) {
          return [PSCustomObject]@{
            Name   = $_.Name
            Value  = $_.Value
            Status = 'Skipped'
          }
        }
        try {
          [__ComObject].InvokeMember($_.Name, [BindingFlags]::SetProperty, $null, $InputObject, $_.Value)
          return [PSCustomObject]@{
            Name   = $_.Name
            Value  = $_.Value
            Status = 'Success'
          }
        } catch {
          Write-Warning -Message $_.Exception.Message
          return [PSCustomObject]@{
            Name   = $_.Name
            Value  = $_.Value
            Status = 'Failed'
          }
        }
      }
    } |
    Out-Host
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

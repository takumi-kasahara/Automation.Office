using module .\DialogSuppressor.psm1
using namespace System.Collections.Generic
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Reflection
using namespace System.Runtime.CompilerServices

Set-StrictMode -Version Latest

function Get-ObjectProperty {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [switch]
    $Recurse,
    [ValidateRange(0, [int]::MaxValue)]
    [int]
    $Depth = 1
  )
  try {
    $visited = [HashSet[int]]::new()
    function Get-RecursiveObjectProperty {
      [CmdletBinding()]
      [OutputType([PSCustomObject])]
      param (
        [Parameter(Mandatory)]
        [__ComObject]
        $TargetObject,
        [Parameter(Mandatory)]
        [int]
        $RemainingDepth,
        [switch]
        $Recursive
      )
      if ($null -eq $TargetObject) {
        return [PSCustomObject]@{}
      }
      $identity = [RuntimeHelpers]::GetHashCode($TargetObject)
      if ($visited.Contains($identity)) {
        return [PSCustomObject]@{}
      }
      [void]$visited.Add($identity)

      $properties = [PSCustomObject]@{}
      $memberProperties = @(
        Get-Member -InputObject $TargetObject -MemberType Properties -ErrorAction SilentlyContinue
      )
      $memberProperties |
      ForEach-Object {
        $propertyInfo = $TargetObject.GetType().GetProperty($_.Name)
        $property = $null
        if ($propertyInfo) {
          if (-not $propertyInfo.CanRead) {
            return
          }
          $type = $propertyInfo.PropertyType
        } else {
          $property = $TargetObject.PSObject.Properties | Where-Object -Property Name -Value $_.Name -EQ | Select-Object -First 1
          if (-not $property) {
            return
          }
          if (-not $property.IsGettable) {
            return
          }
          $type = [type]$property.TypeNameOfValue
        }

        if ($type.IsPrimitive -or $type.IsEnum -or $type -eq [datetime]) {
          try {
            $value = $TargetObject.$($_.Name)
          } catch {
            return
          }
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $value
          return
        }
        if (-not $Recursive -or $RemainingDepth -le 0) {
          return
        }
        try {
          $value = $TargetObject.$($_.Name)
        } catch {
          return
        }
        if ($null -eq $value -or $value -isnot [__ComObject]) {
          return
        }
        $child = Get-RecursiveObjectProperty -TargetObject $value -RemainingDepth ($RemainingDepth - 1) -Recursive:$Recursive
        if (@($child.PSObject.Properties).Count -gt 0) {
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $child
        }
      }
      return $properties
    }
    return Get-RecursiveObjectProperty -TargetObject $InputObject -RemainingDepth $Depth -Recursive:$Recurse
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

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
  [SuppressMessage('PSReviewUnusedParameter', 'Name', Justification = 'Parameter is used by the nested recursive helper, which PSScriptAnalyzer does not track')]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $InputObject,
    [string[]]
    $Name,
    [ValidateRange(1, [int]::MaxValue)]
    [int]
    $Depth = 1
  )
  try {
    $activity = 'Processing Object Properties'
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
        [string]
        $ParentName = [string]::Empty
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
      if ($Name -and $RemainingDepth -eq $Depth) {
        $memberProperties = $memberProperties |
        Where-Object -Property Name -In -Value $Name
      }
      $memberProperties |
      ForEach-Object {
        $displayName = if ($ParentName) {
          "$ParentName.$($_.Name)"
        } else {
          $_.Name
        }
        # Some properties can initialize the mail client and block inside the COM getter.
        if ($_.Name -in 'MailEnvelope') {
          Write-Warning -Message "$($displayName): Skipped because accessing it can block automation."
          return
        }
        Write-Progress -Activity $activity -Status "Processing: $displayName"
        $propertyInfo = try {
          $TargetObject.GetType().GetProperty($_.Name)
        } catch {
          Write-Warning -Message "$($displayName): $($_.Exception.Message)"
          $null
        }
        if ($propertyInfo) {
          if (-not $propertyInfo.CanRead) {
            Write-Warning -Message "$displayName is write-only."
            return
          }
          $type = $propertyInfo.PropertyType
        } else {
          $property = $TargetObject.PSObject.Properties | Where-Object -Property Name -Value $_.Name -EQ | Select-Object -First 1
          if (-not $property.IsGettable) {
            Write-Warning -Message "$displayName is write-only."
            return
          }
          $type = [type]$property.TypeNameOfValue
        }
        $value = $TargetObject.$($_.Name)
        if ($null -ne $value) {
          try {
            $type = $value.GetType()
          } catch {
            Write-Warning -Message "$($displayName): $($_.Exception.Message)"
          }
        }
        if ($type.IsPrimitive -or $type.IsEnum -or $type -eq [datetime] -or $type -eq [string]) {
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $value
          return
        }
        if ($RemainingDepth -eq 0) {
          return
        }
        if ($null -eq $value) {
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $null
          return
        }
        if ($value -isnot [__ComObject]) {
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $value
          return
        }
        $child = Get-RecursiveObjectProperty -TargetObject $value -RemainingDepth ($RemainingDepth - 1) -ParentName $displayName
        if (@($child.PSObject.Properties).Count -gt 0) {
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $child
        } else {
          $properties | Add-Member -MemberType NoteProperty -Name $_.Name -Value $null
        }
      }
      if ($null -ne $Name -and @($Name).Count -eq 1) {
        return $properties.$Name
      }
      return $properties
    }
    return Get-RecursiveObjectProperty -TargetObject $InputObject -RemainingDepth $Depth
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Progress -Activity $activity -Completed
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
          return
        }
        try {
          [__ComObject].InvokeMember($_.Name, [BindingFlags]::SetProperty, $null, $InputObject, $_.Value)
        } catch {
          Write-Warning -Message "$($_.Name)`t$($_.Exception.Message)"
        }
      }
    } | Out-Host
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

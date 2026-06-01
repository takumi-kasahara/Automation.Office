using namespace Microsoft.Office.Interop.Excel

Set-StrictMode -Version Latest

function Get-ExcelDocumentProperty {
  <#
  .SYNOPSIS
    Gets document properties from workbooks.

  .DESCRIPTION
    Opens one or more workbooks and returns document properties as a `PSCustomObject`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected workbooks via `-PasswordToOpen` and `-PasswordToModify`.

    When `-Custom` is specified, the cmdlet reads custom document properties instead of built-in properties.

  .PARAMETER Path
    Specifies workbook paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies workbook paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Filters returned properties by name. If not specified, all properties are returned.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

  .PARAMETER Custom
    Gets custom document properties.

  .EXAMPLE
    Get-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title

    Gets the built-in `Title` property from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-ExcelDocumentProperty -LiteralPath "$env:TEMP\Workbook.xlsx" -PasswordToOpen $password -Name Title

    Gets properties from a file protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -PasswordToModify $password -Name Title

    Gets properties from a file protected with a modify password.

  .EXAMPLE
    Get-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Custom -Name MyCustomProperty

    Gets a custom document property.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Each NoteProperty corresponds to a file property.

  .NOTES
    If a file cannot be opened (for example, incorrect password), the cmdlet throws an error record and continues processing remaining items.
  #>
  [CmdletBinding(DefaultParameterSetName = 'PathSet')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [Parameter(Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Name,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $Custom
  )
  begin {
    $app = New-ExcelObject
  }
  process {
    try {
      $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
        'PathSet' {
          Get-Item -Path $Path -Force
        }
        'LiteralPathSet' {
          Get-Item -LiteralPath $LiteralPath -Force
        }
      }
      $items |
      ForEach-Object {
        $file = Open-ExcelFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly
        try {
          return Get-DocumentProperty -InputObject $file -Name $Name -Custom:$Custom
        }
        finally {
          $file.Close()
        }
      }
    }
    catch {
      try {
        if ($app) {
          $app.Quit()
        }
      }
      finally {
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
  end {
    try {
      if ($app) {
        $app.Quit()
      }
    }
    finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
}
function Get-ExcelPropertyValue {
  <#
  .SYNOPSIS
    Gets a single document property value from workbooks.

  .DESCRIPTION
    Wrapper for `Get-ExcelDocumentProperty` that returns only the value of `-Name`.

  .PARAMETER Path
    Specifies workbook paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies workbook paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the document property name to read.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

  .PARAMETER Custom
    Gets custom document properties.

  .EXAMPLE
    Get-ExcelPropertyValue -Path "$env:TEMP\Workbook.xlsx" -Name Title

  .OUTPUTS
    System.Object
      The value of the specified property.

  .NOTES
    This cmdlet calls `Get-ExcelDocumentProperty` internally.
  #>
  [CmdletBinding(DefaultParameterSetName = 'PathSet')]
  [OutputType([object])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [Parameter(Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Name,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $Custom
  )
  process {
    Get-ExcelDocumentProperty @PSBoundParameters |
    ForEach-Object {
      if ($null -eq $_) {
        return
      }
      $_.PSObject.Properties.Value
    }
  }
}
function Set-ExcelDocumentProperty {
  <#
  .SYNOPSIS
    Sets built-in or custom document properties on workbooks.

  .DESCRIPTION
    Opens one or more workbooks and updates a built-in or custom document property.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
    You can use `-WhatIf` and `-Confirm` to preview or confirm updates.

    By default, read-only items and read-only recommended workbooks throw an exception.
    Use `-Force` to ignore the read-only recommended flag when opening a file.

  .PARAMETER Path
    Specifies workbook paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies workbook paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the document property name to update. When `-Custom` is specified, the cmdlet updates a custom document property.

  .PARAMETER Value
    Specifies the value assigned to the property.

  .PARAMETER InputObject
    Specifies a `PSObject` whose properties are mapped to document property names and values.
    Use this parameter to update multiple properties at once.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

  .PARAMETER Custom
    Updates a custom document property instead of a built-in property.

  .PARAMETER Force
    Forces update processing when a file is marked as read-only recommended.

  .PARAMETER PassThru
    Returns the updated document property object when specified.

  .EXAMPLE
    Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title -Value 'Monthly report'

    Sets the `Title` built-in document property.

  .EXAMPLE
    Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Custom -Name MyProperty -Value 'Custom value'

    Sets the `MyProperty` custom document property.

  .EXAMPLE
    Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -InputObject ([PSCustomObject]@{ Title = 'Monthly report'; Subject = 'Sales' })

    Updates multiple built-in document properties in a single call.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title -Value 'Protected book' -PasswordToOpen $password -Force

    Updates a protected workbook.

  .EXAMPLE
    Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title -Value 'Monthly report' -PassThru

    Updates the `Title` built-in document property and returns the updated property object.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      When `-PassThru` is specified.

  .NOTES
    If a file cannot be updated (for example, read-only state, invalid property name, or file access issues), the cmdlet throws an exception and continues processing remaining items.
  #>
  [CmdletBinding(DefaultParameterSetName = 'ValuePathSet', SupportsShouldProcess)]
  [OutputType([PSCustomObject], [void])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'ValuePathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'PSObjectPathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'ValueLiteralPathSet', ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'PSObjectLiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [Parameter(Mandatory, ParameterSetName = 'ValuePathSet', Position = 1, ValueFromPipeline)]
    [Parameter(Mandatory, ParameterSetName = 'ValueLiteralPathSet', Position = 1, ValueFromPipeline)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Name,
    [Parameter(Mandatory, ParameterSetName = 'ValuePathSet', Position = 2)]
    [Parameter(Mandatory, ParameterSetName = 'ValueLiteralPathSet', Position = 2)]
    [object]
    $Value,
    [Parameter(Mandatory, ParameterSetName = 'PSObjectPathSet', Position = 1, ValueFromPipeline)]
    [Parameter(Mandatory, ParameterSetName = 'PSObjectLiteralPathSet', Position = 1, ValueFromPipeline)]
    [PSObject]
    $InputObject,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $Custom,
    [switch]
    $PassThru,
    [switch]
    $Force
  )
  begin {
    $app = New-ExcelObject
  }
  process {
    try {
      $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
        { $_ -in 'ValuePathSet', 'PSObjectPathSet' } {
          Get-Item -Path $Path -Force
        }
        { $_ -in 'ValueLiteralPathSet', 'PSObjectLiteralPathSet' } {
          Get-Item -LiteralPath $LiteralPath -Force
        }
      }
      $items |
      ForEach-Object {
        $fullName = $_.FullName
        $target = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
          { $_ -in 'ValuePathSet', 'ValueLiteralPathSet' } {
            "Item: $fullName Property: $Name Value: $Value"
          }
          { $_ -in 'PSObjectPathSet', 'PSObjectLiteralPathSet' } {
            "Item: $fullName Properties: $($InputObject.PSObject.Properties.Name -join ', ')"
          }
        }
        if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($target, 'Set DocumentProperty'))) {
          return
        }
        if ($_.IsReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $_))
        }
        $file = Open-ExcelFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Force:$Force
        try {
          if ($file.ReadOnly) {
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $_))
          }
          switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
            { $_ -in 'ValuePathSet', 'ValueLiteralPathSet' } {
              $file.Saved = -not (Set-DocumentProperty -InputObject $file -Properties ([PSCustomObject]@{ $Name = $Value }) -Custom:$Custom)
            }
            { $_ -in 'PSObjectPathSet', 'PSObjectLiteralPathSet' } {
              $file.Saved = -not (Set-DocumentProperty -InputObject $file -Properties $InputObject -Custom:$Custom)
            }
          }
          if (-not $file.Saved) {
            $file.Save()
          }
          if ($PassThru) {
            switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
              { $_ -in 'ValuePathSet', 'ValueLiteralPathSet' } {
                return Get-DocumentProperty -InputObject $file -Name $Name -Custom:$Custom
              }
              { $_ -in 'PSObjectPathSet', 'PSObjectLiteralPathSet' } {
                return Get-DocumentProperty -InputObject $file -Name $InputObject.PSObject.Properties.Name -Custom:$Custom
              }
            }
          }
        }
        finally {
          $file.Close()
        }
      }
    }
    catch {
      try {
        if ($app) {
          $app.Quit()
        }
      }
      finally {
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
  end {
    try {
      if ($app) {
        $app.Quit()
      }
    }
    finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
}
function Remove-ExcelDocumentProperty {
  <#
  .SYNOPSIS
    Removes document properties from workbooks.

  .DESCRIPTION
    Opens one or more workbooks and removes document information by using `Workbook.RemoveDocumentInformation`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    You can use `-WhatIf` and `-Confirm` to preview or confirm removals.

    By default, read-only items and read-only recommended workbooks throw an exception.
    Use `-Force` to ignore the read-only recommended flag when opening a file.

  .PARAMETER Path
    Specifies workbook paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies workbook paths literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

  .PARAMETER RemoveDocInfoType
    Specifies the document information type removed by Excel.

    The default value is `xlRDIDocumentProperties`.

  .PARAMETER Force
    Forces removal processing when a file is marked as read-only recommended.

  .EXAMPLE
    Remove-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx"

    Removes document properties from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Remove-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -PasswordToOpen $password -Force

    Removes document properties from a protected workbook.

  .EXAMPLE
    Remove-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -RemoveDocInfoType xlRDIAll

    Removes all supported document information categories.

  .OUTPUTS
    None.

  .NOTES
    If a file cannot be updated (for example, read-only state or file access issues), the cmdlet throws an exception and continues processing remaining items.
  #>
  [CmdletBinding(DefaultParameterSetName = 'PathSet', SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlremovedocinfotype?view=excel-pia
    [Microsoft.Office.Interop.Excel.XlRemoveDocInfoType]
    $RemoveDocInfoType = [Microsoft.Office.Interop.Excel.XlRemoveDocInfoType]::xlRDIDocumentProperties,
    [switch]
    $Force
  )
  begin {
    $app = New-ExcelObject
  }
  process {
    try {
      $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
        'PathSet' {
          Get-Item -Path $Path -Force
        }
        'LiteralPathSet' {
          Get-Item -LiteralPath $LiteralPath -Force
        }
      }
      $items |
      ForEach-Object {
        $target = "Item: $($_.FullName) RemoveDocInfoType: $RemoveDocInfoType"
        if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($target, 'Remove DocumentProperty'))) {
          return
        }
        if ($_.IsReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $_))
        }
        $file = Open-ExcelFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Force:$Force
        try {
          if ($file.ReadOnly) {
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $_))
          }
          $file.RemoveDocumentInformation($RemoveDocInfoType)
          $file.Saved = $false
          $file.Save()
        }
        finally {
          $file.Close()
        }
      }
    }
    catch {
      try {
        if ($app) {
          $app.Quit()
        }
      }
      finally {
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
  end {
    try {
      if ($app) {
        $app.Quit()
      }
    }
    finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
}

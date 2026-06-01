using namespace Microsoft.Office.Interop.PowerPoint

Set-StrictMode -Version Latest

function Get-PowerPointDocumentProperty {
  <#
  .SYNOPSIS
    Gets document properties from presentations.

  .DESCRIPTION
    Opens one or more presentations and returns document properties as a `PSCustomObject`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected presentations via `-PasswordToOpen` and `-PasswordToModify`.

    When `-Custom` is specified, the cmdlet reads custom document properties instead of built-in properties.

  .PARAMETER Path
    Specifies presentation paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies presentation paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Filters returned properties by name. If not specified, all properties are returned.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER Custom
    Gets custom document properties.

  .EXAMPLE
    Get-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Name Title

    Gets the built-in `Title` property from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-PowerPointDocumentProperty -LiteralPath "$env:TEMP\Presentation.pptx" -PasswordToOpen $password -Name Title

    Gets properties from a file protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -PasswordToModify $password -Name Title

    Gets properties from a file protected with a modify password.

  .EXAMPLE
    Get-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Custom -Name MyCustomProperty

    Gets a custom document property.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Each NoteProperty corresponds to a file property.

  .NOTES
    If a file cannot be opened (for example, file access issues), the cmdlet writes an error record and continues processing remaining items.
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
    $app = New-PowerPointObject
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
        $file = Open-PowerPointFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly
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
function Get-PowerPointPropertyValue {
  <#
  .SYNOPSIS
    Gets a single document property value from presentations.

  .DESCRIPTION
    Wrapper for `Get-PowerPointDocumentProperty` that returns only the value of `-Name`.

  .PARAMETER Path
    Specifies presentation paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies presentation paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the document property name to read.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER Custom
    Gets custom document properties.

  .EXAMPLE
    Get-PowerPointPropertyValue -Path "$env:TEMP\Presentation.pptx" -Name Title

  .OUTPUTS
    System.Object
      The value of the specified property.

  .NOTES
    This cmdlet calls `Get-PowerPointDocumentProperty` internally.
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
    Get-PowerPointDocumentProperty @PSBoundParameters |
    ForEach-Object {
      if ($null -eq $_) {
        return
      }
      $_.PSObject.Properties.Value
    }
  }
}
function Set-PowerPointDocumentProperty {
  <#
  .SYNOPSIS
    Sets built-in or custom document properties on presentations.

  .DESCRIPTION
    Opens one or more presentations and updates a built-in or custom document property.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
    You can use `-WhatIf` and `-Confirm` to preview or confirm updates.

    If a target is read-only, the cmdlet throws an exception.

  .PARAMETER Path
    Specifies presentation paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies presentation paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the document property name to update. When `-Custom` is specified, the cmdlet updates a custom document property.

  .PARAMETER Value
    Specifies the value assigned to the property.

  .PARAMETER InputObject
    Specifies a `PSObject` whose properties are mapped to document property names and values.
    Use this parameter to update multiple properties at once.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER Custom
    Updates a custom document property instead of a built-in property.

  .PARAMETER PassThru
    Returns the updated document property object when specified.

  .EXAMPLE
    Set-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Name Title -Value 'Monthly report'

    Sets the `Title` built-in document property.

  .EXAMPLE
    Set-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Custom -Name MyProperty -Value 'Custom value'

    Sets the `MyProperty` custom document property.

  .EXAMPLE
    Set-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -InputObject ([PSCustomObject]@{ Title = 'Monthly report'; Subject = 'Sales' })

    Updates multiple built-in document properties in a single call.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Set-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Name Title -Value 'Protected deck' -PasswordToOpen $password

    Updates a protected presentation.

  .EXAMPLE
    Set-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Name Title -Value 'Monthly report' -PassThru

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
    $PassThru
  )
  begin {
    $app = New-PowerPointObject
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
        if (-not $PSCmdlet.ShouldProcess($target, 'Set DocumentProperty')) {
          return
        }
        if ($_.IsReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $_))
        }
        $file = Open-PowerPointFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify
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
function Remove-PowerPointDocumentProperty {
  <#
  .SYNOPSIS
    Removes document properties from presentations.

  .DESCRIPTION
    Opens one or more presentations and removes document information by using `Presentation.RemoveDocumentInformation`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    You can use `-WhatIf` and `-Confirm` to preview or confirm removals.

    Read-only files and read-only presentations throw an exception.

  .PARAMETER Path
    Specifies presentation paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies presentation paths literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER RemoveDocInfoType
    Specifies the document information type removed by PowerPoint.

    The default value is `ppRDIDocumentProperties`.

  .EXAMPLE
    Remove-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx"

    Removes document properties from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Remove-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -PasswordToOpen $password

    Removes document properties from a protected presentation.

  .EXAMPLE
    Remove-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -RemoveDocInfoType ppRDIAll

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
    # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
    [Microsoft.Office.Interop.PowerPoint.PpRemoveDocInfoType]
    $RemoveDocInfoType = [Microsoft.Office.Interop.PowerPoint.PpRemoveDocInfoType]::ppRDIDocumentProperties
  )
  begin {
    $app = New-PowerPointObject
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
        if (-not $PSCmdlet.ShouldProcess($target, 'Remove DocumentProperty')) {
          return
        }
        if ($_.IsReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $_))
        }
        $file = Open-PowerPointFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify
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

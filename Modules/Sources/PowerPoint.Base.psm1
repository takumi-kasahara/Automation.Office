using namespace Microsoft.Office.Core
using namespace Microsoft.Office.Interop.PowerPoint
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Net
using namespace System.Runtime.InteropServices

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Set-StrictMode -Version Latest

#region Private
function New-PowerPointObject {
  [CmdletBinding()]
  [OutputType([__ComObject])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for PowerPoint application and does not change persistent state')]
  param (
    [switch]
    $NoSetup
  )
  # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application
  while ($true) {
    try {
      $app = New-Object -ComObject PowerPoint.Application
      if ($app -and $app.hWnd) {
        break
      }
    }
    catch [COMException] {
      Write-Warning -Message $_.Exception.Message
    }
    Start-Sleep -Milliseconds 100
  }
  try {
    if (-not $NoSetup) {
      # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppalertlevel
      $app.DisplayAlerts = [PpAlertLevel]::ppAlertsNone
    }
    return $app
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
  finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Open-PowerPointFile {
  [CmdletBinding()]
  [OutputType([__ComObject])]
  param(
    [Parameter(Mandatory)]
    [__ComObject]
    $Application,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $ReadOnly
  )
  $resolved = (Resolve-Path -LiteralPath $Path).Path
  $passwordToOpenString = if ($null -eq $PasswordToOpen) {
    [type]::Missing
  }
  else {
    [NetworkCredential]::new([string]::Empty, $PasswordToOpen).Password
  }
  $passwordToModifyString = if ($null -eq $PasswordToModify) {
    [type]::Missing
  }
  else {
    [NetworkCredential]::new([string]::Empty, $PasswordToModify).Password
  }
  $readOnlyValue = if ($ReadOnly) {
    [MsoTriState]::msoTrue
  }
  else {
    [MsoTriState]::msoFalse
  }
  $dialogSuppressor = if ($PasswordToOpen -or $PasswordToModify) {
    Start-NUIDialogSuppressor -TargetExe 'POWERPNT.EXE'
  }
  else {
    $null
  }
  try {
    # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open
    return $Application.Presentations.Open(
      "$resolved::$passwordToOpenString::$passwordToModifyString" # FileName
      , $readOnlyValue                                            # ReadOnly
      , [type]::Missing                                           # Untitled
      , [MsoTriState]::msoFalse                                   # WithWindow
    )
  }
  finally {
    try {
      if ($dialogSuppressor) {
        Stop-NUIDialogSuppressor -Job $dialogSuppressor
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
#endregion
#region Public
function New-PowerPointFile {
  <#
  .SYNOPSIS
    Creates a new PowerPoint presentation file.

  .DESCRIPTION
    Creates a file at the specified path by automating PowerPoint through COM.

    If the destination file already exists, the command stops unless `-Force` is specified.
    You can use `-WhatIf` and `-Confirm` to preview or confirm the file creation or overwrite operation.

  .PARAMETER Path
    Specifies the destination path of the presentation file to create.

    This parameter does not support wildcards because it represents a new file path.

  .PARAMETER FileFormat
    Specifies the PowerPoint file format used when saving the presentation.

    The default value is `ppSaveAsOpenXMLPresentation`.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  Pass a `SecureString` value. If omitted, no open password is set.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

    Pass a `SecureString` value. If omitted, no modify password is set.

  .PARAMETER ReadOnlyRecommended
    Saves the presentation with the read-only recommended flag set.

    When this switch is specified, PowerPoint recommends that users open the presentation as read-only.

  .PARAMETER Force
    Overwrites an existing file at `-Path`.

    Without this switch, the cmdlet stops when the destination file already exists.

  .PARAMETER RemovePersonalInformation
    Removes personal information from the new presentation file.

  .EXAMPLE
    New-PowerPointFile -Path "$env:TEMP\Presentation.pptx"

    Creates a new presentation.

  .EXAMPLE
    New-PowerPointFile -Path "$env:TEMP\Presentation.pptm" -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -Force

    Creates or overwrites a macro-enabled presentation.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    New-PowerPointFile -Path "$env:TEMP\Presentation.pptx" -PasswordToOpen $password

    Creates a file protected with an open password.

  .EXAMPLE
    New-PowerPointFile -Path "$env:TEMP\Presentation.pptx" -ReadOnlyRecommended

    Creates a file that recommends opening as read-only.

  .OUTPUTS
    System.IO.FileInfo
      Returns the created presentation.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Path,
    # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
    [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]
    $FileFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentation,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $ReadOnlyRecommended,
    [switch]
    $Force,
    [switch]
    $RemovePersonalInformation
  )
  process {
    $passwordToOpenString = if ($null -eq $PasswordToOpen) {
      $null
    }
    else {
      [NetworkCredential]::new([string]::Empty, $PasswordToOpen).Password
    }
    $passwordToModifyString = if ($null -eq $PasswordToModify) {
      $null
    }
    else {
      [NetworkCredential]::new([string]::Empty, $PasswordToModify).Password
    }
    $resolved = [Path]::GetFullPath($Path)
    $exists = Test-Path -LiteralPath $resolved
    if ($exists -and -not $Force) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
    }
    $action = if ($exists) {
      'Overwrite PowerPoint file'
    }
    else {
      'Create PowerPoint file'
    }
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($resolved, $action))) {
      return
    }
    if ($exists) {
      Remove-Item -LiteralPath $resolved -Force -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $app = New-PowerPointObject
    try {
      # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.add
      $file = $app.Presentations.Add([MsoTriState]::msoFalse)
      try {
        if ($passwordToOpenString) {
          $file.Password = $passwordToOpenString
        }
        if ($passwordToModifyString) {
          $file.WritePassword = $passwordToModifyString
        }
        if ($RemovePersonalInformation) {
          $file.RemovePersonalInformation = [MsoTriState]::msoTrue
        }
        # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas2
        $file.SaveCopyAs2(
          $resolved                         # FileName
          , $FileFormat                     # FileFormat
          , [type]::Missing                 # EmbedTrueTypeFonts
          , $ReadOnlyRecommended.IsPresent  # ReadOnlyRecommended
        )
        return Get-Item -LiteralPath $resolved -Force
      }
      finally {
        $file.Close()
      }
    }
    finally {
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
}
function Get-PowerPointAppProperty {
  <#
  .SYNOPSIS
    Gets PowerPoint application properties.

  .DESCRIPTION
    Creates a PowerPoint Application COM object with setup disabled and returns its properties
    as a custom object by using `Get-ObjectProperty`.

    This cmdlet is intended for capturing a snapshot of current application-level settings
    that can be inspected or later passed to `Set-PowerPointAppProperty`.

  .EXAMPLE
    Get-PowerPointAppProperty

    Returns PowerPoint application properties.

  .EXAMPLE
    $properties = Get-PowerPointAppProperty
    $properties | Format-List

    Captures and inspects all available properties.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Returns a custom object that contains PowerPoint application properties.

  .NOTES
    The PowerPoint COM object is always closed and released after property retrieval.
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()
  $app = New-PowerPointObject -NoSetup
  try {
    Get-ObjectProperty -InputObject $app
  }
  finally {
    $app.Quit()
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Set-PowerPointAppProperty {
  <#
  .SYNOPSIS
    Sets PowerPoint application properties.

  .DESCRIPTION
    Creates a PowerPoint Application COM object with setup disabled and applies the specified
    properties to the object by using `Set-ObjectProperty`.

    This cmdlet is commonly used with an object captured by `Get-PowerPointAppProperty`.
    If `-Properties` contains a property that does not exist on the PowerPoint Application
    object, the cmdlet throws an error.

  .PARAMETER Properties
    Specifies the property set to apply to the PowerPoint Application object.

    Provide a `PSCustomObject` whose property names match writable properties on the
    PowerPoint Application COM object.

  .EXAMPLE
    $properties = Get-PowerPointAppProperty
    Set-PowerPointAppProperty -Properties $properties

    Applies a captured property snapshot back to PowerPoint.

  .EXAMPLE
    $properties = [PSCustomObject]@{ DisplayAlerts = 0; Visible = $false }
    Set-PowerPointAppProperty -Properties $properties

    Sets selected PowerPoint application properties.

  .OUTPUTS
    None.

  .NOTES
    The PowerPoint COM object is always released after applying properties.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory)]
    [PSObject]
    $Properties
  )
  if (-not $PSCmdlet.ShouldProcess('POWERPNT.EXE', 'Set PowerPoint application properties')) {
    return
  }
  try {
    $app = New-PowerPointObject -NoSetup
    Set-ObjectProperty -InputObject $app -Properties $Properties
  }
  finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Get-PowerPointFileProperty {
  <#
  .SYNOPSIS
    Gets presentation file properties.

  .DESCRIPTION
    Opens one or more presentations and returns file properties as a `PSCustomObject`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected presentations via `-PasswordToOpen` and `-PasswordToModify`.

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

  .EXAMPLE
    Get-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final

    Gets the `Final` file property from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-PowerPointFileProperty -LiteralPath "$env:TEMP\Presentation.pptx" -PasswordToOpen $password -Name Final

    Gets file properties from a file protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -PasswordToModify $password -Name Final

    Gets file properties from a file protected with a modify password.

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
    $PasswordToModify = $null
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
          $properties = Get-ObjectProperty -InputObject $file
          if ($Name) {
            $selected = [PSCustomObject]@{}
            foreach ($propertyName in $Name) {
              $property = $properties.PSObject.Properties |
              Where-Object -Property Name -EQ $propertyName |
              Select-Object -First 1
              if ($property) {
                $selected | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
              }
            }
            return $selected
          }
          return $properties
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
function Set-PowerPointFileProperty {
  <#
  .SYNOPSIS
    Sets presentation file properties.

  .DESCRIPTION
    Opens one or more presentations and updates a file property.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
    You can use `-WhatIf` and `-Confirm` to preview or confirm updates.

    If a target is read-only, the cmdlet throws an exception.

  .PARAMETER Path
    Specifies presentation paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies presentation paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the file property name to update.

  .PARAMETER Value
    Specifies the value assigned to the property.

  .PARAMETER InputObject
    Specifies a `PSObject` whose properties are mapped to file property names and values.
    Use this parameter to update multiple properties at once.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER PassThru
    Returns the updated file property object when specified.

  .EXAMPLE
    Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final -Value $true

    Sets a single file property.

  .EXAMPLE
    Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -InputObject ([PSCustomObject]@{ Final = $true })

    Updates multiple file properties in a single call.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final -Value $true -PasswordToOpen $password

    Updates file properties on a protected presentation.

  .EXAMPLE
    Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final -Value $true -PassThru

    Updates the property and returns the updated value.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      When `-PassThru` is specified.

    None.

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
        $item = $_
        $properties = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
          { $_ -in 'ValuePathSet', 'ValueLiteralPathSet' } {
            [PSCustomObject]@{ $Name = $Value }
          }
          { $_ -in 'PSObjectPathSet', 'PSObjectLiteralPathSet' } {
            $InputObject
          }
        }
        $target = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
          { $_ -in 'ValuePathSet', 'ValueLiteralPathSet' } {
            "Item: $($item.FullName) Property: $Name Value: $Value"
          }
          { $_ -in 'PSObjectPathSet', 'PSObjectLiteralPathSet' } {
            "Item: $($item.FullName) Properties: $($InputObject.PSObject.Properties.Name -join ', ')"
          }
        }
        if (-not $PSCmdlet.ShouldProcess($target, 'Set FileProperty')) {
          return
        }
        if ($item.IsReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $item))
        }
        $file = Open-PowerPointFile -Application $app -Path $item.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify
        try {
          if ($file.ReadOnly -ne [MsoTriState]::msoFalse) {
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $item))
          }
          Set-ObjectProperty -InputObject $file -Properties $properties | Out-Null
          if (-not $file.Saved) {
            $file.Save()
          }
          if ($PassThru) {
            $updated = Get-ObjectProperty -InputObject $file
            $propertyNames = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
              { $_ -in 'ValuePathSet', 'ValueLiteralPathSet' } {
                $Name
              }
              { $_ -in 'PSObjectPathSet', 'PSObjectLiteralPathSet' } {
                $InputObject.PSObject.Properties.Name
              }
            }
            return $updated | Select-Object -Property $propertyNames
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
function Test-PowerPointExtension {
  <#
  .SYNOPSIS
    Determines whether one or more files use a PowerPoint-compatible presentation extension.

  .DESCRIPTION
    Returns `$true` when any supplied path points to an existing file with a supported PowerPoint extension.
    Supports wildcard input through `-Path` and exact paths through `-LiteralPath`.

  .PARAMETER Path
    Specifies one or more file paths to evaluate. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies one or more exact file paths to evaluate.

  .EXAMPLE
    Test-PowerPointExtension -Path "$env:TEMP\Presentation.pptx"

    Returns true when the specified file exists and has a supported PowerPoint extension.

  .EXAMPLE
    Test-PowerPointExtension -LiteralPath "$env:TEMP\Presentation.ppsx"

    Returns true when the exact file path exists and is recognized as a PowerPoint presentation.

  .EXAMPLE
    Test-PowerPointExtension -Path "$env:TEMP\*.txt"

    Returns false when matched files do not use a supported PowerPoint extension.

  .OUTPUTS
    System.Boolean
      Returns `$true` if a supported PowerPoint file exists; otherwise, returns `$false`.
  #>
  [CmdletBinding(DefaultParameterSetName = 'PathSet')]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -IsValid })]
    [string[]]
    $LiteralPath
  )
  process {
    try {
      $items = @(
        switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
          'PathSet' {
            Get-Item -Path $Path -Force
          }
          'LiteralPathSet' {
            Get-Item -LiteralPath $LiteralPath -Force
          }
        }
      )
    }
    catch [ItemNotFoundException] {
      return $false
    }
    return @(
      $items |
      Where-Object {
        if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) {
          return $false
        }
        return [Path]::GetExtension($_.FullName) -imatch '^\.(odp|pptx|ppt|pptm|ppsx|pps|ppsm|potx|pot|potm|ppam|ppa)$'
      }
    ).Count -gt 0
  }
}
#endregion

using namespace Microsoft.Office.Interop.Word
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Net
using namespace System.Runtime.InteropServices

Add-Type -AssemblyName Microsoft.Office.Interop.Word
Set-StrictMode -Version Latest

#region Private
function New-WordObject {
  [CmdletBinding()]
  [OutputType([__ComObject])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for Word application and does not change persistent state')]
  param (
    [switch]
    $NoSetup
  )
  # https://learn.microsoft.com/en-us/office/vba/api/word.application
  while ($true) {
    try {
      $app = New-Object -ComObject Word.Application
      if ($app -and $app.Application) {
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
      # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdalertlevel?view=word-pia
      $app.DisplayAlerts = [WdAlertLevel]::wdAlertsNone
      $app.Visible = $false
      $app.ScreenUpdating = $false
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
function Open-WordFile {
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
  $dialogSuppressor = if ($PasswordToOpen -or $PasswordToModify) {
    Start-NUIDialogSuppressor -TargetExe 'WINWORD.EXE'
  }
  else {
    $null
  }
  try {
    # https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
    return $Application.Documents.Open(
      $resolved                 # FileName
      , [type]::Missing         # ConfirmConversions
      , $ReadOnly.IsPresent     # ReadOnly
      , [type]::Missing         # AddToRecentFiles
      , $passwordToOpenString   # PasswordDocument
      , [type]::Missing         # PasswordTemplate
      , [type]::Missing         # Revert
      , $passwordToModifyString # WritePasswordDocuments
      , [type]::Missing         # WritePasswordTemplate
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
function New-WordFile {
  <#
  .SYNOPSIS
    Creates a new Word document file.

  .DESCRIPTION
    Creates a file at the specified path by automating Word through COM.

    If the destination file already exists, the command stops unless `-Force` is specified.
    You can use `-WhatIf` and `-Confirm` to preview or confirm the file creation or overwrite operation.

  .PARAMETER Path
    Specifies the destination path of the document file to create.

    This parameter does not support wildcards because it represents a new file path.

  .PARAMETER FileFormat
    Specifies the Word file format used when saving the document.

    The default value is `wdFormatXMLDocument`.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the document.

    Pass a `SecureString` value. If omitted, no open password is set.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the document.

    Pass a `SecureString` value. If omitted, no modify password is set.

  .PARAMETER ReadOnlyRecommended
    Saves the document with the read-only recommended flag set.

    When this switch is specified, Word recommends that users open the document as read-only.

  .PARAMETER Force
    Overwrites an existing file at `-Path`.

    Without this switch, the cmdlet stops when the destination file already exists.

  .PARAMETER RemovePersonalInformation
    Removes personal information from the new document file.

  .PARAMETER Initialize
    Specifies a script block to initialize the document before saving.
    The script block receives the document object as its first argument.

  .EXAMPLE
    New-WordFile -Path "$env:TEMP\Document.docx"

    Creates a new document.

  .EXAMPLE
    New-WordFile -Path "$env:TEMP\Document.docm" -FileFormat wdFormatXMLDocumentMacroEnabled -Force

    Creates or overwrites a macro-enabled document.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    New-WordFile -Path "$env:TEMP\Document.docx" -PasswordToOpen $password

    Creates a file protected with an open password.

  .EXAMPLE
    New-WordFile -Path "$env:TEMP\Document.docx" -ReadOnlyRecommended

    Creates a file that recommends opening as read-only.

  .EXAMPLE
    New-WordFile -Path "$env:TEMP\Document.docx" -Initialize {
      param($Document)
      $Document.Range().Text = 'Hello, World!'
    }

    Creates a document and adds text to it.

  .OUTPUTS
    System.IO.FileInfo
      Returns the created document.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Path,
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat?view=word-pia
    [Microsoft.Office.Interop.Word.WdSaveFormat]
    $FileFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $ReadOnlyRecommended,
    [switch]
    $Force,
    [switch]
    $RemovePersonalInformation,
    [ScriptBlock]
    $Initialize
  )
  process {
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
    $resolved = [Path]::GetFullPath($Path)
    $exists = Test-Path -LiteralPath $resolved
    if ($exists -and -not $Force) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
    }
    $action = if ($exists) {
      'Overwrite Word file'
    }
    else {
      'Create Word file'
    }
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($resolved, $action))) {
      return
    }
    if ($exists) {
      Remove-Item -LiteralPath $resolved -Force -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $app = New-WordObject
    try {
      # https://learn.microsoft.com/en-us/office/vba/api/word.documents.add
      $file = $app.Documents.Add()
      try {
        if ($RemovePersonalInformation) {
          $file.RemovePersonalInformation = $true
        }
        if ($Initialize) {
          & $Initialize $file
        }
        # https://learn.microsoft.com/en-us/office/vba/api/word.saveas2
        $file.SaveAs2(
          $resolved                         # FileName
          , $FileFormat                     # FileFormat
          , [type]::Missing                 # LockComments
          , $passwordToOpenString           # Password
          , [type]::Missing                 # AddToRecentFiles
          , $passwordToModifyString         # WritePassword
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
function Get-WordAppProperty {
  <#
  .SYNOPSIS
    Gets Word application properties.

  .DESCRIPTION
    Creates a Word Application COM object with setup disabled and returns its properties
    as a custom object by using `Get-ObjectProperty`.

    This cmdlet is intended for capturing a snapshot of current application-level settings
    that can be inspected or later passed to `Set-WordAppProperty`.

  .EXAMPLE
    Get-WordAppProperty

    Returns Word application properties.

  .EXAMPLE
    $properties = Get-WordAppProperty
    $properties | Format-List

    Captures and inspects all available properties.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Returns a custom object that contains Word application properties.

  .NOTES
    The Word COM object is always closed and released after property retrieval.
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()
  $app = New-WordObject -NoSetup
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
function Set-WordAppProperty {
  <#
  .SYNOPSIS
    Sets Word application properties.

  .DESCRIPTION
    Creates a Word Application COM object with setup disabled and applies the specified
    properties to the object by using `Set-ObjectProperty`.

    This cmdlet is commonly used with an object captured by `Get-WordAppProperty`.
    If `-Properties` contains a property that does not exist on the Word Application
    object, the cmdlet throws an error.

  .PARAMETER Properties
    Specifies the property set to apply to the Word Application object.

    Provide a `PSCustomObject` whose property names match writable properties on the
    Word Application COM object.

  .EXAMPLE
    $properties = Get-WordAppProperty
    Set-WordAppProperty -Properties $properties

    Applies a captured property snapshot back to Word.

  .EXAMPLE
    $properties = [PSCustomObject]@{ DisplayAlerts = 0; Visible = $false }
    Set-WordAppProperty -Properties $properties

    Sets selected Word application properties.

  .OUTPUTS
    None.

  .NOTES
    The Word COM object is always released after applying properties.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory)]
    [PSObject]
    $Properties
  )
  if (-not $PSCmdlet.ShouldProcess('WINWORD.EXE', 'Set Word application properties')) {
    return
  }
  try {
    $app = New-WordObject -NoSetup
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
function Get-WordFileProperty {
  <#
  .SYNOPSIS
    Gets document file properties.

  .DESCRIPTION
    Opens one or more documents and returns file properties as a `PSCustomObject`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected documents via `-PasswordToOpen` and `-PasswordToModify`.

  .PARAMETER Path
    Specifies document paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies document paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Filters returned properties by name. If not specified, all properties are returned.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the document.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the document.

  .EXAMPLE
    Get-WordFileProperty -Path "$env:TEMP\Document.docx" -Name RemovePersonalInformation

    Gets the `RemovePersonalInformation` file property from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-WordFileProperty -LiteralPath "$env:TEMP\Document.docx" -PasswordToOpen $password -Name RemovePersonalInformation

    Gets file properties from a file protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-WordFileProperty -Path "$env:TEMP\Document.docx" -PasswordToModify $password -Name RemovePersonalInformation

    Gets file properties from a file protected with a modify password.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Each NoteProperty corresponds to a file property.

  .NOTES
    If a file cannot be opened (for example, an invalid `-PasswordToOpen` value), the cmdlet throws an exception and continues processing remaining items.
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
    $app = New-WordObject
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
        $file = Open-WordFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly
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
function Set-WordFileProperty {
  <#
  .SYNOPSIS
    Sets document file properties.

  .DESCRIPTION
    Opens one or more documents and updates a file property.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
    You can use `-WhatIf` and `-Confirm` to preview or confirm updates.

    If a target is read-only, the cmdlet throws an exception.

  .PARAMETER Path
    Specifies document paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies document paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the file property name to update.

  .PARAMETER Value
    Specifies the value assigned to the property.

  .PARAMETER InputObject
    Specifies a `PSObject` whose properties are mapped to file property names and values.
    Use this parameter to update multiple properties at once.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the document.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the document.

  .PARAMETER PassThru
    Returns the updated file property object when specified.

  .EXAMPLE
    Set-WordFileProperty -Path "$env:TEMP\Document.docx" -Name RemovePersonalInformation -Value $true

    Sets a single file property.

  .EXAMPLE
    Set-WordFileProperty -Path "$env:TEMP\Document.docx" -InputObject ([PSCustomObject]@{ RemovePersonalInformation = $true })

    Updates multiple file properties in a single call.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Set-WordFileProperty -Path "$env:TEMP\Document.docx" -Name RemovePersonalInformation -Value $true -PasswordToOpen $password

    Updates file properties on a protected document.

  .EXAMPLE
    Set-WordFileProperty -Path "$env:TEMP\Document.docx" -Name RemovePersonalInformation -Value $true -PassThru

    Updates the property and returns the updated value.

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
    $PassThru
  )
  begin {
    $app = New-WordObject
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
        $file = Open-WordFile -Application $app -Path $item.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify
        try {
          if ($file.ReadOnly) {
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
function Test-WordExtension {
  <#
  .SYNOPSIS
    Determines whether one or more files use a Word-compatible document extension.

  .DESCRIPTION
    Returns `$true` when any supplied path points to an existing file with a supported Word extension.
    Supports wildcard input through `-Path` and exact paths through `-LiteralPath`.

  .PARAMETER Path
    Specifies one or more file paths to evaluate. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies one or more exact file paths to evaluate.

  .EXAMPLE
    Test-WordExtension -Path "$env:TEMP\Document.docx"

    Returns true when the specified document exists and has a supported Word extension.

  .EXAMPLE
    Test-WordExtension -LiteralPath "$env:TEMP\Document.docm"

    Returns true when the exact file path exists and is recognized as a Word document.

  .EXAMPLE
    Test-WordExtension -Path "$env:TEMP\*.txt"

    Returns false when matched files do not use a supported Word extension.

  .OUTPUTS
    System.Boolean
      Returns `$true` if a supported Word file exists; otherwise, returns `$false`.
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
        return [Path]::GetExtension($_.FullName) -imatch '^\.(odt|docx|docm|doc|dotx|dotm|dot)$'
      }
    ).Count -gt 0
  }
}
#endregion

using namespace Microsoft.Office.Interop.Excel
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Net
using namespace System.Runtime.InteropServices

Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Set-StrictMode -Version Latest

#region Private
function New-ExcelObject {
  [CmdletBinding()]
  [OutputType([__ComObject])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for Excel application and does not change persistent state')]
  param (
    [switch]
    $NoSetup
  )
  # https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)
  while ($true) {
    try {
      $app = New-Object -ComObject Excel.Application
      if ($app -and $app.hWnd) {
        break
      }
    } catch [COMException] {
      Write-Warning -Message $_.Exception.Message
    }
    Start-Sleep -Milliseconds 100
  }
  try {
    if (-not $NoSetup) {
      $app.DisplayAlerts = $false
      $app.Visible = $false
      $app.EnableEvents = $false
      $app.ScreenUpdating = $false
    }
    return $app
  } catch {
    try {
      if ($app) {
        $app.Quit()
      }
    } finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
    $PSCmdlet.ThrowTerminatingError($_)
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
#endregion
#region Public
function New-ExcelFile {
  <#
  .SYNOPSIS
    Creates a new Excel workbook file.

  .DESCRIPTION
    Creates a file at the specified path by automating Excel through COM.

    If the destination file already exists, the command stops unless `-Force` is specified.
    You can use `-WhatIf` and `-Confirm` to preview or confirm the file creation or overwrite operation.

  .PARAMETER Path
    Specifies the destination path of the workbook file to create.

    This parameter does not support wildcards because it represents a new file path.

  .PARAMETER FileFormat
    Specifies the Excel file format used when saving the workbook.

    The default value is `xlWorkbookDefault`.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

    Pass a `SecureString` value. If omitted, no open password is set.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

    Pass a `SecureString` value. If omitted, no modify password is set.

  .PARAMETER ReadOnlyRecommended
    Saves the workbook with the read-only recommended flag set.

    When this switch is specified, Excel recommends that users open the workbook as read-only.

  .PARAMETER Force
    Overwrites an existing file at `-Path`.

    Without this switch, the cmdlet stops when the destination file already exists.

  .PARAMETER RemovePersonalInformation
    Removes personal information from the new workbook file.

  .PARAMETER Initialize
    Specifies a script block to initialize the workbook before saving.
    The script block receives the workbook object as its first argument.

  .EXAMPLE
    New-ExcelFile -Path "$env:TEMP\Book.xlsx"

    Creates a new workbook.

  .EXAMPLE
    New-ExcelFile -Path "$env:TEMP\Book.xlsm" -FileFormat xlOpenXMLWorkbookMacroEnabled -Force

    Creates or overwrites a macro-enabled workbook.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    New-ExcelFile -Path "$env:TEMP\Book.xlsx" -PasswordToOpen $password

    Creates a file protected with an open password.

  .EXAMPLE
    New-ExcelFile -Path "$env:TEMP\Book.xlsx" -ReadOnlyRecommended

    Creates a file that recommends opening as read-only.

  .EXAMPLE
    New-ExcelFile -Path "$env:TEMP\Book.xlsx" -Initialize {
      param($Workbook)
      $Workbook.Worksheets.Item(1).Name = 'Data'
    }

    Creates a workbook and renames the first worksheet.

  .OUTPUTS
    System.IO.FileInfo
      Returns the created workbook.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Path,
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
    [Microsoft.Office.Interop.Excel.XlFileFormat]
    $FileFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook,
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
    } else {
      [NetworkCredential]::new([string]::Empty, $PasswordToOpen).Password
    }
    $passwordToModifyString = if ($null -eq $PasswordToModify) {
      [type]::Missing
    } else {
      [NetworkCredential]::new([string]::Empty, $PasswordToModify).Password
    }
    $resolved = [Path]::GetFullPath($Path)
    $exists = Test-Path -LiteralPath $resolved
    if ($exists -and -not $Force) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
    }
    $action = if ($exists) {
      'Overwrite Excel file'
    } else {
      'Create Excel file'
    }
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($resolved, $action))) {
      return
    }
    if ($exists) {
      Remove-Item -LiteralPath $resolved -Force -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $app = New-ExcelObject
    try {
      # https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.add
      $file = $app.Workbooks.Add()
      try {
        if ($RemovePersonalInformation) {
          $file.RemovePersonalInformation = $true
        }
        if ($Initialize) {
          & $Initialize $file
        }
        # https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
        $file.SaveAs(
          $resolved                        # FileName
          , $FileFormat                    # FileFormat
          , $passwordToOpenString          # Password
          , $passwordToModifyString        # WriteResPassword
          , $ReadOnlyRecommended.IsPresent # ReadOnlyRecommended
        )
        return Get-Item -LiteralPath $resolved -Force
      } finally {
        $file.Close()
      }
    } finally {
      try {
        if ($app) {
          $app.Quit()
        }
      } finally {
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
    }
  }
}
function Open-ExcelFile {
  <#
  .SYNOPSIS
    Opens an Excel workbook file.

  .DESCRIPTION
    Opens a workbook file by automating Excel through COM.

    If `-Application` is not specified, the cmdlet creates a new Excel Application
    object, opens the workbook, and releases the application after the action completes.

  .PARAMETER Path
    Specifies the path to the workbook file to open.

    This parameter does not support wildcards because it represents an existing file path.

  .PARAMETER Application
    Specifies the Excel Application COM object to use for opening the workbook.

    If not specified, a new Excel Application object is created automatically.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

    Pass a `SecureString` value. If omitted, no open password is used.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

    Pass a `SecureString` value. If omitted, no modify password is used.

  .PARAMETER ReadOnly
    Opens the workbook in read-only mode.

  .PARAMETER Force
    Ignores the read-only recommended flag when opening the workbook.

  .PARAMETER Action
    Specifies a script block to execute with the opened workbook.

    The script block receives the workbook object as its first argument.
    When this parameter is specified, the workbook is automatically closed after
    the action completes.

  .EXAMPLE
    Open-ExcelFile -Path "$env:TEMP\Book.xlsx"

    Opens a workbook and returns the workbook object.

  .EXAMPLE
    Open-ExcelFile -Path "$env:TEMP\Book.xlsx" -Action {
      param($Workbook)
      $Workbook.Worksheets.Item(1).Name = 'Sheet1'
    }

    Opens a workbook, renames the first worksheet, and closes the workbook.

  .EXAMPLE
    $app = New-ExcelObject
    Open-ExcelFile -Application $app -Path "$env:TEMP\Book.xlsx" -Action {
      param($Workbook)
      $Workbook.Worksheets.Item(1).Name = 'Sheet1'
    }
    $app.Quit()

    Opens a workbook using an existing Excel Application object.

  .OUTPUTS
    __ComObject
      Returns the opened workbook object.
  #>
  [OutputType([__ComObject])]
  param(
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [__ComObject]
    $Application,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $ReadOnly,
    [switch]
    $Force,
    [ScriptBlock]
    $Action
  )
  process {
    $app = if ($Application) {
      $Application
    } else {
      New-ExcelObject
    }
    $shouldDisposeApp = -not $Application
    try {
      $resolved = (Resolve-Path -LiteralPath $Path).Path
      $passwordToOpenString = if ($null -eq $PasswordToOpen) {
        [type]::Missing
      } else {
        [NetworkCredential]::new([string]::Empty, $PasswordToOpen).Password
      }
      $passwordToModifyString = if ($null -eq $PasswordToModify) {
        [type]::Missing
      } else {
        [NetworkCredential]::new([string]::Empty, $PasswordToModify).Password
      }
      $dialogSuppressor = if ($PasswordToOpen -or $PasswordToModify) {
        Start-NUIDialogSuppressor -TargetExe 'EXCEL.EXE'
      } else {
        $null
      }
      try {
        # https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
        $file = $app.Workbooks.Open(
          $resolved                  # FileName
          , [type]::Missing         # UpdateLinks
          , $ReadOnly.IsPresent     # ReadOnly
          , [type]::Missing         # Format
          , $passwordToOpenString   # Password
          , $passwordToModifyString # WriteResPassword
          , $Force.IsPresent        # IgnoreReadOnlyRecommended
        )
        try {
          if ($Action) {
            return & $Action $file
          }
        } finally {
          $file.Close()
        }
      } finally {
        try {
          if ($dialogSuppressor) {
            Stop-NUIDialogSuppressor -Job $dialogSuppressor
          }
        } finally {
          Get-Variable |
          Where-Object -Property Value -Is [__ComObject] |
          Clear-Variable -Force -WhatIf:$false -Confirm:$false
          [GC]::Collect()
          [GC]::WaitForPendingFinalizers()
        }
      }
    } finally {
      if ($shouldDisposeApp -and $app) {
        $app.Quit()
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
    }
  }
}
function Get-ExcelAppProperty {
  <#
  .SYNOPSIS
    Gets Excel application properties.

  .DESCRIPTION
    Creates an Excel Application COM object with setup disabled and returns its properties
    as a custom object by using `Get-ObjectProperty`.

    This cmdlet is intended for capturing a snapshot of current application-level settings
    that can be inspected or later passed to `Set-ExcelAppProperty`.

  .EXAMPLE
    Get-ExcelAppProperty

    Returns Excel application properties.

  .EXAMPLE
    $properties = Get-ExcelAppProperty
    $properties | Format-List

    Captures and inspects all available properties.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Returns a custom object that contains Excel application properties.

  .NOTES
    The Excel COM object is always closed and released after property retrieval.
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()
  $app = New-ExcelObject -NoSetup
  try {
    Get-ObjectProperty -InputObject $app
  } finally {
    $app.Quit()
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Set-ExcelAppProperty {
  <#
  .SYNOPSIS
    Sets Excel application properties.

  .DESCRIPTION
    Creates an Excel Application COM object with setup disabled and applies the specified
    properties to the object by using `Set-ObjectProperty`.

    This cmdlet is commonly used with an object captured by `Get-ExcelAppProperty`.
    If `-Properties` contains a property that does not exist on the Excel Application
    object, the cmdlet throws an error.

  .PARAMETER Properties
    Specifies the property set to apply to the Excel Application object.

    Provide a `PSCustomObject` whose property names match writable properties on the
    Excel Application COM object.

  .EXAMPLE
    $properties = Get-ExcelAppProperty
    Set-ExcelAppProperty -Properties $properties

    Applies a captured property snapshot back to Excel.

  .EXAMPLE
    $properties = [PSCustomObject]@{ DisplayAlerts = $false; Visible = $false }
    Set-ExcelAppProperty -Properties $properties

    Sets selected Excel application properties.

  .OUTPUTS
    None.

  .NOTES
    The Excel COM object is always released after applying properties.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory)]
    [PSObject]
    $Properties
  )
  if (-not $PSCmdlet.ShouldProcess('EXCEL.EXE', 'Set Excel application properties')) {
    return
  }
  try {
    $app = New-ExcelObject -NoSetup
    Set-ObjectProperty -InputObject $app -Properties $Properties
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Get-ExcelFileProperty {
  <#
  .SYNOPSIS
    Gets workbook file properties.

  .DESCRIPTION
    Opens one or more workbooks and returns file properties as a `PSCustomObject`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected workbooks via `-PasswordToOpen` and `-PasswordToModify`.

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

  .EXAMPLE
    Get-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -Name SaveLinkValues

    Gets the `SaveLinkValues` workbook property from a file.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-ExcelFileProperty -LiteralPath "$env:TEMP\Book.xlsx" -PasswordToOpen $password -Name SaveLinkValues

    Gets file properties from a file protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -PasswordToModify $password -Name SaveLinkValues

    Gets file properties from a file protected with a modify password.

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
    $PasswordToModify = $null
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
        Open-ExcelFile -Application $app -Path $_.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly -Action {
          param(
            [Parameter(Mandatory)]
            [Workbook]
            $file
          )
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
      }
    } catch {
      try {
        if ($app) {
          $app.Quit()
        }
      } finally {
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
    } finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
}
function Set-ExcelFileProperty {
  <#
  .SYNOPSIS
    Sets workbook file properties.

  .DESCRIPTION
    Opens one or more workbooks and updates workbook file properties by using `Set-ObjectProperty`.

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
    Specifies the workbook property name to update.

  .PARAMETER Value
    Specifies the value assigned to the property.

  .PARAMETER InputObject
    Specifies a `PSObject` whose properties are mapped to workbook property names and values.
    Use this parameter to update multiple properties at once.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

  .PARAMETER PassThru
    Returns the updated workbook property object when specified.

  .PARAMETER Force
    Forces update processing when a file is marked as read-only recommended.

  .EXAMPLE
    Set-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -Name SaveLinkValues -Value $false

    Sets a single workbook property.

  .EXAMPLE
    Set-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -InputObject ([PSCustomObject]@{ SaveLinkValues = $false; CheckCompatibility = $false })

    Updates multiple workbook properties in a single call.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Set-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -Name SaveLinkValues -Value $false -PasswordToOpen $password -Force

    Updates file properties on a protected workbook.

  .EXAMPLE
    Set-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -Name SaveLinkValues -Value $false -PassThru

    Updates the property and returns the updated value.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      When `-PassThru` is specified.

    None.

  .NOTES
    The workbook and Excel COM objects are always closed and released after applying properties.
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
        if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($target, 'Set FileProperty'))) {
          return
        }
        if ($item.IsReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $item))
        }
        Open-ExcelFile -Application $app -Path $item.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Force:$Force -Action {
          param(
            [Parameter(Mandatory)]
            [Workbook]
            $file
          )
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
      }
    } catch {
      try {
        if ($app) {
          $app.Quit()
        }
      } finally {
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
    } finally {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
    }
  }
}
function Test-ExcelExtension {
  <#
  .SYNOPSIS
    Checks whether the specified file paths are Excel-compatible documents.

  .DESCRIPTION
    Returns `$true` when any supplied path points to an existing file with a supported Excel extension.
    Supports wildcard input through `-Path` and exact paths through `-LiteralPath`.

  .PARAMETER Path
    Specifies one or more file paths to evaluate. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies one or more exact file paths to evaluate.

  .EXAMPLE
    Test-ExcelExtension -Path "$env:TEMP\Book.xlsx"

    Returns true when the specified file exists and has a supported Excel extension.

  .EXAMPLE
    Test-ExcelExtension -LiteralPath "$env:TEMP\Book.xlsx"

    Returns true when the exact file path exists and is recognized as a Excel presentation.

  .EXAMPLE
    Test-ExcelExtension -Path "$env:TEMP\*.txt"

    Returns false when matched files do not use a supported Excel extension.

  .OUTPUTS
    System.Boolean
      Returns `$true` if a supported Excel file exists; otherwise, returns `$false`.
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
    } catch [ItemNotFoundException] {
      return $false
    }
    return @(
      $items |
      Where-Object {
        if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) {
          return $false
        }
        return [Path]::GetExtension($_.FullName) -imatch '^\.(ods|xl|xlsx|xlsb|xlam|xltx|xltm|xls|xlt|xlm|xlw)$'
      }
    ).Count -gt 0
  }
}
#endregion

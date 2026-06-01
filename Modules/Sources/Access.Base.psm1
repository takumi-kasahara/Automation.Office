using namespace Microsoft.Office.Interop.Access
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Net
using namespace System.Runtime.InteropServices

Add-Type -AssemblyName Microsoft.Office.Interop.Access
Set-StrictMode -Version Latest

#region Private
function New-AccessObject {
  [CmdletBinding()]
  [OutputType([__ComObject])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for Access application and does not change persistent state')]
  param ()
  # https://learn.microsoft.com/en-us/office/vba/api/access.application
  while ($true) {
    try {
      $app = New-Object -ComObject Access.Application
      if ($app -and $app.hWndAccessApp) {
        break
      }
    }
    catch [COMException] {
      Write-Warning -Message $_.Exception.Message
    }
    Start-Sleep -Milliseconds 100
  }
  try {
    $app.Visible = $false
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
function Open-AccessFile {
  [CmdletBinding()]
  [OutputType([void])]
  param(
    [Parameter(Mandatory)]
    [__ComObject]
    $Application,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null
  )
  $resolved = (Resolve-Path -LiteralPath $Path).Path
  $passwordString = if ($null -eq $Password) {
    $null
  }
  else {
    [NetworkCredential]::new([string]::Empty, $Password).Password
  }
  if ($Password) {
    # Validate credentials through DAO first to avoid Access password UI prompts.
    # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-opendatabase-method-dao
    $Application.DBEngine.OpenDatabase(
      $resolved                 # Name
      , $false                  # Options
      , $true                   # ReadOnly
      , ";PWD=$passwordString"  # Connect
    ).Close()
  }
  # https://learn.microsoft.com/en-us/office/vba/api/access.application.opencurrentdatabase
  $Application.OpenCurrentDatabase(
    $resolved         # filePath
    , $true           # exclusive
    , $passwordString # bstrPassword
  )
}
#endregion
#region Public
function New-AccessFile {
  <#
  .SYNOPSIS
    Creates a new Access database file.

  .DESCRIPTION
    Creates a database at the specified path by automating Access through COM.

    If the destination file already exists, the command stops unless `-Force` is specified.
    Because this cmdlet supports `ShouldProcess`, you can use `-WhatIf` and `-Confirm`
    to preview or confirm the file creation or overwrite operation.

  .PARAMETER Path
    Specifies the destination path of the database file to create.

    This parameter does not support wildcards because it represents a new file path.

  .PARAMETER FileFormat
    Specifies the Access file format used when creating the database.

    The default value is `acNewDatabaseFormatUserDefault`.

  .PARAMETER Password
    Specifies the password required to open the database.

    Pass a `SecureString` value. If omitted, no database password is set.

  .PARAMETER Force
    Overwrites an existing file at `-Path`.

    Without this switch, the cmdlet stops when the destination file already exists.

  .PARAMETER RemovePersonalInformation
    Removes personal information from the new database file.

  .EXAMPLE
    New-AccessFile -Path "$env:TEMP\Database.accdb"

    Creates a new database in the temporary directory.

  .EXAMPLE
    New-AccessFile -Path "$env:TEMP\Database.mdb" -FileFormat acNewDatabaseFormatAccess2000 -Force

    Creates or overwrites a legacy Access database.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    New-AccessFile -Path "$env:TEMP\Database.accdb" -Password $password

    Creates a database protected with an open password.

  .EXAMPLE
    New-AccessFile -Path "$env:TEMP\Database.accdb" -RemovePersonalInformation

    Creates a new database and removes personal information from the file.

  .OUTPUTS
    System.IO.FileInfo
      Returns the created database.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Path,
    # https://learn.microsoft.com/en-us/office/vba/api/access.acnewdatabaseformat
    [Microsoft.Office.Interop.Access.AcNewDatabaseFormat]
    $FileFormat = [Microsoft.Office.Interop.Access.AcNewDatabaseFormat]::acNewDatabaseFormatUserDefault,
    [SecureString]
    $Password = $null,
    [switch]
    $Force,
    [switch]
    $RemovePersonalInformation
  )
  process {
    $passwordString = if ($null -eq $Password) {
      $null
    }
    else {
      [NetworkCredential]::new([string]::Empty, $Password).Password
    }
    $resolved = [Path]::GetFullPath($Path)
    $exists = Test-Path -LiteralPath $resolved
    if ($exists -and -not $Force) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
    }
    $action = if ($exists) {
      'Overwrite Access file'
    }
    else {
      'Create Access file'
    }
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($resolved, $action))) {
      return
    }
    if ($exists) {
      Remove-Item -LiteralPath $resolved -Force -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $app = New-AccessObject
    try {
      # https://learn.microsoft.com/en-us/office/vba/api/access.application.newcurrentdatabase
      $app.NewCurrentDatabase(
        $resolved     # FilePath
        , $FileFormat # FileFormat
      )
      if ($passwordString) {
        # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-opendatabase-method-dao
        $database = $app.DBEngine.OpenDatabase(
          $resolved         # Name
          , $false          # Options
          , $false          # ReadOnly
          , [string]::Empty # Connect
        )
        try {
          # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/database-newpassword-method-dao
          $database.NewPassword(
            [string]::Empty   # bstrOld
            , $passwordString # bstrNew
          )
        }
        finally {
          $database.Close()
        }
      }
      # https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.removepersonalinformation
      if ($RemovePersonalInformation) {
        $app.CurrentProject.RemovePersonalInformation = $true
      }
      return Get-Item -LiteralPath $resolved -Force
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
function Get-AccessAppProperty {
  <#
  .SYNOPSIS
    Gets Access application properties.

  .DESCRIPTION
    Creates an Access Application COM object and returns its properties as a `PSCustomObject`.

    This cmdlet is intended for capturing a snapshot of current application-level settings
    that can be inspected or later passed to `Set-AccessAppProperty`.

  .EXAMPLE
    Get-AccessAppProperty

    Returns Access application properties.

  .EXAMPLE
    $properties = Get-AccessAppProperty
    $properties | Format-List

    Captures and inspects all available properties.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Returns a custom object that contains Access application properties.

  .NOTES
    The Access COM object is always closed and released after property retrieval.
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()
  $app = New-AccessObject
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
function Set-AccessAppProperty {
  <#
  .SYNOPSIS
    Sets Access application properties.

  .DESCRIPTION
    Creates an Access Application COM object and applies the specified properties to the object
    by using `Set-ObjectProperty`.

    This cmdlet is commonly used with an object captured by `Get-AccessAppProperty`.
    If `-Properties` contains a property that does not exist on the Access Application object,
    the cmdlet throws an error.

  .PARAMETER Properties
    Specifies the property set to apply to the Access Application object.

    Provide a `PSCustomObject` whose property names match writable properties on the
    Access Application COM object.

  .EXAMPLE
    $properties = Get-AccessAppProperty
    Set-AccessAppProperty -Properties $properties

    Applies a captured property snapshot back to Access.

  .EXAMPLE
    $properties = [PSCustomObject]@{ Visible = $false }
    Set-AccessAppProperty -Properties $properties

    Sets selected Access application properties.

  .OUTPUTS
    None.

  .NOTES
    The Access COM object is always released after applying properties.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory)]
    [PSObject]
    $Properties
  )
  if (-not $PSCmdlet.ShouldProcess('MSACCESS.EXE', 'Set Access application properties')) {
    return
  }
  try {
    $app = New-AccessObject
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
function Get-AccessFileProperty {
  <#
  .SYNOPSIS
    Gets database file properties.

  .DESCRIPTION
    Opens one or more databases and returns database file properties as a `PSCustomObject`.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected databases via `-PasswordToOpen`.

  .PARAMETER Path
    Specifies database paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies database paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Filters returned properties by name. If not specified, all file properties are returned.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the database.

  .EXAMPLE
    Get-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation

    Gets the `RemovePersonalInformation` file property from a database.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-AccessFileProperty -LiteralPath "$env:TEMP\Database.accdb" -PasswordToOpen $password -Name RemovePersonalInformation

    Gets file properties from a database protected with a password.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      Each NoteProperty corresponds to a file property.

  .NOTES
    If a database cannot be opened (for example, incorrect password), the cmdlet throws an error record and continues processing remaining items.
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
    $PasswordToOpen = $null
  )
  begin {
    $app = New-AccessObject
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
        Open-AccessFile -Application $app -Path $_.FullName -Password $PasswordToOpen
        try {
          $properties = Get-ObjectProperty -InputObject $app.CurrentProject
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
          $app.CloseCurrentDatabase()
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
function Set-AccessFileProperty {
  <#
  .SYNOPSIS
    Sets database file properties.

  .DESCRIPTION
    Opens one or more databases and updates a file property.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
    Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
    You can use `-WhatIf` and `-Confirm` to preview or confirm updates.

    If a target is read-only, the cmdlet throws an exception.

  .PARAMETER Path
    Specifies database paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies database paths literally. Wildcards are not interpreted.

  .PARAMETER Name
    Specifies the file property name to update.

  .PARAMETER Value
    Specifies the value assigned to the property.

  .PARAMETER InputObject
    Specifies a `PSObject` whose properties are mapped to file property names and values.
    Use this parameter to update multiple properties at once.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the database.

  .PARAMETER PassThru
    Returns the updated file property object when specified.

  .EXAMPLE
    Set-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation -Value $true

    Sets a single file property.

  .EXAMPLE
    Set-AccessFileProperty -Path "$env:TEMP\Database.accdb" -InputObject ([PSCustomObject]@{ RemovePersonalInformation = $true })

    Updates multiple file properties in a single call.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Set-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation -Value $true -PasswordToOpen $password

    Updates file properties on a protected database.

  .EXAMPLE
    Set-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation -Value $true -PassThru

    Updates the property and returns the updated value.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
      When `-PassThru` is specified.

    None.

  .NOTES
    If a database cannot be updated (for example, read-only state, invalid property name, or file access issues), the cmdlet throws an exception and continues processing remaining items.
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
    [switch]
    $PassThru
  )
  begin {
    $app = New-AccessObject
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
      $items | ForEach-Object {
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
        Open-AccessFile -Application $app -Path $item.FullName -Password $PasswordToOpen
        try {
          Set-ObjectProperty -InputObject $app.CurrentProject -Properties $properties | Out-Null
          if ($PassThru) {
            $updated = Get-ObjectProperty -InputObject $app.CurrentProject
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
          $app.CloseCurrentDatabase()
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
#endregion

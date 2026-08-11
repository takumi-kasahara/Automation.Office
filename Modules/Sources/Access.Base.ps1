using assembly Microsoft.Office.Interop.Access
using module .\Base.psm1
using namespace Microsoft.Office.Interop.Access
using namespace Microsoft.Office.Interop.Access.Dao
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Net
using namespace System.Runtime.InteropServices

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
    } catch [COMException] {
      Write-Warning -Message $_.Exception.Message
    }
    Start-Sleep -Milliseconds 100
  }
  try {
    $app.Visible = $false
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
  }
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

  .PARAMETER InitializeDb
    Specifies a script block to initialize the database before saving.
    The script block receives the DAO database object as its first argument.

  .PARAMETER InitializeProject
    Specifies a script block to initialize the database project before saving.
    The script block receives the CurrentProject object as its first argument.

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

  .EXAMPLE
    New-AccessFile -Path "$env:TEMP\Database.accdb" -InitializeDb {
      param($Database)
      $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
    }

    Creates a database and creates a table using DAO.

  .EXAMPLE
    New-AccessFile -Path "$env:TEMP\Database.accdb" -InitializeProject {
      param($Project)
      return $Project.Connection
    }

    Returns the connection string of the new database project.

  .NOTES
    The -Password and -RemovePersonalInformation parameters cannot be specified together.
    Setting RemovePersonalInformation triggers a save operation that displays a password dialog,
    which may cause the cmdlet to hang if a password is required for a password-protected databasetirmation dialog in password-protected databases, causing the cmdlet to hang.
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
    $RemovePersonalInformation,
    [ScriptBlock]
    $InitializeDb,
    [ScriptBlock]
    $InitializeProject
  )
  process {
    if ($null -ne $Password -and $RemovePersonalInformation.IsPresent) {
      throw '-Password and -RemovePersonalInformation parameters cannot be specified together.'
    }
    $passwordString = if ($null -eq $Password) {
      $null
    } else {
      [NetworkCredential]::new([string]::Empty, $Password).Password
    }
    $resolved = [Path]::GetFullPath($Path)
    $exists = Test-Path -LiteralPath $resolved
    if ($exists -and -not $Force) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
    }
    $action = if ($exists) {
      'Overwrite Access file'
    } else {
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
      $dialogSuppressor = Start-DialogSuppressor -TargetExe 'MSACCESS.EXE'
      if ($Password) {
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
        } finally {
          $database.Close()
        }
      }
      # https://learn.microsoft.com/en-us/office/vba/api/access.currentproject.removepersonalinformation
      if ($RemovePersonalInformation) {
        $app.CurrentProject.RemovePersonalInformation = $true
      }
      if ($InitializeDb) {
        $database = $app.CurrentDb()
        try {
          & $InitializeDb $database
        } finally {
          $database.Close()
        }
      }
      if ($InitializeProject) {
        & $InitializeProject $app.CurrentProject
      }
      return Get-Item -LiteralPath $resolved -Force
    } finally {
      try {
        if ($app) {
          $app.CloseCurrentDatabase()
          $app.Quit()
        }
        if ($dialogSuppressor) {
          Stop-DialogSuppressor -Job $dialogSuppressor
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
function Open-AccessFile {
  <#
  .SYNOPSIS
    Opens an Access database file.

  .DESCRIPTION
    Opens a database file by automating Access through COM.

    If `-Application` is not specified, the cmdlet creates a new Access Application
    object, opens the database, and releases the application after the action completes.

  .PARAMETER Path
    Specifies the path to the database file to open.

    This parameter does not support wildcards because it represents an existing file path.

  .PARAMETER Application
    Specifies the Access Application COM object to use for opening the database.

    If not specified, a new Access Application object is created automatically.

  .PARAMETER Password
    Specifies the password required to open the database.

    Pass a `SecureString` value. If omitted, no password is used.

  .PARAMETER ActionDb
    Specifies a script block to execute with the DAO database object.

    The script block receives the database object as its first argument.
    When this parameter is specified, the database is automatically closed after
    the action completes.

  .PARAMETER ActionProject
    Specifies a script block to execute with the CurrentProject object.

    The script block receives the CurrentProject object as its first argument.
    When this parameter is specified, the database is automatically closed after
    the action completes.

  .EXAMPLE
    Open-AccessFile -Path "$env:TEMP\Database.accdb"

    Opens a database and returns the CurrentProject object.

  .EXAMPLE
    Open-AccessFile -Path "$env:TEMP\Database.accdb" -ActionDb {
      param($Database)
      $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
    }

    Opens a database, creates a table using DAO, and closes the database.

  .EXAMPLE
    Open-AccessFile -Path "$env:TEMP\Database.accdb" -ActionProject {
      param($Project)
      return $Project.Connection
    }

    Opens a database and returns the connection string.

  .OUTPUTS
    __ComObject
      Returns the opened CurrentProject object when no Action parameter is specified.

  .NOTES
    If neither ActionDb nor ActionProject is specified, do not return anything.
  #>
  [CmdletBinding()]
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
    $Password = $null,
    [ScriptBlock]
    $ActionDb,
    [ScriptBlock]
    $ActionProject
  )
  process {
    $app = if ($Application) {
      $Application
    } else {
      New-AccessObject
    }
    try {
      $resolved = (Resolve-Path -LiteralPath $Path).Path
      $passwordString = if ($null -eq $Password) {
        $null
      } else {
        [NetworkCredential]::new([string]::Empty, $Password).Password
      }
      try {
        # Validate credentials through DAO first to avoid Access password UI prompts.
        # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-opendatabase-method-dao
        $app.DBEngine.OpenDatabase(
          $resolved                 # Name
          , $false                  # Options
          , $true                   # ReadOnly
          , ";PWD=$passwordString"  # Connect
        ).Close()
        # https://learn.microsoft.com/en-us/office/vba/api/access.application.opencurrentdatabase
        $app.OpenCurrentDatabase(
          $resolved         # filePath
          , $true           # exclusive
          , $passwordString # bstrPassword
        )
        if ($ActionDb) {
          $db = $app.CurrentDb()
          try {
            & $ActionDb $db
          } finally {
            $db.Close()
          }
        }
        if ($ActionProject) {
          & $ActionProject $app.CurrentProject
        }
        if (-not $Application) {
          $app.CloseCurrentDatabase()
        }
      } finally {
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
    } finally {
      if ($app) {
        try {
          $app.Quit()
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

  .PARAMETER Password
    Specifies the password required to open the database.

  .EXAMPLE
    Get-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation

    Gets the `RemovePersonalInformation` file property from a database.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-AccessFileProperty -LiteralPath "$env:TEMP\Database.accdb" -Password $password -Name RemovePersonalInformation

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
    $Password = $null
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
        Open-AccessFile -Application $app -Path $_.FullName -Password $Password
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
        } finally {
          $app.CloseCurrentDatabase()
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

  .PARAMETER Password
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
    Set-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation -Value $true -Password $password

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
    $Password = $null,
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
        Open-AccessFile -Application $app -Path $item.FullName -Password $Password
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
        } finally {
          $app.CloseCurrentDatabase()
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
function Export-AccessDatabase {
  <#
  .SYNOPSIS
    Exports data from an Access database table to a text file or spreadsheet.

  .DESCRIPTION
    Exports the contents of a table from an Access database to a text file (CSV, delimited, HTML) or spreadsheet (Excel) by automating Access through COM.

    The cmdlet supports two parameter sets:
    - **TextSet** (default): Exports to text formats using `DoCmd.TransferText`. Supports delimited, fixed-width, and HTML export types.
    - **SpreadsheetSet**: Exports to Excel formats using `DoCmd.TransferSpreadsheet`.

    If the destination file already exists, the cmdlet stops unless `-Force` is specified. With `-Force`, the existing file is overwritten. If `-NoClobber` is specified, the cmdlet throws an error when the destination exists.

    Because this cmdlet supports `ShouldProcess`, you can use `-WhatIf` and `-Confirm` to preview or confirm the export operation.

  .PARAMETER Path
    Specifies the path to the Access database file to export from.

    This parameter does not support wildcards because it represents an existing file path.

  .PARAMETER Password
    Specifies the password required to open the database.

    Pass a `SecureString` value. If omitted, no password is used.

  .PARAMETER TableName
    Specifies the name of the table to export.

  .PARAMETER Destination
    Specifies the destination file path for the exported data.

    This parameter does not support wildcards because it represents a new file path.

  .PARAMETER Force
    Overwrites an existing file at `-Destination`.

    Without this switch, the cmdlet stops when the destination file already exists.

  .PARAMETER NoClobber
    Prevents overwriting an existing file at `-Destination`.

    If the destination file exists, the cmdlet throws an error. This switch takes precedence over `-Force`.

  .PARAMETER TransferType
    Specifies the text transfer type for the export.

    This parameter is only valid with the **TextSet** parameter set. The default value is `acExportDelim`.

  .PARAMETER CodePage
    Specifies the code page to use for the exported text file.

    This parameter is only valid with the **TextSet** parameter set. The default value is `1200` (Unicode).

  .PARAMETER SpreadsheetType
    Specifies the spreadsheet type for the export.

    This parameter is only valid with the **SpreadsheetSet** parameter set. The default value is `acSpreadsheetTypeExcel12Xml`.

  .PARAMETER Range
    Specifies the range of cells to export in the spreadsheet.

    This parameter is only valid with the **SpreadsheetSet** parameter set.

  .EXAMPLE
    Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.txt"

    Exports the Employees table to a delimited text file.

  .EXAMPLE
    Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.txt" -HasFieldNames

    Exports the Employees table to a text file with field names as the first row.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.txt" -Password $password

    Exports data from a password-protected database.

  .EXAMPLE
    Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.html" -TransferType acExportHTML

    Exports the Employees table to an HTML file.

  .EXAMPLE
    Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.xlsx" -SpreadsheetType acSpreadsheetTypeExcel12Xml

    Exports the Employees table to an Excel spreadsheet.

  .EXAMPLE
    Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.xlsx" -SpreadsheetType acSpreadsheetTypeExcel12Xml -Range 'A1:C10'

    Exports the Employees table to an Excel spreadsheet with a specific range.

  .OUTPUTS
    System.IO.FileInfo
      Returns a FileInfo object representing the exported file.

  .NOTES
    The `-Password` and `-RemovePersonalInformation` parameters cannot be specified together in the underlying `New-AccessFile` cmdlet, but this cmdlet only reads from an existing database.
    The cmdlet creates a temporary file during text export to work around a `TransferText` limitation with file names containing multiple periods.
  #>
  [CmdletBinding(DefaultParameterSetName = 'TextSet', SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $TableName,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [Parameter(ParameterSetName = 'TextSet')]
    [Microsoft.Office.Interop.Access.AcTextTransferType]
    $TransferType = [Microsoft.Office.Interop.Access.AcTextTransferType]::acExportDelim,
    [Parameter(ParameterSetName = 'TextSet')]
    [int]
    $CodePage = 1200,
    [Parameter(ParameterSetName = 'SpreadsheetSet')]
    [Microsoft.Office.Interop.Access.AcSpreadSheetType]
    $SpreadsheetType = [Microsoft.Office.Interop.Access.AcSpreadSheetType]::acSpreadsheetTypeExcel12Xml,
    [Parameter(ParameterSetName = 'SpreadsheetSet')]
    [string]
    $Range
  )
  process {
    $resolvedPath = [Path]::GetFullPath($Path)
    $resolvedDestination = [Path]::GetFullPath($Destination)
    $exists = Test-Path -LiteralPath $resolvedDestination
    if ($exists -and $NoClobber) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolvedDestination))
    }
    if ($exists -and -not $Force) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolvedDestination))
    }
    $isReadOnly = $false
    if ($exists -and $Force) {
      $destinationItem = Get-Item -LiteralPath $resolvedDestination -Force
      $isReadOnly = $destinationItem.IsReadOnly
      if ($isReadOnly) {
        $destinationItem.IsReadOnly = $false
      }
    }
    $action = if ($PSCmdlet.ParameterSetName -eq 'SpreadsheetSet') {
      'Export Access data to spreadsheet'
    } else {
      'Export Access data to text'
    }
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($resolvedDestination, $action))) {
      return
    }
    if ($exists) {
      Remove-Item -LiteralPath $resolvedDestination -Force -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $app = New-AccessObject
    try {
      Open-AccessFile -Application $app -Path $resolvedPath -Password $Password
      try {
        switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
          'TextSet' {
            # Create a temporary file with a single period for TransferText
            $tempDestination = [Path]::GetTempFileName()
            try {
              $app.DoCmd.TransferText(
                $TransferType       # TransferType
                , [type]::Missing   # SpecificationName
                , $TableName        # TableName
                , $tempDestination  # FileName
                , $true             # HasFieldNames
                , [type]::Missing   # HTMLTableName
                , $CodePage         # CodePage
              )
              # Rename temp file to actual destination
              if (Test-Path -LiteralPath $tempDestination) {
                if (Test-Path -LiteralPath $resolvedDestination) {
                  Remove-Item -LiteralPath $resolvedDestination -Force -WhatIf:$WhatIfPreference -Confirm:$false
                }
                Rename-Item -LiteralPath $tempDestination -NewName $resolvedDestination -Force
                return Get-Item -LiteralPath $resolvedDestination -Force
              }
            } finally {
              if (Test-Path -LiteralPath $tempDestination) {
                Remove-Item -LiteralPath $tempDestination -Force -WhatIf:$false -Confirm:$false
              }
            }
          }
          'SpreadsheetSet' {
            $app.DoCmd.TransferSpreadsheet(
              [Microsoft.Office.Interop.Access.AcDataTransferType]::acExport  # TransferType
              , $SpreadsheetType                                              # SpreadsheetType
              , $TableName                                                    # TableName
              , $resolvedDestination                                          # FileName
              , $true                                                         # HasFieldNames
              , $(if ($Range) { $Range } else { [type]::Missing })            # Range
            )
            return Get-Item -LiteralPath $resolvedDestination -Force
          }
        }
      } finally {
        $app.CloseCurrentDatabase()
      }
    } finally {
      try {
        if ($app) {
          $app.Quit()
        }
      } finally {
        if ($exists -and $Force -and $isReadOnly -and (Test-Path -LiteralPath $resolvedDestination)) {
          (Get-Item -LiteralPath $resolvedDestination -Force).IsReadOnly = $true
        }
        Get-Variable |
        Where-Object -Property Value -Is [__ComObject] |
        Clear-Variable -Force -WhatIf:$false -Confirm:$false
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
      }
    }
  }
}
function Import-AccessDatabase {
  <#
  .SYNOPSIS
    Imports data from a text file or spreadsheet into an Access database table.

  .DESCRIPTION
    Imports data from a text file (CSV, delimited, HTML) or spreadsheet (Excel) into an Access database table by automating Access through COM.

    The cmdlet supports two parameter sets:
    - **TextSet** (default): Imports from text formats using `DoCmd.TransferText`. Supports delimited, fixed-width, and HTML import types.
    - **SpreadsheetSet**: Imports from Excel formats using `DoCmd.TransferSpreadsheet`.

    If the destination database is read-only, the cmdlet throws an error.

    Because this cmdlet supports `ShouldProcess`, you can use `-WhatIf` and `-Confirm` to preview or confirm the import operation.

  .PARAMETER Path
    Specifies the path to the Access database file to import into.

    This parameter does not support wildcards because it represents an existing file path.

  .PARAMETER Password
    Specifies the password required to open the database.

    Pass a `SecureString` value. If omitted, no password is used.

  .PARAMETER Source
    Specifies the path to the source file to import from.

    This parameter does not support wildcards because it represents an existing file path.

  .PARAMETER TableName
    Specifies the name of the Access table to import data into.

  .PARAMETER TransferType
    Specifies the text transfer type for the import.

    This parameter is only valid with the **TextSet** parameter set. The default value is `acImportDelim`.

  .PARAMETER CodePage
    Specifies the code page to use for the imported text file.

    This parameter is only valid with the **TextSet** parameter set. The default value is `1200` (Unicode).

  .PARAMETER HasFieldNames
    Indicates that the first row of the source file contains field names that should be used as column headers.

    This parameter is only valid with the **TextSet** parameter set.

  .PARAMETER SpreadsheetType
    Specifies the spreadsheet type for the import.

    This parameter is only valid with the **SpreadsheetSet** parameter set. The default value is `acSpreadsheetTypeExcel12Xml`.

  .PARAMETER Range
    Specifies the range of cells to import from the spreadsheet.

    This parameter is only valid with the **SpreadsheetSet** parameter set.

  .EXAMPLE
    Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.csv" -TableName 'Employees'

    Imports data from a CSV file into the Employees table.

  .EXAMPLE
    Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.csv" -TableName 'Employees' -HasFieldNames

    Imports data from a CSV file with field names in the first row.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.csv" -TableName 'Employees' -Password $password

    Imports data into a password-protected database.

  .EXAMPLE
    Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.xlsx" -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml

    Imports data from an Excel spreadsheet into the Employees table.

  .EXAMPLE
    Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.xlsx" -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml -Range 'A1:C10'

    Imports data from a specific range in an Excel spreadsheet.

  .OUTPUTS
    None.

  .NOTES
    The cmdlet creates a temporary file during text import to work around a `TransferText` limitation with file names containing multiple periods.
    If the destination database is read-only, the cmdlet throws an error.
  #>
  [CmdletBinding(DefaultParameterSetName = 'TextSet', SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Alias('FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Source,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $TableName,
    [Parameter(ParameterSetName = 'TextSet')]
    [Microsoft.Office.Interop.Access.AcTextTransferType]
    $TransferType = [Microsoft.Office.Interop.Access.AcTextTransferType]::acImportDelim,
    [Parameter(ParameterSetName = 'TextSet')]
    [int]
    $CodePage = 1200,
    [Parameter(ParameterSetName = 'SpreadsheetSet')]
    [Microsoft.Office.Interop.Access.AcSpreadSheetType]
    $SpreadsheetType = [Microsoft.Office.Interop.Access.AcSpreadSheetType]::acSpreadsheetTypeExcel12Xml,
    [Parameter(ParameterSetName = 'SpreadsheetSet')]
    [string]
    $Range
  )
  process {
    $resolvedPath = [Path]::GetFullPath($Path)
    $resolvedSource = [Path]::GetFullPath($Source)
    $action = if ($PSCmdlet.ParameterSetName -eq 'SpreadsheetSet') {
      'Import spreadsheet data into Access'
    } else {
      'Import text data into Access'
    }
    if (-not $PSCmdlet.ShouldProcess($resolvedPath, $action)) {
      return
    }
    $item = Get-Item -LiteralPath $resolvedPath -Force
    if ($item.IsReadOnly) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $item))
    }
    $app = New-AccessObject
    try {
      Open-AccessFile -Application $app -Path $resolvedPath -Password $Password
      try {
        switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
          'TextSet' {
            # Create a temporary file with a single period for TransferText
            $tempSource = [Path]::GetTempFileName()
            try {
              Copy-Item -LiteralPath $resolvedSource -Destination $tempSource -Force
              $app.DoCmd.TransferText(
                $TransferType                                               # TransferType
                , [type]::Missing                                           # SpecificationName
                , $TableName                                                # TableName
                , $tempSource                                               # FileName
                , $true                                                     # HasFieldNames
                , [type]::Missing                                           # HTMLTableName
                , $CodePage                                                 # CodePage
              )
            } finally {
              if (Test-Path -LiteralPath $tempSource) {
                Remove-Item -LiteralPath $tempSource -Force -WhatIf:$false -Confirm:$false
              }
            }
          }
          'SpreadsheetSet' {
            $app.DoCmd.TransferSpreadsheet(
              [Microsoft.Office.Interop.Access.AcDataTransferType]::acImport  # TransferType
              , $SpreadsheetType                                              # SpreadsheetType
              , $TableName                                                    # TableName
              , $resolvedSource                                               # FileName
              , $true                                                         # HasFieldNames
              , $(if ($Range) { $Range } else { [type]::Missing })            # Range
            )
          }
        }
      } finally {
        $app.CloseCurrentDatabase()
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
#endregion

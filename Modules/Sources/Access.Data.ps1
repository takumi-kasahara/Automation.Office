using namespace System.Data
using namespace System.Data.Odbc
using namespace System.Data.OleDb
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Net

Import-Module -Name ($PSScriptRoot | Join-Path -ChildPath 'Data.psm1')
Set-StrictMode -Version Latest

function Get-AccessTable {
  <#
  .SYNOPSIS
    Gets tables in Access database files.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB/ODBC and returns table entries.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Access database file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Access database file paths literally. Wildcards are not interpreted.

  .PARAMETER Password
    Specifies the password required to open a protected Access database.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
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
    [SecureString]
    $Password = $null
  )
  process {
    $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'PathSet' {
        Get-Item -Path $Path -Force
      }
      'LiteralPathSet' {
        Get-Item -LiteralPath $LiteralPath -Force
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName -Password $Password
      try {
        $schema = Get-DbTableSchema -Connection $connection
        foreach ($row in $schema) {
          $name = [string]$row['TABLE_NAME']
          $type = [string]$row['TABLE_TYPE']
          if ($type -eq 'SYSTEM TABLE') {
            continue
          }
          [PSCustomObject]@{
            Path = $item.FullName
            Name = $name
            Type = $type
          }
        }
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}
function Get-AccessTableColumn {
  <#
  .SYNOPSIS
    Gets table column metadata in Access database files.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB/ODBC and returns column metadata for specified tables.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Access database file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Access database file paths literally. Wildcards are not interpreted.

  .PARAMETER Table
    Specifies one or more table names whose column metadata is returned.

  .PARAMETER Password
    Specifies the password required to open a protected Access database.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
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
    [Parameter(Mandatory, Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Table,
    [SecureString]
    $Password = $null
  )
  process {
    $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'PathSet' {
        Get-Item -Path $Path -Force
      }
      'LiteralPathSet' {
        Get-Item -LiteralPath $LiteralPath -Force
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName -Password $Password
      try {
        $schema = Get-DbTableSchema -Connection $connection
        $tableNames = $schema | ForEach-Object { [string]$_.TABLE_NAME }
        foreach ($tableName in $Table) {
          if ($tableNames -notcontains $tableName) {
            throw [ArgumentException]::new("Table not found: $tableName", 'Table')
          }
          $columns = Get-DbColumn -Connection $connection -ObjectName $tableName
          foreach ($column in $columns) {
            [PSCustomObject]@{
              Path       = $item.FullName
              ObjectType = 'Table'
              ObjectName = $tableName
              ColumnName = $column.ColumnName
              Ordinal    = [int]$column.Ordinal
              DataType   = $column.DataType.FullName
            }
          }
        }
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}
function Get-AccessView {
  <#
  .SYNOPSIS
    Gets views in Access database files.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB and returns view entries.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Access database file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Access database file paths literally. Wildcards are not interpreted.

  .PARAMETER Password
    Specifies the password required to open a protected Access database.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
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
    [SecureString]
    $Password = $null
  )
  process {
    $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'PathSet' {
        Get-Item -Path $Path -Force
      }
      'LiteralPathSet' {
        Get-Item -LiteralPath $LiteralPath -Force
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName -Password $Password
      try {
        $schema = Get-DbViewSchema -Connection $connection
        foreach ($row in $schema) {
          $name = [string]$row['TABLE_NAME']
          [PSCustomObject]@{
            Path = $item.FullName
            Name = $name
            Type = 'VIEW'
          }
        }
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}
function Get-AccessViewColumn {
  <#
  .SYNOPSIS
    Gets view column metadata in Access database files.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB/ODBC and returns column metadata for specified views.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Access database file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Access database file paths literally. Wildcards are not interpreted.

  .PARAMETER View
    Specifies one or more view names whose column metadata is returned.

  .PARAMETER Password
    Specifies the password required to open a protected Access database.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
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
    [Parameter(Mandatory, Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $View,
    [SecureString]
    $Password = $null
  )
  process {
    $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'PathSet' {
        Get-Item -Path $Path -Force
      }
      'LiteralPathSet' {
        Get-Item -LiteralPath $LiteralPath -Force
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName -Password $Password
      try {
        $schema = Get-DbViewSchema -Connection $connection
        $viewNames = $schema | ForEach-Object { [string]$_.TABLE_NAME }
        foreach ($viewName in $View) {
          if ($viewNames -notcontains $viewName) {
            throw [ArgumentException]::new("View not found: $viewName", 'View')
          }
          $columns = Get-DbColumn -Connection $connection -ObjectName $viewName
          foreach ($column in $columns) {
            [PSCustomObject]@{
              Path       = $item.FullName
              ObjectType = 'View'
              ObjectName = $viewName
              ColumnName = $column.ColumnName
              Ordinal    = [int]$column.Ordinal
              DataType   = $column.DataType.FullName
            }
          }
        }
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}
function Invoke-AccessQuery {
  <#
  .SYNOPSIS
    Executes a SQL statement against Access database files and returns rows or affected-row metadata.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB and executes the specified SQL statement.

    For `SELECT` statements, rows are returned as PSCustomObject with Path, ObjectType, ObjectName, and column properties.

    For `INSERT`, `UPDATE`, and `DELETE` statements, a single PSCustomObject with Path, ObjectType, ObjectName, and RecordsAffected is returned.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Access database file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Access database file paths literally. Wildcards are not interpreted.

  .PARAMETER Table
    Specifies one or more table names to query.

  .PARAMETER View
    Specifies one or more view names to query.

  .PARAMETER Columns
    Specifies the columns to return when querying tables or views by name.

    When omitted, all columns are returned.

  .PARAMETER Query
    Specifies a SQL statement to execute against the database. `SELECT`, `INSERT`, `UPDATE`, and `DELETE` are supported.

  .PARAMETER Password
    Specifies the password required to open a protected Access database.

  .NOTES
    SQL syntax is based on the Microsoft Access SQL dialect. For a complete SQL reference, see:
    https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/microsoft-access-sql-reference

    For connection string details, see:
    https://www.connectionstrings.com/access/

  .OUTPUTS
    System.Management.Automation.PSCustomObject
  #>
  [CmdletBinding(DefaultParameterSetName = 'TablePathSet')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'TablePathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'ViewPathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'QueryPathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'TableLiteralPathSet', ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'ViewLiteralPathSet', ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'QueryLiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [Parameter(Mandatory, ParameterSetName = 'TablePathSet', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'TableLiteralPathSet', Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Table,
    [Parameter(Mandatory, ParameterSetName = 'ViewPathSet', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'ViewLiteralPathSet', Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $View,
    [Parameter(ParameterSetName = 'TablePathSet')]
    [Parameter(ParameterSetName = 'TableLiteralPathSet')]
    [Parameter(ParameterSetName = 'ViewPathSet')]
    [Parameter(ParameterSetName = 'ViewLiteralPathSet')]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Columns,
    [Parameter(Mandatory, ParameterSetName = 'QueryPathSet', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'QueryLiteralPathSet', Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Query,
    [SecureString]
    $Password = $null
  )
  process {
    $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      { $_ -in 'TablePathSet', 'ViewPathSet', 'QueryPathSet' } {
        Get-Item -Path $Path -Force
      }
      { $_ -in 'TableLiteralPathSet', 'ViewLiteralPathSet', 'QueryLiteralPathSet' } {
        Get-Item -LiteralPath $LiteralPath -Force
      }
    }
    $targets = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      { $_ -in 'TablePathSet', 'TableLiteralPathSet' } {
        $Table | ForEach-Object {
          [PSCustomObject]@{
            Name       = $_
            ObjectType = 'Table'
          }
        }
      }
      { $_ -in 'ViewPathSet', 'ViewLiteralPathSet' } {
        $View | ForEach-Object {
          [PSCustomObject]@{
            Name       = $_
            ObjectType = 'View'
          }
        }
      }
      { $_ -in 'QueryPathSet', 'QueryLiteralPathSet' } {
        [PSCustomObject]@{
          Name       = $Query.TrimStart()
          ObjectType = 'Query'
        }
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName -Password $Password
      try {
        foreach ($target in $targets) {
          $sql = if ($target.ObjectType -eq 'Query') {
            $target.Name
          } else {
            $escapedName = ConvertTo-SqlIdentifier -Name $target.Name
            $selectList = if ($Columns) {
              ($Columns | ForEach-Object { ConvertTo-SqlIdentifier -Name $_ }) -join ', '
            } else {
              '*'
            }
            "SELECT $selectList FROM $escapedName"
          }
          $queryOutput = Invoke-Query -Connection $connection -Query $sql
          $dataTable = $queryOutput.Table
          if (-not ($dataTable -is [DataTable])) {
            throw [InvalidOperationException]::new("Invalid query result type: $($dataTable.GetType().FullName)")
          }
          if ($target.ObjectType -eq 'Query' -and $queryOutput.PSObject.Properties.Name -contains 'RecordsAffected' -and $dataTable.Rows.Count -eq 0) {
            [PSCustomObject]@{
              Path            = $item.FullName
              ObjectType      = $target.ObjectType
              ObjectName      = $target.Name
              RecordsAffected = $queryOutput.RecordsAffected
            }
          } else {
            foreach ($rowData in $dataTable.Rows) {
              $result = [ordered]@{}
              $result.Path = $item.FullName
              $result.ObjectType = $target.ObjectType
              $result.ObjectName = $target.Name
              foreach ($column in $dataTable.Columns) {
                $value = $rowData[$column.ColumnName]
                $result[$column.ColumnName] = if ($value -is [DBNull]) { $null } else { $value }
              }
              [PSCustomObject]$result
            }
          }
        }
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}

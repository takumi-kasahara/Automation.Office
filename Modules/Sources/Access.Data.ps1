using namespace System.Collections.Generic
using namespace System.Data
using namespace System.Data.Odbc
using namespace System.Data.OleDb
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Net

Add-Type -AssemblyName System.Data
Set-StrictMode -Version Latest

#region Private
function Get-OdbcConnectionString {
  [CmdletBinding()]
  [OutputType([string[]])]
  param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null
  )
  $resolved = (Resolve-Path -LiteralPath $Path).Path

  $extension = [Path]::GetExtension($resolved)
  $drivers = Get-OdbcDriver | Where-Object -Property Name -Like "*[*]$extension*" | Select-Object -ExpandProperty Name
  return $drivers |
  ForEach-Object {
    # https://learn.microsoft.com/en-us/sql/odbc/microsoft/sqldriverconnect-access-driver?view=sql-server-ver17
    $driver = [string]$_
    $builder = [OdbcConnectionStringBuilder]::new()
    $builder.Driver = $driver
    $builder['Dbq'] = $resolved
    $builder['Uid'] = 'Admin'
    if ($Password) {
      $passwordString = [NetworkCredential]::new([string]::Empty, $Password).Password
      $builder['Pwd'] = $passwordString
    }
    return $builder.ConnectionString
  }
}
function Get-OleDbConnectionString {
  [CmdletBinding()]
  [OutputType([string[]])] # Returns an array of OLE DB connection strings
  param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null
  )
  $resolved = (Resolve-Path -LiteralPath $Path).Path
  $extension = [Path]::GetExtension($resolved)
  if ($extension -notin '.accdb', '.mdb') {
    throw [ArgumentException]::new("Unsupported extension: $extension. Supported extensions are .accdb and .mdb.", 'Path')
  }

  $availableProviders = [HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $elements = [OleDbEnumerator]::new().GetElements()
  foreach ($row in $elements.Rows) {
    if ($row['SOURCES_TYPE'] -eq 1) {
      [void]$availableProviders.Add([string]$row['SOURCES_NAME'])
    }
  }
  $preferredProviders = switch -Exact -CaseSensitive ($extension) {
    '.accdb' {
      @(
        'Microsoft.ACE.OLEDB.16.0'
        'Microsoft.ACE.OLEDB.15.0'
        'Microsoft.ACE.OLEDB.12.0'
      )
    }
    '.mdb' {
      @(
        'Microsoft.ACE.OLEDB.16.0'
        'Microsoft.ACE.OLEDB.15.0'
        'Microsoft.ACE.OLEDB.12.0'
        'Microsoft.Jet.OLEDB.4.0'
      )
    }
  }
  $providers = $preferredProviders | Where-Object { $availableProviders.Contains($_) }
  return $providers |
  ForEach-Object {
    # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/microsoft-ole-db-provider-for-microsoft-jet
    $provider = [string]$_
    $builder = [OleDbConnectionStringBuilder]::new()
    $builder.Provider = $provider
    $builder.DataSource = $resolved
    $builder['Persist Security Info'] = $false
    if ($Password) {
      $passwordString = [NetworkCredential]::new([string]::Empty, $Password).Password
      $builder['Jet OLEDB:Database Password'] = $passwordString
    }
    return $builder.ConnectionString
  }
}
function Open-DbConnection {
  [CmdletBinding()]
  [OutputType([System.Data.IDbConnection])]
  param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null
  )
  $messages = @()
  foreach ($connectionString in (Get-OleDbConnectionString -Path $Path -Password $Password)) {
    $connection = [OleDbConnection]::new([string]$connectionString)
    try {
      $connection.Open()
      return [IDbConnection]$connection
    } catch {
      $messages += $_.Exception.Message
      if ($connection) {
        $connection.Dispose()
      }
    }
  }
  foreach ($connectionString in (Get-OdbcConnectionString -Path $Path -Password $Password)) {
    $connection = [OdbcConnection]::new($connectionString)
    try {
      $connection.Open()
      return [IDbConnection]$connection
    } catch {
      $messages += $_.Exception.Message
      if ($connection) {
        $connection.Dispose()
      }
    }
  }
  throw [InvalidOperationException]::new("Failed to open Access file: $Path. $(($messages | Select-Object -Unique) -join ' | ')")
}
function ConvertTo-AccessIdentifier {
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Name
  )
  return "[$($Name.Replace(']', ']]'))]"
}
function ConvertTo-AccessDataObject {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory)]
    [System.Data.DataRow]
    $Row,
    [Parameter(Mandatory)]
    [string]
    $Path,
    [Parameter(Mandatory)]
    [string]
    $ObjectType,
    [Parameter(Mandatory)]
    [string]
    $ObjectName
  )
  $result = [ordered]@{
    Path       = $Path
    ObjectType = $ObjectType
    ObjectName = $ObjectName
  }
  foreach ($column in $Row.Table.Columns) {
    $value = $Row[$column.ColumnName]
    $result[$column.ColumnName] = if ($value -is [DBNull]) {
      $null
    } else {
      $value
    }
  }
  return [PSCustomObject]$result
}
function Get-AccessTableSchema {
  [CmdletBinding()]
  [OutputType([System.Data.DataTable])]
  param (
    [Parameter(Mandatory)]
    [IDbConnection]
    $Connection
  )
  if ($Connection -is [OleDbConnection]) {
    return [DataTable]$Connection.GetOleDbSchemaTable([OleDbSchemaGuid]::Tables, $null)
  }
  if ($Connection -is [OdbcConnection]) {
    return [DataTable]$Connection.GetSchema('Tables')
  }
  throw [InvalidOperationException]::new("Unsupported connection type: $($Connection.GetType().FullName)")
}
function Get-AccessViewSchema {
  [CmdletBinding()]
  [OutputType([System.Data.DataTable])]
  param (
    [Parameter(Mandatory)]
    [IDbConnection]
    $Connection
  )
  if ($Connection -is [OleDbConnection]) {
    return [DataTable]$Connection.GetOleDbSchemaTable([OleDbSchemaGuid]::Views, $null)
  }
  if ($Connection -is [OdbcConnection]) {
    return [DataTable]$Connection.GetSchema('Views')
  }
  throw [InvalidOperationException]::new("Unsupported connection type: $($Connection.GetType().FullName)")
}
function Invoke-Query {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory)]
    [IDbConnection]
    $Connection,
    [Parameter(Mandatory)]
    [string]
    $Query
  )
  $dataTable = [DataTable]::new()
  if ($Connection -is [OleDbConnection]) {
    $adapter = [OleDbDataAdapter]::new($Query, $Connection)
    try {
      [void]$adapter.Fill($dataTable)
    } finally {
      if ($adapter) {
        $adapter.Dispose()
      }
    }
    return [PSCustomObject]@{ Table = $dataTable }
  }
  if ($Connection -is [OdbcConnection]) {
    $adapter = [OdbcDataAdapter]::new($Query, $Connection)
    try {
      [void]$adapter.Fill($dataTable)
    } finally {
      if ($adapter) {
        $adapter.Dispose()
      }
    }
    return [PSCustomObject]@{ Table = $dataTable }
  }
  throw [InvalidOperationException]::new("Unsupported connection type: $($Connection.GetType().FullName)")
}
function Test-AccessSelectQuery {
  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Query
  )
  return $Query.TrimStart() -match '^(?i)SELECT\b'
}
function Get-AccessDataColumn {
  [CmdletBinding()]
  [OutputType([System.Data.DataColumn[]])]
  param (
    [Parameter(Mandatory)]
    [IDbConnection]
    $Connection,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $ObjectName
  )
  $escapedName = ConvertTo-AccessIdentifier -Name $ObjectName
  $sql = "SELECT * FROM $escapedName WHERE 1 = 0"
  $queryOutput = Invoke-Query -Connection $Connection -Query $sql
  $dataTable = $queryOutput.Table
  if (-not ($dataTable -is [DataTable])) {
    throw [InvalidOperationException]::new("Invalid query result type: $($dataTable.GetType().FullName)")
  }
  return $dataTable.Columns
}
#endregion
#region Public
function Get-AccessTable {
  <#
  .SYNOPSIS
    Gets user tables in Access database files.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB and returns table entries.

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
        $schema = Get-AccessTableSchema -Connection $connection
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
        $schema = Get-AccessViewSchema -Connection $connection
        foreach ($row in $schema) {
          $name = [string]$row['TABLE_NAME']
          $type = [string]$row['TABLE_TYPE']
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
        foreach ($tableName in $Table) {
          $columns = Get-AccessDataColumn -Connection $connection -ObjectName $tableName
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
        foreach ($viewName in $View) {
          $columns = Get-AccessDataColumn -Connection $connection -ObjectName $viewName
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
function Get-AccessData {
  <#
  .SYNOPSIS
    Gets rows from Access tables or views.

  .DESCRIPTION
    Opens one or more Access database files (.accdb, .mdb) through OLE DB and returns rows from the specified table or view.

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
    Specifies a SELECT statement to execute against the database.

    Only SELECT statements are allowed. Data modification statements such as INSERT, UPDATE,
    DELETE, and DDL statements are rejected.

  .PARAMETER Password
    Specifies the password required to open a protected Access database.

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
        if (-not (Test-AccessSelectQuery -Query $Query)) {
          throw [ArgumentException]::new('Only SELECT statements are allowed for -Query. Use a SELECT statement that returns rows.', 'Query')
        }
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
            $escapedName = ConvertTo-AccessIdentifier -Name $target.Name
            $selectList = if ($Columns) {
              ($Columns | ForEach-Object { ConvertTo-AccessIdentifier -Name $_ }) -join ', '
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
          foreach ($row in @($dataTable.Rows)) {
            ConvertTo-AccessDataObject -Row $row -Path $item.FullName -ObjectType $target.ObjectType -ObjectName $target.Name
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
#endregion

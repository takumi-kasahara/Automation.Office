using namespace System.Data
using namespace System.Data.Odbc
using namespace System.Data.OleDb
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO

Add-Type -AssemblyName System.Data
Set-StrictMode -Version Latest

#region Private
function Get-ExcelTableSchema {
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
#endregion
#region Public
function Get-ExcelTable {
  <#
  .SYNOPSIS
    Gets tables in Excel workbook files.

  .DESCRIPTION
    Opens one or more Excel workbook files (.xlsx, .xls, .xlsm, .xlsb) through OLE DB/ODBC and returns table entries.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Excel workbook file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Excel workbook file paths literally. Wildcards are not interpreted.

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
        $schema = Get-ExcelTableSchema -Connection $connection
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
function Get-ExcelTableColumn {
  <#
  .SYNOPSIS
    Gets table column metadata in Excel workbook files.

  .DESCRIPTION
    Opens one or more Excel workbook files (.xlsx, .xls, .xlsm, .xlsb) through OLE DB/ODBC and returns column metadata for specified tables.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Excel workbook file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Excel workbook file paths literally. Wildcards are not interpreted.

  .PARAMETER Table
    Specifies one or more table names whose column metadata is returned.

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
        $schema = Get-ExcelTableSchema -Connection $connection
        foreach ($row in $schema) {
          $tableName = $row['TABLE_NAME']
          if ($Table -contains $tableName) {
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
        }
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}
function Get-ExcelData {
  <#
  .SYNOPSIS
    Gets rows from Excel tables.

  .DESCRIPTION
    Opens one or more Excel workbook files (.xlsx, .xls, .xlsm, .xlsb) through OLE DB/ODBC and returns rows from the specified table.

    This cmdlet does not use COM objects.

  .PARAMETER Path
    Specifies Excel workbook file paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies Excel workbook file paths literally. Wildcards are not interpreted.

  .PARAMETER Table
    Specifies one or more table names to query.

  .PARAMETER Columns
    Specifies the columns to return when querying tables by name.

    When omitted, all columns are returned.

  .PARAMETER Query
    Specifies a SELECT statement to execute against the database.

  .OUTPUTS
    System.Management.Automation.PSCustomObject
  #>
  [CmdletBinding(DefaultParameterSetName = 'TablePathSet')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'TablePathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'QueryPathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'TableLiteralPathSet', ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'QueryLiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [Parameter(Mandatory, ParameterSetName = 'TablePathSet', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'TableLiteralPathSet', Position = 1)]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Table,
    [Parameter(ParameterSetName = 'TablePathSet')]
    [Parameter(ParameterSetName = 'TableLiteralPathSet')]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Columns,
    [Parameter(Mandatory, ParameterSetName = 'QueryPathSet', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'QueryLiteralPathSet', Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Query
  )
  process {
    $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      { $_ -in 'TablePathSet', 'QueryPathSet' } {
        Get-Item -Path $Path -Force
      }
      { $_ -in 'TableLiteralPathSet', 'QueryLiteralPathSet' } {
        Get-Item -LiteralPath $LiteralPath -Force
      }
    }
    $targets = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      { $_ -in 'TablePathSet', 'TableLiteralPathSet' } {
        $Table | ForEach-Object {
          [PSCustomObject]@{ Name = $_ }
        }
      }
      { $_ -in 'QueryPathSet', 'QueryLiteralPathSet' } {
        [PSCustomObject]@{ Name = $Query.TrimStart() }
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName
      try {
        foreach ($target in $targets) {
          $sql = if ($PSCmdlet.ParameterSetName -like 'Query*') {
            $target.Name
          } else {
            $selectList = if ($Columns) {
              ($Columns | ForEach-Object { "[$_]" }) -join ', '
            } else {
              '*'
            }
            "SELECT $selectList FROM [$($target.Name)]"
          }
          $queryOutput = Invoke-Query -Connection $connection -Query $sql
          $dataTable = $queryOutput.Table
          if (-not ($dataTable -is [DataTable])) {
            throw [InvalidOperationException]::new("Invalid query result type: $($dataTable.GetType().FullName)")
          }
          foreach ($rowData in $dataTable.Rows) {
            $result = [ordered]@{}
            $result.Path = $item.FullName
            $result.TableName = $target.Name
            foreach ($column in $dataTable.Columns) {
              $value = $rowData[$column.ColumnName]
              $result[$column.ColumnName] = if ($value -is [DBNull]) { $null } else { $value }
            }
            [PSCustomObject]$result
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

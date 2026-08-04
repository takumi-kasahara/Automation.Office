using namespace System.Data
using namespace System.Data.Odbc
using namespace System.Data.OleDb
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO

Add-Type -AssemblyName System.Data
Set-StrictMode -Version Latest

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
        $schema = Get-DbTableSchema -Connection $connection
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

  .PARAMETER Address
    Specifies one or more cell ranges to query within the corresponding table.

    Use A1-style notation such as A1:B10. The range is appended to the table name
    in the form [Sheet1$A1:B10]. This parameter must have the same number of
    elements as -Table.

    Reference: [Import from Excel or Export to Excel with SQL Server Integration Services (SSIS)](https://learn.microsoft.com/en-us/sql/integration-services/load-data-to-from-excel-with-ssis?view=sql-server-ver17)

  .PARAMETER Columns
    Specifies the columns to return when querying tables by name.

    When omitted, all columns are returned.

  .PARAMETER Query
    Specifies a SELECT statement to execute against the database.

    Only SELECT statements are allowed. Data modification statements such as INSERT, UPDATE, DELETE, and DDL statements are rejected.

  .PARAMETER NoHeader
    Specifies that the first row of the Excel range does not contain column names.

    When this switch is specified, the OLE DB connection uses `HDR=NO` instead of
    the default `HDR=YES`. Column names appear as F1, F2, F3, and so on.

    Reference: [Initializing the Microsoft Excel driver](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/initializing-the-microsoft-excel-driver)

  .PARAMETER NoIMEX
    Disables IMEX mode for the OLE DB connection.

    When this switch is specified, the OLE DB connection uses `IMEX=0` instead of
    the default `IMEX=1`. With IMEX disabled, the driver may return null for cells
    whose data type does not match the guessed column type.

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
    $Address,
    [Parameter(ParameterSetName = 'TablePathSet')]
    [Parameter(ParameterSetName = 'TableLiteralPathSet')]
    [ValidateCount(1, [int]::MaxValue)]
    [string[]]
    $Columns,
    [Parameter(Mandatory, ParameterSetName = 'QueryPathSet', Position = 1)]
    [Parameter(Mandatory, ParameterSetName = 'QueryLiteralPathSet', Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Query,
    [Parameter(ParameterSetName = 'TablePathSet')]
    [Parameter(ParameterSetName = 'TableLiteralPathSet')]
    [Parameter(ParameterSetName = 'QueryPathSet')]
    [Parameter(ParameterSetName = 'QueryLiteralPathSet')]
    [switch]
    $NoHeader,
    [Parameter(ParameterSetName = 'TablePathSet')]
    [Parameter(ParameterSetName = 'TableLiteralPathSet')]
    [Parameter(ParameterSetName = 'QueryPathSet')]
    [Parameter(ParameterSetName = 'QueryLiteralPathSet')]
    [switch]
    $NoIMEX
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
        if ($Address -and $Address.Count -ne $Table.Count) {
          throw [ArgumentException]::new('The number of elements in -Address must match the number of elements in -Table.', 'Address')
        }
        for ($i = 0; $i -lt $Table.Count; $i++) {
          $name = $Table[$i]
          if ($Address) {
            $name = "$name$($Address[$i])"
          }
          [PSCustomObject]@{
            Name       = $name
            ObjectType = 'Table'
          }
        }
      }
      { $_ -in 'QueryPathSet', 'QueryLiteralPathSet' } {
        if (-not (Test-SelectQuery -Query $Query)) {
          throw [ArgumentException]::new('Only SELECT statements are allowed for -Query. Use a SELECT statement that returns rows.', 'Query')
        }
        [PSCustomObject]@{
          Name       = $Query.TrimStart()
          ObjectType = 'Query'
        }
      }
    }
    foreach ($item in $items) {
      $connection = Open-DbConnection -Path $item.FullName -NoHeader:$NoHeader -NoIMEX:$NoIMEX
      try {
        foreach ($target in $targets) {
          $sql = if ($target.ObjectType -eq 'Query') {
            $target.Name
          } else {
            $escapedName = if ($target.Name -match '\$[A-Za-z]+\d+(:[A-Za-z]+\d+)?$') {
              $target.Name -replace '([\]\\])', '$1$1' -replace '^(.+)$', '[$1]'
            } else {
              ConvertTo-SqlIdentifier -Name $target.Name
            }
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
      } finally {
        if ($connection) {
          $connection.Dispose()
        }
      }
    }
  }
}

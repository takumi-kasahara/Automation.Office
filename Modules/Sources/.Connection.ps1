using namespace System.Collections.Generic
using namespace System.Data
using namespace System.Data.Odbc
using namespace System.Data.OleDb
using namespace System.IO
using namespace System.Net

Add-Type -AssemblyName System.Data
Set-StrictMode -Version Latest

function ConvertTo-SqlIdentifier {
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
function Get-OdbcConnectionString {
  [CmdletBinding()]
  [OutputType([string[]])]
  param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $Password = $null,
    [switch]
    $ReadOnly
  )
  $resolved = (Resolve-Path -LiteralPath $Path).Path

  $extension = [Path]::GetExtension($resolved)
  $drivers = Get-OdbcDriver | Where-Object -Property Name -Like "*[*]$extension*" | Select-Object -ExpandProperty Name
  return $drivers |
  ForEach-Object {
    $driver = [string]$_
    $builder = [OdbcConnectionStringBuilder]::new()
    $builder.Driver = $driver
    $builder['Dbq'] = $resolved
    if ($extension -in '.accdb', '.mdb') {
      # https://learn.microsoft.com/en-us/sql/odbc/microsoft/sqldriverconnect-access-driver?view=sql-server-ver17
      if ($Password) {
        $passwordString = [NetworkCredential]::new([string]::Empty, $Password).Password
        $builder['Pwd'] = $passwordString
      }
    } elseif ($extension -in '.xlsx', '.xls', '.xlsm', '.xlsb') {
      if ($ReadOnly) {
        $builder['ReadOnly'] = 1
      }
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
    $Password = $null,
    [switch]
    $NoHeader,
    [switch]
    $NoIMEX
  )
  $resolved = (Resolve-Path -LiteralPath $Path).Path
  $extension = [Path]::GetExtension($resolved)
  if ($extension -notin '.accdb', '.mdb', '.xlsx', '.xls', '.xlsm', '.xlsb') {
    throw [ArgumentException]::new("Unsupported extension: $extension. Supported extensions are .accdb, .mdb, .xlsx, .xls, .xlsm, .xlsb.", 'Path')
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
    '.xlsx' {
      @(
        'Microsoft.ACE.OLEDB.16.0'
        'Microsoft.ACE.OLEDB.15.0'
        'Microsoft.ACE.OLEDB.12.0'
      )
    }
    '.xls' {
      @(
        'Microsoft.ACE.OLEDB.16.0'
        'Microsoft.ACE.OLEDB.15.0'
        'Microsoft.ACE.OLEDB.12.0'
        'Microsoft.Jet.OLEDB.4.0'
      )
    }
    '.xlsm' {
      @(
        'Microsoft.ACE.OLEDB.16.0'
        'Microsoft.ACE.OLEDB.15.0'
        'Microsoft.ACE.OLEDB.12.0'
      )
    }
    '.xlsb' {
      @(
        'Microsoft.ACE.OLEDB.16.0'
        'Microsoft.ACE.OLEDB.15.0'
        'Microsoft.ACE.OLEDB.12.0'
      )
    }
  }
  $providers = $preferredProviders | Where-Object { $availableProviders.Contains($_) }
  return $providers |
  ForEach-Object {
    $provider = [string]$_
    $builder = [OleDbConnectionStringBuilder]::new()
    $builder.Provider = $provider
    $builder['Data Source'] = $resolved
    if ($extension -in '.accdb', '.mdb') {
      # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/microsoft-ole-db-provider-for-microsoft-jet
      $builder['Persist Security Info'] = $false
      if ($Password) {
        $passwordString = [NetworkCredential]::new([string]::Empty, $Password).Password
        $builder['Jet OLEDB:Database Password'] = $passwordString
      }
    } elseif ($extension -in '.xlsx', '.xls', '.xlsm', '.xlsb') {
      # https://learn.microsoft.com/en-us/sql/integration-services/connection-manager/excel-connection-manager?view=sql-server-ver17
      $HDR = if ($NoHeader) { 'NO' } else { 'YES' }
      $IMEX = if ($NoIMEX) { '0' } else { '1' }
      $builder['Extended Properties'] = switch -Exact -CaseSensitive ($extension) {
        '.xlsx' { "Excel 12.0 Xml;HDR=$HDR;IMEX=$IMEX" }
        '.xlsm' { "Excel 12.0 Macro;HDR=$HDR;IMEX=$IMEX" }
        '.xlsb' { "Excel 12.0 Binary;HDR=$HDR;IMEX=$IMEX" }
        '.xls' { "Excel 8.0;HDR=$HDR;IMEX=$IMEX" }
      }
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
    $Password = $null,
    [switch]
    $NoHeader,
    [switch]
    $NoIMEX,
    [switch]
    $ReadOnly
  )
  $messages = @()
  foreach ($connectionString in (Get-OleDbConnectionString -Path $Path -Password $Password -NoHeader:$NoHeader -NoIMEX:$NoIMEX)) {
    $connection = [OleDbConnection]::new([string]$connectionString)
    try {
      $connection.Open()
      $PSCmdlet.WriteVerbose("Opened OLE DB connection using provider: $($connection.Provider)")
      return [IDbConnection]$connection
    } catch {
      $messages += $_.Exception.Message
      if ($connection) {
        $connection.Dispose()
      }
    }
  }
  foreach ($connectionString in (Get-OdbcConnectionString -Path $Path -Password $Password -ReadOnly:$ReadOnly)) {
    $connection = [OdbcConnection]::new($connectionString)
    try {
      $connection.Open()
      $PSCmdlet.WriteVerbose("Opened ODBC connection using driver: $($connection.Driver)")
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
function Get-DbTableSchema {
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
function Get-DbViewSchema {
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
function Get-DbColumn {
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
  $escapedName = ConvertTo-SqlIdentifier -Name $ObjectName
  $sql = "SELECT * FROM $escapedName WHERE 1 = 0"
  $queryOutput = Invoke-Query -Connection $Connection -Query $sql
  $dataTable = $queryOutput.Table
  if (-not ($dataTable -is [DataTable])) {
    throw [InvalidOperationException]::new("Invalid query result type: $($dataTable.GetType().FullName)")
  }
  return $dataTable.Columns
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
  $normalizedQuery = $Query.TrimStart()
  $isDml = $normalizedQuery -imatch '^\s*(INSERT|UPDATE|DELETE)\b'
  if ($isDml) {
    $command = $Connection.CreateCommand()
    try {
      $command.CommandText = $Query
      $recordsAffected = $command.ExecuteNonQuery()
      return [PSCustomObject]@{ Table = [DataTable]::new(); RecordsAffected = $recordsAffected }
    } finally {
      if ($command) {
        $command.Dispose()
      }
    }
  }
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
    return [PSCustomObject]@{ Table = $dataTable; RecordsAffected = $dataTable.Rows.Count }
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
    return [PSCustomObject]@{ Table = $dataTable; RecordsAffected = $dataTable.Rows.Count }
  }
  throw [InvalidOperationException]::new("Unsupported connection type: $($Connection.GetType().FullName)")
}

using namespace System.Collections.Generic
using namespace System.Data
using namespace System.Data.Odbc
using namespace System.Data.OleDb
using namespace System.IO
using namespace System.Net

Add-Type -AssemblyName System.Data
Set-StrictMode -Version Latest

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
    $driver = [string]$_
    $builder = [OdbcConnectionStringBuilder]::new()
    $builder.Driver = $driver
    if ($extension -in '.accdb', '.mdb') {
      # https://learn.microsoft.com/en-us/sql/odbc/microsoft/sqldriverconnect-access-driver?view=sql-server-ver17
      $builder['DBQ'] = $resolved
      $builder['UID'] = 'Admin'
      if ($Password) {
        $passwordString = [NetworkCredential]::new([string]::Empty, $Password).Password
        $builder['PWD'] = $passwordString
      }
    } elseif ($extension -in '.xlsx', '.xls', '.xlsm', '.xlsb') {
      $builder['DBQ'] = $resolved
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
      # https://learn.microsoft.com/en-us/power-automate/desktop-flows/how-to/sql-queries-excel
      # https://learn.microsoft.com/en-us/sql/integration-services/connection-manager/excel-connection-manager?view=sql-server-ver17
      $builder['Extended Properties'] = switch -Exact -CaseSensitive ($extension) {
        '.xlsx' { 'Excel 12.0 Xml;HDR=YES;IMEX=1' }
        '.xlsm' { 'Excel 12.0 Macro;HDR=YES;IMEX=1' }
        '.xlsb' { 'Excel 12.0 Binary;HDR=YES;IMEX=1' }
        '.xls' { 'Excel 8.0;HDR=YES;IMEX=1' }
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

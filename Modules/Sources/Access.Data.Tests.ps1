using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest

InModuleScope 'Automation.Office' {
  BeforeAll {
    function Get-Password {
      [CmdletBinding()]
      [OutputType([SecureString])]
      [SuppressMessage('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Used in tests to generate random passwords for verification purposes')]
      param ()
      return ConvertTo-SecureString -String ([guid]::NewGuid().ToString('N').SubString(0, 20)) -AsPlainText -Force
    }
    function Get-TempFile {
      [CmdletBinding()]
      [OutputType([string])]
      param (
        [string]
        $Extension = '.accdb'
      )
      return $env:TEMP | Join-Path -ChildPath "Database.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function Initialize-AccessFixture {
      param(
        [string]
        $Path,
        [SecureString]
        $Password = $null
      )
      New-AccessFile -Path $Path -Password $Password -Force -InitializeDb {
        param($db)
        $db.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255), Department TEXT(255))')
        $db.Execute("INSERT INTO Employees (Id, Name, Department) VALUES (1, 'Alice', 'Sales')")
        $db.Execute("INSERT INTO Employees (Id, Name, Department) VALUES (2, 'Bob', 'Engineering')")
        $db.CreateQueryDef('vwSalesEmployees', "SELECT Id, Name FROM Employees WHERE Department = 'Sales'") | Out-Null
      } | Out-Null
    }
  }
  Describe 'Get-AccessTable' {
    BeforeEach {
      $accdbPath = Get-TempFile -Extension '.accdb'
      $mdbPath = Get-TempFile -Extension '.mdb'
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $accdbPath) {
        Remove-Item -LiteralPath $accdbPath -Force
      }
      if (Test-Path -LiteralPath $mdbPath) {
        Remove-Item -LiteralPath $mdbPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets tables by Path' {
        Initialize-AccessFixture -Path $accdbPath
        $tables = Get-AccessTable -Path $accdbPath
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by LiteralPath' {
        Initialize-AccessFixture -Path $accdbPath
        $tables = Get-AccessTable -LiteralPath $accdbPath
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipeline' {
        Initialize-AccessFixture -Path $accdbPath
        $tables = $accdbPath | Get-AccessTable
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $accdbPath
        $tables = [PSCustomObject]@{ PSPath = $accdbPath } | Get-AccessTable
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-AccessFixture -Path $accdbPath
        $table = Get-AccessTable -LiteralPath $accdbPath | Where-Object -Property Name -EQ 'Employees' | Select-Object -First 1
        $table | Should -Not -BeNullOrEmpty
        $table.Path | Should -Be $accdbPath
        $table.Type | Should -Be 'TABLE'
        $table.Name | Should -Be 'Employees'
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        Initialize-AccessFixture -Path $mdbPath
        $tables = Get-AccessTable -LiteralPath $mdbPath
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $tables = Get-AccessTable -LiteralPath $accdbPath -Password $password
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'throws when password is incorrect' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        { Get-AccessTable -LiteralPath $accdbPath -Password (Get-Password) } | Should -Throw
      }
      It 'throws when password is not provided for protected database' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        { Get-AccessTable -LiteralPath $accdbPath } | Should -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $txtPath = Get-TempFile -Extension '.txt'
        New-Item -Path $txtPath -ItemType File -Force | Out-Null
        { Get-AccessTable -LiteralPath $txtPath } | Should -Throw
      }
    }
  }
  Describe 'Get-AccessTableColumn' {
    BeforeEach {
      $accdbPath = Get-TempFile -Extension '.accdb'
      $mdbPath = Get-TempFile -Extension '.mdb'
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $accdbPath) {
        Remove-Item -LiteralPath $accdbPath -Force
      }
      if (Test-Path -LiteralPath $mdbPath) {
        Remove-Item -LiteralPath $mdbPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table columns by Path' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = Get-AccessTableColumn -LiteralPath $accdbPath -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by LiteralPath' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = Get-AccessTableColumn -LiteralPath $accdbPath -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipeline' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = $accdbPath | Get-AccessTableColumn -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = [PSCustomObject]@{ PSPath = $accdbPath } | Get-AccessTableColumn -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = Get-AccessTableColumn -LiteralPath $accdbPath -Table Employees
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should -Not -BeNullOrEmpty
        $name | Should -Not -BeNullOrEmpty
        $id.Path | Should -Be $accdbPath
        $id.ObjectType | Should -Be 'Table'
        $id.ObjectName | Should -Be 'Employees'
        $id.Ordinal | Should -BeOfType [int]
        $name.DataType | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        Initialize-AccessFixture -Path $mdbPath
        $columns = Get-AccessTableColumn -LiteralPath $mdbPath -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $columns = Get-AccessTableColumn -LiteralPath $accdbPath -Table Employees -Password $password
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'throws when password is incorrect' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        { Get-AccessTableColumn -LiteralPath $accdbPath -Table Employees } | Should -Throw
        { Get-AccessTableColumn -LiteralPath $accdbPath -Table Employees -Password (Get-Password) } | Should -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        Initialize-AccessFixture -Path $accdbPath
        { Get-AccessTableColumn -LiteralPath $accdbPath -Table NotExists } | Should -Throw
      }
    }
  }
  Describe 'Get-AccessView' {
    BeforeEach {
      $accdbPath = Get-TempFile -Extension '.accdb'
      $mdbPath = Get-TempFile -Extension '.mdb'
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $accdbPath) {
        Remove-Item -LiteralPath $accdbPath -Force
      }
      if (Test-Path -LiteralPath $mdbPath) {
        Remove-Item -LiteralPath $mdbPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets views by Path' {
        Initialize-AccessFixture -Path $accdbPath
        $views = Get-AccessView -Path $accdbPath
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
      It 'gets views by LiteralPath' {
        Initialize-AccessFixture -Path $accdbPath
        $views = Get-AccessView -LiteralPath $accdbPath
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
      It 'gets views by ValueFromPipeline' {
        Initialize-AccessFixture -Path $accdbPath
        $views = $accdbPath | Get-AccessView
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
      It 'gets views by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $accdbPath
        $views = [PSCustomObject]@{ PSPath = $accdbPath } | Get-AccessView
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-AccessFixture -Path $accdbPath
        $view = Get-AccessView -LiteralPath $accdbPath | Where-Object -Property Name -EQ 'vwSalesEmployees' | Select-Object -First 1
        $view | Should -Not -BeNullOrEmpty
        $view.Path | Should -Be $accdbPath
        $view.Type | Should -Match 'VIEW'
        $view.Name | Should -Be 'vwSalesEmployees'
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        Initialize-AccessFixture -Path $mdbPath
        $views = Get-AccessView -LiteralPath $mdbPath
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
      It 'gets views from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $views = Get-AccessView -LiteralPath $accdbPath -Password $password
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
      It 'throws when password is incorrect' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        { Get-AccessView -LiteralPath $accdbPath } | Should -Throw
        { Get-AccessView -LiteralPath $accdbPath -Password (Get-Password) } | Should -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $txtPath = Get-TempFile -Extension '.txt'
        New-Item -Path $txtPath -ItemType File -Force | Out-Null
        { Get-AccessView -LiteralPath $txtPath } | Should -Throw
      }
    }
  }
  Describe 'Get-AccessViewColumn' {
    BeforeEach {
      $accdbPath = Get-TempFile -Extension '.accdb'
      $mdbPath = Get-TempFile -Extension '.mdb'
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $accdbPath) {
        Remove-Item -LiteralPath $accdbPath -Force
      }
      if (Test-Path -LiteralPath $mdbPath) {
        Remove-Item -LiteralPath $mdbPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets view columns by Path' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = Get-AccessViewColumn -Path $accdbPath -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets view columns by LiteralPath' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = Get-AccessViewColumn -LiteralPath $accdbPath -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets view columns by ValueFromPipeline' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = $accdbPath | Get-AccessViewColumn -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets view columns by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = [PSCustomObject]@{ PSPath = $accdbPath } | Get-AccessViewColumn -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-AccessFixture -Path $accdbPath
        $columns = Get-AccessViewColumn -LiteralPath $accdbPath -View vwSalesEmployees
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should -Not -BeNullOrEmpty
        $name | Should -Not -BeNullOrEmpty
        $id.Path | Should -Be $accdbPath
        $id.ObjectType | Should -Be 'View'
        $id.ObjectName | Should -Be 'vwSalesEmployees'
        $id.Ordinal | Should -BeOfType [int]
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        Initialize-AccessFixture -Path $mdbPath
        $columns = Get-AccessViewColumn -LiteralPath $mdbPath -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets view columns from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $columns = Get-AccessViewColumn -LiteralPath $accdbPath -View vwSalesEmployees -Password $password
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'throws when password is incorrect' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        { Get-AccessViewColumn -LiteralPath $accdbPath -View vwSalesEmployees } | Should -Throw
        { Get-AccessViewColumn -LiteralPath $accdbPath -View vwSalesEmployees -Password (Get-Password) } | Should -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for missing view' {
        Initialize-AccessFixture -Path $accdbPath
        { Get-AccessViewColumn -LiteralPath $accdbPath -View NotExists } | Should -Throw
      }
    }
  }
  Describe 'Invoke-AccessQuery' {
    BeforeEach {
      $accdbPath = Get-TempFile -Extension '.accdb'
      $mdbPath = Get-TempFile -Extension '.mdb'
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $accdbPath) {
        Remove-Item -LiteralPath $accdbPath -Force
      }
      if (Test-Path -LiteralPath $mdbPath) {
        Remove-Item -LiteralPath $mdbPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table data by Path' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = Invoke-AccessQuery -Path $accdbPath -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets table data by LiteralPath' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets table data by ValueFromPipeline' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = $accdbPath | Invoke-AccessQuery -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets table data by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = [PSCustomObject]@{ PSPath = $accdbPath } | Invoke-AccessQuery -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets data with selected columns' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Table Employees -Columns Id, Name
        $rows | Should -HaveCount 2
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Id'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Name'
        @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'Department'
      }
      It 'gets view data with selected columns' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -View vwSalesEmployees -Columns Id, Name
        $rows | Should -HaveCount 1
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Id'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Name'
      }
      It 'gets data by Query' {
        Initialize-AccessFixture -Path $accdbPath
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'SELECT [Id], [Name] FROM Employees WHERE [Id] = 1'
        $rows | Should -HaveCount 1
        $rows[0].Id | Should -Be 1
        $rows[0].Name | Should -Be 'Alice'
      }
    }
    Context 'Input' {
      It 'inserts rows and returns affected count' {
        Initialize-AccessFixture -Path $accdbPath
        $result = Invoke-AccessQuery -LiteralPath $accdbPath -Query "INSERT INTO [Employees] (Id, Name, Department) VALUES (3, 'Carol', 'Marketing')"
        $result | Should -Not -BeNullOrEmpty
        $result.Path | Should -Be $accdbPath
        $result.ObjectType | Should -Be 'Query'
        $result.ObjectName | Should -Be "INSERT INTO [Employees] (Id, Name, Department) VALUES (3, 'Carol', 'Marketing')"
        $result.RecordsAffected | Should -Be 1
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'SELECT * FROM [Employees] WHERE [Id] = 3'
        $rows | Should -HaveCount 1
        $rows[0].Name | Should -Be 'Carol'
        $rows[0].Department | Should -Be 'Marketing'
      }
      It 'updates rows and returns affected count' {
        Initialize-AccessFixture -Path $accdbPath
        $result = Invoke-AccessQuery -LiteralPath $accdbPath -Query "UPDATE [Employees] SET [Department] = 'Marketing' WHERE [Id] = 1"
        $result | Should -Not -BeNullOrEmpty
        $result.RecordsAffected | Should -Be 1
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'SELECT * FROM [Employees] WHERE [Id] = 1'
        $rows | Should -HaveCount 1
        $rows[0].Department | Should -Be 'Marketing'
      }
      It 'deletes rows and returns affected count' {
        Initialize-AccessFixture -Path $accdbPath
        $result = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'DELETE FROM [Employees] WHERE [Id] = 2'
        $result | Should -Not -BeNullOrEmpty
        $result.RecordsAffected | Should -Be 1
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'SELECT * FROM [Employees]'
        $rows | Should -HaveCount 1
      }
      It 'returns zero affected count when no rows match' {
        Initialize-AccessFixture -Path $accdbPath
        $result = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'DELETE FROM [Employees] WHERE [Id] = 999'
        $result.RecordsAffected | Should -Be 0
      }
    }
    Context 'Output' {
      It 'returns row metadata and column values for table data' {
        Initialize-AccessFixture -Path $accdbPath
        $row = Invoke-AccessQuery -LiteralPath $accdbPath -Table Employees | Where-Object -Property Id -EQ 1 | Select-Object -First 1
        $row | Should -Not -BeNullOrEmpty
        $row.Path | Should -Be $accdbPath
        $row.ObjectType | Should -Be 'Table'
        $row.ObjectName | Should -Be 'Employees'
        $row.Name | Should -Be 'Alice'
        $row.Department | Should -Be 'Sales'
      }
      It 'returns row metadata and column values for view data' {
        Initialize-AccessFixture -Path $accdbPath
        $row = Invoke-AccessQuery -LiteralPath $accdbPath -View vwSalesEmployees | Select-Object -First 1
        $row | Should -Not -BeNullOrEmpty
        $row.ObjectType | Should -Be 'View'
        $row.ObjectName | Should -Be 'vwSalesEmployees'
        $row.Name | Should -Be 'Alice'
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        Initialize-AccessFixture -Path $mdbPath
        $rows = Invoke-AccessQuery -LiteralPath $mdbPath -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets table data from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Table Employees -Password $password
        $rows | Should -HaveCount 2
      }
      It 'gets view data from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -View vwSalesEmployees -Password $password
        $rows | Should -HaveCount 1
      }
      It 'gets query data from database protected with Password' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        $rows = Invoke-AccessQuery -LiteralPath $accdbPath -Query 'SELECT * FROM Employees WHERE [Id] = 1' -Password $password
        $rows | Should -HaveCount 1
      }
      It 'throws when password is incorrect' {
        Initialize-AccessFixture -Path $accdbPath -Password $password
        { Invoke-AccessQuery -LiteralPath $accdbPath -Table Employees } | Should -Throw
        { Invoke-AccessQuery -LiteralPath $accdbPath -Table Employees -Password (Get-Password) } | Should -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws when Query is used with Columns' {
        { Invoke-AccessQuery -LiteralPath $accdbPath -Query 'SELECT * FROM Employees' -Columns Id } | Should -Throw
      }
      It 'throws for missing table' {
        { Invoke-AccessQuery -LiteralPath $accdbPath -Table NotExists } | Should -Throw
      }
      It 'throws for missing view' {
        { Invoke-AccessQuery -LiteralPath $accdbPath -View NotExists } | Should -Throw
      }
    }
  }
}

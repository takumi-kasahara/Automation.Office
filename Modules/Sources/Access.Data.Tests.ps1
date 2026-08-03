using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'
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
      Initialize-AccessFixture -Path $accdbPath
      $mdbPath = Get-TempFile -Extension '.mdb'
      Initialize-AccessFixture -Path $mdbPath
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
        $tables = Get-AccessTable -Path $accdbPath
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by LiteralPath' {
        $tables = Get-AccessTable -LiteralPath $accdbPath
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipeline' {
        $tables = $accdbPath | Get-AccessTable
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        $tables = [PSCustomObject]@{ PSPath = $accdbPath } | Get-AccessTable
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        $table = Get-AccessTable -Path $accdbPath | Where-Object -Property Name -EQ 'Employees' | Select-Object -First 1
        $table | Should -Not -BeNullOrEmpty
        $table.Path | Should -Be $accdbPath
        $table.Type | Should -Not -BeNullOrEmpty
      }
      It 'works with mdb files' {
        $tables = Get-AccessTable -Path $mdbPath
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'gets tables from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $tables = Get-AccessTable -Path $path -Password $password
          ($tables | Where-Object -Property Name -EQ 'Employees') | Should -Not -BeNullOrEmpty
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is incorrect' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessTable -Path $path -Password (Get-Password) } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is not provided for protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessTable -Path $path } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $txtPath = Get-TempFile -Extension '.txt'
        try {
          Set-Content -LiteralPath $txtPath -Encoding UTF8 -Value 'x'
          { Get-AccessTable -LiteralPath $txtPath } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $txtPath) {
            Remove-Item -LiteralPath $txtPath -Force
          }
        }
      }
    }
  }
  Describe 'Get-AccessView' {
    BeforeEach {
      $path = Get-TempFile -Extension '.accdb'
      Initialize-AccessFixture -Path $path -Password $password
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets views by Path' {
        $views = Get-AccessView -Path $path
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
      It 'gets views by LiteralPath' {
        $views = Get-AccessView -LiteralPath $path
        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        $view = Get-AccessView -Path $path | Where-Object -Property Name -EQ 'vwSalesEmployees' | Select-Object -First 1
        $view | Should -Not -BeNullOrEmpty
        $view.Path | Should -Be $path
        $view.Type | Should -Match 'VIEW'
      }
    }
    Context 'Other parameters' {
      It 'gets views from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $views = Get-AccessView -Path $path -Password $password
          ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should -Not -BeNullOrEmpty
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is incorrect' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessView -Path $path -Password (Get-Password) } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is not provided for protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessView -Path $path } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
    }
  }
  Describe 'Get-AccessTableColumn' {
    BeforeEach {
      $path = Get-TempFile -Extension '.accdb'
      Initialize-AccessFixture -Path $path
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table columns by Path' {
        $columns = Get-AccessTableColumn -Path $path -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by LiteralPath' {
        $columns = Get-AccessTableColumn -LiteralPath $path -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        $columns = [PSCustomObject]@{ PSPath = $path } | Get-AccessTableColumn -Table Employees
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        $columns = Get-AccessTableColumn -Path $path -Table Employees
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should -Not -BeNullOrEmpty
        $name | Should -Not -BeNullOrEmpty
        $id.Path | Should -Be $path
        $id.ObjectType | Should -Be 'Table'
        $id.ObjectName | Should -Be 'Employees'
        $id.Ordinal | Should -BeOfType [int]
        $name.DataType | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'gets table columns from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $columns = Get-AccessTableColumn -Path $path -Table Employees -Password $password
          $columns | Should -Not -BeNullOrEmpty
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is incorrect' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessTableColumn -Path $path -Table Employees -Password (Get-Password) } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is not provided for protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessTableColumn -Path $path -Table Employees } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        { Get-AccessTableColumn -Path $path -Table NotExists } | Should -Throw
      }
    }
  }
  Describe 'Get-AccessViewColumn' {
    BeforeEach {
      $path = Get-TempFile -Extension '.accdb'
      Initialize-AccessFixture -Path $path
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets view columns by Path' {
        $columns = Get-AccessViewColumn -Path $path -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets view columns by LiteralPath' {
        $columns = Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets view columns by ValueFromPipelineByPropertyName' {
        $columns = [PSCustomObject]@{ PSPath = $path } | Get-AccessViewColumn -View vwSalesEmployees
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        $columns = Get-AccessViewColumn -Path $path -View vwSalesEmployees
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should -Not -BeNullOrEmpty
        $name | Should -Not -BeNullOrEmpty
        $id.Path | Should -Be $path
        $id.ObjectType | Should -Be 'View'
        $id.ObjectName | Should -Be 'vwSalesEmployees'
        $id.Ordinal | Should -BeOfType [int]
      }
    }
    Context 'Other parameters' {
      It 'gets view columns from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $columns = Get-AccessViewColumn -Path $path -View vwSalesEmployees -Password $password
          $columns | Should -Not -BeNullOrEmpty
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is incorrect' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessViewColumn -Path $path -View vwSalesEmployees -Password (Get-Password) } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is not provided for protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessViewColumn -Path $path -View vwSalesEmployees } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'throws for missing view' {
        { Get-AccessViewColumn -Path $path -View NotExists } | Should -Throw
      }
    }
  }
  Describe 'Get-AccessData' {
    BeforeEach {
      $path = Get-TempFile -Extension '.accdb'
      Initialize-AccessFixture -Path $path
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table data by Path' {
        $rows = Get-AccessData -Path $path -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets table data by LiteralPath' {
        $rows = Get-AccessData -LiteralPath $path -Table Employees
        $rows | Should -HaveCount 2
      }
      It 'gets view data by Path' {
        $rows = Get-AccessData -Path $path -View vwSalesEmployees
        $rows | Should -HaveCount 1
      }
      It 'gets table data by ValueFromPipelineByPropertyName' {
        $rows = [PSCustomObject]@{ PSPath = $path } | Get-AccessData -Table Employees
        $rows | Should -HaveCount 2
      }
    }

    Context 'Query' {
      It 'gets data by Query with Path' {
        $rows = Get-AccessData -Path $path -Query 'SELECT Id, Name FROM Employees WHERE Id = 1'
        $rows | Should -HaveCount 1
        $rows[0].Id | Should -Be 1
        $rows[0].Name | Should -Be 'Alice'
      }
      It 'gets data by Query with leading whitespace and lowercase select' {
        $rows = Get-AccessData -Path $path -Query "`r`n  select Id, Name FROM Employees WHERE Id = 1"
        $rows | Should -HaveCount 1
        $rows[0].Id | Should -Be 1
        $rows[0].Name | Should -Be 'Alice'
      }
      It 'gets data by Query with LiteralPath' {
        $rows = Get-AccessData -LiteralPath $path -Query 'SELECT * FROM Employees WHERE Id = 2'
        $rows | Should -HaveCount 1
        $rows[0].Id | Should -Be 2
        $rows[0].Name | Should -Be 'Bob'
      }
      It 'gets data by Query with ValueFromPipeline' {
        $rows = $path | Get-AccessData -Query 'SELECT * FROM Employees WHERE Id = 1'
        $rows | Should -HaveCount 1
        $rows[0].Name | Should -Be 'Alice'
      }
      It 'gets data by Query with ValueFromPipelineByPropertyName' {
        $rows = [PSCustomObject]@{ PSPath = $path } | Get-AccessData -Query 'SELECT * FROM Employees WHERE Id = 2'
        $rows | Should -HaveCount 1
        $rows[0].Name | Should -Be 'Bob'
      }
      It 'gets multiple rows by Query' {
        $rows = Get-AccessData -Path $path -Query 'SELECT * FROM Employees ORDER BY Id'
        $rows | Should -HaveCount 2
        $rows[0].Id | Should -Be 1
        $rows[1].Id | Should -Be 2
      }
    }
    Context 'Other parameters' {
      It 'gets data with selected columns' {
        $rows = Get-AccessData -Path $path -Table Employees -Columns Id, Name
        $rows | Should -HaveCount 2
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Id'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Name'
        @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'Department'
      }
      It 'gets view data with selected columns' {
        $rows = Get-AccessData -Path $path -View vwSalesEmployees -Columns Id, Name
        $rows | Should -HaveCount 1
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Id'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Name'
      }
    }
    Context 'Output' {
      It 'returns row metadata and column values for table data' {
        $row = Get-AccessData -Path $path -Table Employees | Where-Object -Property Id -EQ 1 | Select-Object -First 1
        $row | Should -Not -BeNullOrEmpty
        $row.Path | Should -Be $path
        $row.ObjectType | Should -Be 'Table'
        $row.ObjectName | Should -Be 'Employees'
        $row.Name | Should -Be 'Alice'
        $row.Department | Should -Be 'Sales'
      }
      It 'returns row metadata and column values for view data' {
        $row = Get-AccessData -Path $path -View vwSalesEmployees | Select-Object -First 1
        $row | Should -Not -BeNullOrEmpty
        $row.ObjectType | Should -Be 'View'
        $row.ObjectName | Should -Be 'vwSalesEmployees'
        $row.Name | Should -Be 'Alice'
      }
      It 'returns query metadata for Query parameter set' {
        $row = Get-AccessData -Path $path -Query 'SELECT Id, Name FROM Employees WHERE Id = 1' | Select-Object -First 1
        $row | Should -Not -BeNullOrEmpty
        $row.ObjectType | Should -Be 'Query'
        $row.ObjectName | Should -Be 'SELECT Id, Name FROM Employees WHERE Id = 1'
      }
    }
    Context 'Other parameters' {
      It 'gets table data from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $rows = Get-AccessData -Path $path -Table Employees -Password $password
          $rows | Should -HaveCount 2
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'gets view data from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $rows = Get-AccessData -Path $path -View vwSalesEmployees -Password $password
          $rows | Should -HaveCount 1
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'gets query data from password-protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          $rows = Get-AccessData -Path $path -Query 'SELECT * FROM Employees WHERE Id = 1' -Password $password
          $rows | Should -HaveCount 1
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is incorrect' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessData -Path $path -Table Employees -Password (Get-Password) } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
      It 'throws when password is not provided for protected database' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.accdb'
        try {
          Initialize-AccessFixture -Path $path -Password $password
          { Get-AccessData -Path $path -Table Employees } | Should -Throw
        } finally {
          if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'throws when Query is not a SELECT statement' {
        { Get-AccessData -Path $path -Query 'DELETE FROM Employees' } | Should -Throw
      }
      It 'throws when Query is used with Columns' {
        { Get-AccessData -Path $path -Query 'SELECT * FROM Employees' -Columns Id } | Should -Throw
      }
      It 'throws for missing table' {
        { Get-AccessData -Path $path -Table NotExists } | Should -Throw
      }
      It 'throws for missing view' {
        { Get-AccessData -Path $path -View NotExists } | Should -Throw
      }
    }
  }
}

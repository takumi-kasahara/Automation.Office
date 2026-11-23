using assembly Microsoft.Office.Interop.Access
using module .\..\Automation.Office.psd1
using namespace Microsoft.Office.Interop.Access
using namespace Microsoft.Office.Interop.Access.Dao
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()
Set-StrictMode -Version Latest

InModuleScope 'Automation.Office' {
  BeforeAll {
    function Get-Password {
      [CmdletBinding()]
      [OutputType([System.Security.SecureString])]
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
      param (
        [string]
        $Path,
        [SecureString]
        $Password = $null
      )
      New-AccessFile -Path $Path -Password $Password -Force -InitializeDb {
        param (
          [Microsoft.Office.Interop.Access.Dao.Database]
          $Database
        )
        $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255), Department TEXT(255))')
        $Database.Execute("INSERT INTO Employees (Id, Name, Department) VALUES (1, 'Alice', 'Sales')")
        $Database.Execute("INSERT INTO Employees (Id, Name, Department) VALUES (2, 'Bob', 'Engineering')")
        $Database.CreateQueryDef('vwSalesEmployees', "SELECT Id, Name FROM Employees WHERE Department = 'Sales'")
      }
    }
  }
  Describe 'Get-AccessTable' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile -Extension '.accdb'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets tables by Path' {
        Initialize-AccessFixture -Path $path

        $tables = Get-AccessTable -Path $path

        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'gets tables by LiteralPath' {
        Initialize-AccessFixture -Path $path

        $tables = Get-AccessTable -LiteralPath $path

        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'gets tables by ValueFromPipeline' {
        Initialize-AccessFixture -Path $path

        $tables = $path | Get-AccessTable

        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $path

        $tables = [PSCustomObject]@{ PSPath = $path } | Get-AccessTable

        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-AccessFixture -Path $path

        $table = Get-AccessTable -LiteralPath $path | Where-Object -Property Name -EQ 'Employees' | Select-Object -First 1

        $table | Should-NotBeNull
        $table.Path | Should-BeString $path
        $table.Type | Should-BeString 'TABLE'
        $table.Name | Should-BeString 'Employees'
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        $path = Get-TempFile -Extension '.mdb'
        Initialize-AccessFixture -Path $path

        $tables = Get-AccessTable -LiteralPath $path

        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'gets tables from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        { Get-AccessTable -LiteralPath $path } | Should-Throw
        { Get-AccessTable -LiteralPath $path -Password (Get-Password) } | Should-Throw
        { Get-AccessTable -LiteralPath $path -Password $password } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $path = Get-TempFile -Extension '.txt'
        New-Item -Path $path -ItemType File -Force | Out-Null

        { Get-AccessTable -LiteralPath $path } | Should-Throw
      }
    }
  }
  Describe 'Get-AccessTableColumn' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile -Extension '.accdb'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table columns by Path' {
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessTableColumn -Path $path -Table 'Employees'

        $columns | Should-NotBeNull
      }
      It 'gets table columns by LiteralPath' {
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessTableColumn -LiteralPath $path -Table 'Employees'

        $columns | Should-NotBeNull
      }
      It 'gets table columns by ValueFromPipeline' {
        Initialize-AccessFixture -Path $path

        $columns = $path | Get-AccessTableColumn -Table 'Employees'

        $columns | Should-NotBeNull
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $path

        $columns = [PSCustomObject]@{ PSPath = $path } | Get-AccessTableColumn -Table 'Employees'

        $columns | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessTableColumn -LiteralPath $path -Table 'Employees'

        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $id | Should-NotBeNull
        $id.Path | Should-BeString $path
        $id.ObjectType | Should-BeString 'Table'
        $id.ObjectName | Should-BeString 'Employees'
        $id.ColumnName | Should-BeString 'Id'
        $id.Ordinal | Should-HaveType ([int])
        $id.DataType | Should-NotBeNull
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $name | Should-NotBeNull
        $name.Path | Should-BeString $path
        $name.ObjectType | Should-BeString 'Table'
        $name.ObjectName | Should-BeString 'Employees'
        $name.ColumnName | Should-BeString 'Name'
        $name.Ordinal | Should-HaveType ([int])
        $name.DataType | Should-NotBeNull
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        $path = Get-TempFile -Extension '.mdb'
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessTableColumn -LiteralPath $path -Table 'Employees'

        $columns | Should-NotBeNull
      }
      It 'gets table columns from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        { Get-AccessTableColumn -LiteralPath $path -Table 'Employees' } | Should-Throw
        { Get-AccessTableColumn -LiteralPath $path -Table 'Employees' -Password (Get-Password) } | Should-Throw
        { Get-AccessTableColumn -LiteralPath $path -Table 'Employees' -Password $password } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        Initialize-AccessFixture -Path $path

        { Get-AccessTableColumn -LiteralPath $path -Table 'Dummy' } | Should-Throw
      }
    }
  }
  Describe 'Invoke-AccessQuery' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile -Extension '.accdb'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table data by Path' {
        Initialize-AccessFixture -Path $path

        $rows = Invoke-AccessQuery -Path $path -Table 'Employees'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets table data by LiteralPath' {
        Initialize-AccessFixture -Path $path

        $rows = Invoke-AccessQuery -LiteralPath $path -Table 'Employees'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets table data by ValueFromPipeline' {
        Initialize-AccessFixture -Path $path

        $rows = $path | Invoke-AccessQuery -Table 'Employees'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets table data by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $path

        $rows = [PSCustomObject]@{ PSPath = $path } | Invoke-AccessQuery -Table 'Employees'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets table data with selected columns' {
        Initialize-AccessFixture -Path $path

        $rows = Invoke-AccessQuery -LiteralPath $path -Table 'Employees' -Columns Id, Name

        $rows | Should-BeCollection -Count 2
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Id'
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Name'
        $rows[0].PSObject.Properties.Name | Should-NotContainCollection 'Department'
      }
      It 'gets view data with selected columns' {
        Initialize-AccessFixture -Path $path

        $rows = Invoke-AccessQuery -LiteralPath $path -View vwSalesEmployees -Columns Id, Name

        $rows | Should-BeCollection -Count 1
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Id'
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Name'
      }
      It 'gets data by Query' {
        Initialize-AccessFixture -Path $path

        $select = 'SELECT [Id], [Name] FROM Employees WHERE [Id] = 1'
        $rows = Invoke-AccessQuery -LiteralPath $path -Query $select

        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Query'
        $rows[0].ObjectName | Should-BeString $select
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'

      }
    }
    Context 'Input' {
      It 'inserts rows and returns affected count' {
        $insert = "INSERT INTO [Employees] (Id, Name, Department) VALUES (3, 'Carol', 'Marketing')"
        $select = 'SELECT * FROM [Employees] WHERE [Id] = 3'
        Initialize-AccessFixture -Path $path

        $result = Invoke-AccessQuery -LiteralPath $path -Query $insert
        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-BeString $insert
        $result.RecordsAffected | Should-Be 1
        $rows = Invoke-AccessQuery -LiteralPath $path -Query $select
        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Query'
        $rows[0].ObjectName | Should-BeString $select
        $rows[0].Id | Should-Be 3
        $rows[0].Name | Should-BeString 'Carol'
        $rows[0].Department | Should-BeString 'Marketing'
      }
      It 'updates rows and returns affected count' {
        $update = "UPDATE [Employees] SET [Department] = 'Marketing' WHERE [Id] = 1"
        $select = 'SELECT * FROM [Employees] WHERE [Id] = 1'
        Initialize-AccessFixture -Path $path

        $result = Invoke-AccessQuery -LiteralPath $path -Query $update
        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-BeString $update
        $result.RecordsAffected | Should-Be 1
        $rows = Invoke-AccessQuery -LiteralPath $path -Query $select
        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Query'
        $rows[0].ObjectName | Should-BeString $select
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Marketing'
      }
      It 'deletes rows and returns affected count' {
        Initialize-AccessFixture -Path $path

        $delete = 'DELETE FROM [Employees] WHERE [Id] = 2'
        $result = Invoke-AccessQuery -LiteralPath $path -Query $delete

        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-Be $delete
        $result.RecordsAffected | Should-Be 1
      }
      It 'returns zero affected count when no rows match' {
        Initialize-AccessFixture -Path $path

        $delete = 'DELETE FROM [Employees] WHERE [Id] = 999'
        $result = Invoke-AccessQuery -LiteralPath $path -Query $delete

        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-Be $delete
        $result.RecordsAffected | Should-Be 0
      }
    }
    Context 'Output' {
      It 'returns row metadata and column values for table data' {
        Initialize-AccessFixture -Path $path

        $row = Invoke-AccessQuery -LiteralPath $path -Table 'Employees' | Where-Object -Property Id -EQ 1 | Select-Object -First 1

        $row | Should-NotBeNull
        $row.Path | Should-BeString $path
        $row.ObjectType | Should-BeString 'Table'
        $row.ObjectName | Should-BeString 'Employees'
        $row.Id | Should-Be 1
        $row.Name | Should-BeString 'Alice'
      }
      It 'returns row metadata and column values for view data' {
        Initialize-AccessFixture -Path $path

        $row = Invoke-AccessQuery -LiteralPath $path -View vwSalesEmployees | Select-Object -First 1

        $row | Should-NotBeNull
        $row.Path | Should-BeString $path
        $row.ObjectType | Should-BeString 'View'
        $row.ObjectName | Should-BeString 'vwSalesEmployees'
        $row.Id | Should-Be 1
        $row.Name | Should-BeString 'Alice'
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        $path = Get-TempFile -Extension '.mdb'
        Initialize-AccessFixture -Path $path

        $rows = Invoke-AccessQuery -LiteralPath $path -Table 'Employees'

        $rows | Should-BeCollection -Count 2
      }
      It 'gets table data from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        $rows = Invoke-AccessQuery -LiteralPath $path -Table 'Employees' -Password $password

        $rows | Should-BeCollection -Count 2
      }
      It 'gets view data from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        $rows = Invoke-AccessQuery -LiteralPath $path -View vwSalesEmployees -Password $password

        $rows | Should-BeCollection -Count 1
      }
      It 'gets query data from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        { Invoke-AccessQuery -LiteralPath $path -Table 'Employees' } | Should-Throw
        { Invoke-AccessQuery -LiteralPath $path -Table 'Employees' -Password (Get-Password) } | Should-Throw
        { Invoke-AccessQuery -LiteralPath $path -Table 'Employees' -Password $password } | Should -Not -Throw
      }
    }
    Context 'Edge cases' -Tag 'CI', 'Mock' {
      It 'throws when Query is used with Columns' {
        { Invoke-AccessQuery -LiteralPath $path -Query 'SELECT * FROM Employees' -Columns Id } | Should-Throw
      }
      It 'throws for missing table' {
        { Invoke-AccessQuery -LiteralPath $path -Table 'Dummy' } | Should-Throw
      }
      It 'throws for missing view' {
        { Invoke-AccessQuery -LiteralPath $path -View 'Dummy' } | Should-Throw
      }
    }
  }
  Describe 'Get-AccessView' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile -Extension '.accdb'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets views by Path' {
        Initialize-AccessFixture -Path $path

        $views = Get-AccessView -Path $path

        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should-NotBeNull
      }
      It 'gets views by LiteralPath' {
        Initialize-AccessFixture -Path $path

        $views = Get-AccessView -LiteralPath $path

        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should-NotBeNull
      }
      It 'gets views by ValueFromPipeline' {
        Initialize-AccessFixture -Path $path

        $views = $path | Get-AccessView

        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should-NotBeNull
      }
      It 'gets views by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $path

        $views = [PSCustomObject]@{ PSPath = $path } | Get-AccessView

        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-AccessFixture -Path $path

        $view = Get-AccessView -LiteralPath $path | Where-Object -Property Name -EQ 'vwSalesEmployees' | Select-Object -First 1

        $view | Should-NotBeNull
        $view.Path | Should-BeString $path
        $view.Type | Should-BeString 'VIEW'
        $view.Name | Should-BeString 'vwSalesEmployees'
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        $path = Get-TempFile -Extension '.mdb'
        Initialize-AccessFixture -Path $path

        $views = Get-AccessView -LiteralPath $path

        ($views | Where-Object -Property Name -EQ 'vwSalesEmployees') | Should-NotBeNull
      }
      It 'gets views from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        { Get-AccessView -LiteralPath $path } | Should-Throw
        { Get-AccessView -LiteralPath $path -Password (Get-Password) } | Should-Throw
        { Get-AccessView -LiteralPath $path -Password $password } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $path = Get-TempFile -Extension '.txt'

        New-Item -Path $path -ItemType File -Force | Out-Null

        { Get-AccessView -LiteralPath $path } | Should-Throw
      }
    }
  }
  Describe 'Get-AccessViewColumn' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile -Extension '.accdb'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets view columns by Path' {
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessViewColumn -Path $path -View vwSalesEmployees

        $columns | Should-NotBeNull
      }
      It 'gets view columns by LiteralPath' {
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees

        $columns | Should-NotBeNull
      }
      It 'gets view columns by ValueFromPipeline' {
        Initialize-AccessFixture -Path $path

        $columns = $path | Get-AccessViewColumn -View vwSalesEmployees

        $columns | Should-NotBeNull
      }
      It 'gets view columns by ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $path

        $columns = [PSCustomObject]@{ PSPath = $path } | Get-AccessViewColumn -View vwSalesEmployees

        $columns | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees

        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $id | Should-NotBeNull
        $id.Path | Should-BeString $path
        $id.ObjectType | Should-BeString 'View'
        $id.ObjectName | Should-BeString 'vwSalesEmployees'
        $id.ColumnName | Should-BeString 'Id'
        $id.Ordinal | Should-HaveType ([int])
        $id.DataType | Should-NotBeNull
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $name | Should-NotBeNull
        $name.Path | Should-BeString $path
        $name.ObjectType | Should-BeString 'View'
        $name.ObjectName | Should-BeString 'vwSalesEmployees'
        $name.ColumnName | Should-BeString 'Name'
        $name.Ordinal | Should-HaveType ([int])
        $name.DataType | Should-NotBeNull
      }
    }
    Context 'Other parameters' {
      It 'works with mdb files' {
        $path = Get-TempFile -Extension '.mdb'
        Initialize-AccessFixture -Path $path

        $columns = Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees

        $columns | Should-NotBeNull
      }
      It 'gets view columns from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        { Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees } | Should-Throw
        { Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees -Password (Get-Password) } | Should-Throw
        { Get-AccessViewColumn -LiteralPath $path -View vwSalesEmployees -Password $password } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for missing view' {
        Initialize-AccessFixture -Path $path

        { Get-AccessViewColumn -LiteralPath $path -View 'Dummy' } | Should-Throw
      }
    }
  }
}

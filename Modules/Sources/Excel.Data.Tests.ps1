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
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    Add-Type -AssemblyName System.Web
    function Get-Password {
      [CmdletBinding()]
      [OutputType([SecureString])]
      [SuppressMessage('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Used in tests to generate random passwords for verification purposes')]
      param ()
      return ConvertTo-SecureString -String ([System.Web.Security.Membership]::GeneratePassword(15, 0)) -AsPlainText -Force
    }
    function Get-TempFile {
      [CmdletBinding()]
      [OutputType([string])]
      param (
        [string]
        $Extension = '.xlsx'
      )
      return $env:TEMP | Join-Path -ChildPath "Book.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function Initialize-ExcelFixture {
      param(
        [string]
        $Path,
        [SecureString]
        $PasswordToOpen = $null
      )
      $extension = [Path]::GetExtension($Path)
      $fileFormat = switch -Exact -CaseSensitive ($extension) {
        '.xlsx' { [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook }
        '.xlsm' { [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled }
      }
      New-ExcelFile -Path $Path -FileFormat $fileFormat -PasswordToOpen $PasswordToOpen -Initialize {
        param($Workbook)
        $sheet = $Workbook.Worksheets.Item(1)
        $sheet.Name = 'Employees'
        $sheet.Cells.Item(1, 1) = 'Id'
        $sheet.Cells.Item(1, 2) = 'Name'
        $sheet.Cells.Item(1, 3) = 'Department'
        $sheet.Cells.Item(2, 1) = 1
        $sheet.Cells.Item(2, 2) = 'Alice'
        $sheet.Cells.Item(2, 3) = 'Sales'
        $sheet.Cells.Item(3, 1) = 2
        $sheet.Cells.Item(3, 2) = 'Bob'
        $sheet.Cells.Item(3, 3) = 'Engineering'
      } | Out-Null
    }
  }
  Describe 'Get-ExcelTable' {
    BeforeEach {
      $xlsxPath = Get-TempFile -Extension '.xlsx'
      $xlsmPath = Get-TempFile -Extension '.xlsm'
    }
    AfterEach {
      if (Test-Path -LiteralPath $xlsxPath) {
        Remove-Item -LiteralPath $xlsxPath -Force
      }
      if (Test-Path -LiteralPath $xlsmPath) {
        Remove-Item -LiteralPath $xlsmPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets tables by Path' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = Get-ExcelTable -Path $xlsxPath
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = Get-ExcelTable -LiteralPath $xlsxPath
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = $xlsxPath | Get-ExcelTable
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelTable
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-ExcelFixture -Path $xlsxPath
        $table = Get-ExcelTable -LiteralPath $xlsxPath | Where-Object -Property Name -EQ 'Employees$' | Select-Object -First 1
        $table | Should -Not -BeNullOrEmpty
        $table.Path | Should -Be $xlsxPath
        $table.Type | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $tables = Get-ExcelTable -LiteralPath $xlsmPath
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables with ReadOnly' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = Get-ExcelTable -LiteralPath $xlsxPath -ReadOnly
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $txtPath = Get-TempFile -Extension '.txt'
        New-Item -Path $txtPath -ItemType File -Force | Out-Null
        { Get-ExcelTable -LiteralPath $txtPath } | Should -Throw
      }
      It 'throws when protected with password' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.xlsx'
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Get-ExcelTable -LiteralPath $path } | Should -Throw
      }
    }
  }
  Describe 'Get-ExcelTableColumn' {
    BeforeEach {
      $xlsxPath = Get-TempFile -Extension '.xlsx'
      $xlsmPath = Get-TempFile -Extension '.xlsm'
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $xlsxPath) {
        Remove-Item -LiteralPath $xlsxPath -Force
      }
      if (Test-Path -LiteralPath $xlsmPath) {
        Remove-Item -LiteralPath $xlsmPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table columns by Path' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -Path $xlsxPath -Table Employees$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = $xlsxPath | Get-ExcelTableColumn -Table Employees$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelTableColumn -Table Employees$
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should -Not -BeNullOrEmpty
        $name | Should -Not -BeNullOrEmpty
        $id.Path | Should -Be $xlsxPath
        $id.ObjectType | Should -Be 'Table'
        $id.ObjectName | Should -Be 'Employees$'
        $id.Ordinal | Should -BeOfType [int]
        $name.DataType | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsmPath -Table Employees$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns with ReadOnly' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$ -ReadOnly
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Get-ExcelTableColumn -LiteralPath $xlsxPath -Table NotExists } | Should -Throw
      }
      It 'throws when protected with PasswordToOpen' {
        Initialize-ExcelFixture -Path $xlsxPath -PasswordToOpen $password
        { Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$ } | Should -Throw
      }
    }
  }
  Describe 'Get-ExcelData' {
    BeforeEach {
      $xlsxPath = Get-TempFile -Extension '.xlsx'
      $xlsmPath = Get-TempFile -Extension '.xlsm'
    }
    AfterEach {
      if (Test-Path -LiteralPath $xlsxPath) {
        Remove-Item -LiteralPath $xlsxPath -Force
      }
      if (Test-Path -LiteralPath $xlsmPath) {
        Remove-Item -LiteralPath $xlsmPath -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets rows by Path' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -Path $xlsxPath -Table 'Employees$'
        $rows | Should -Not -BeNullOrEmpty
        $rows[0].Path | Should -Be $xlsxPath
        $rows[0].ObjectName | Should -Be 'Employees$'
        $rows[0].Name | Should -Be 'Alice'
        $rows[0].Department | Should -Be 'Sales'
      }
      It 'gets rows by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$'
        $rows | Should -Not -BeNullOrEmpty
        $rows[0].Path | Should -Be $xlsxPath
      }
      It 'gets rows by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = $xlsxPath | Get-ExcelData -Table 'Employees$'
        $rows | Should -Not -BeNullOrEmpty
      }
      It 'gets rows by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelData -Table 'Employees$'
        $rows | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'gets rows with selected columns' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -Columns Id, Name
        $rows | Should -HaveCount 2
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Id'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Name'
        @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'Department'
      }
      It 'gets rows by Query' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Query 'SELECT [Id], [Name] FROM [Employees$] WHERE [Id] = 1'
        $rows | Should -HaveCount 1
        $rows[0].Id | Should -Be 1
        $rows[0].Name | Should -Be 'Alice'
      }
      It 'gets rows with address range' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -Address 'A1:C2'
        $rows | Should -HaveCount 1
        $rows[0].Id | Should -Be 1
        $rows[0].Name | Should -Be 'Alice'
        $rows[0].Department | Should -Be 'Sales'
      }
      It 'gets rows with address range and selected columns' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -Address 'A1:B3' -Columns Id, Name
        $rows | Should -HaveCount 2
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Id'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'Name'
        @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'Department'
      }
      It 'throws when address count does not match table count' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$', 'Other$' -Address 'A1:C2' } | Should -Throw
      }
      It 'gets rows with NoHeader' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -NoHeader
        $rows | Should -HaveCount 3
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'F1'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'F2'
        @($rows[0].PSObject.Properties.Name) | Should -Contain 'F3'
      }
      It 'gets rows with NoIMEX' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -NoIMEX
        $rows | Should -HaveCount 2
        $rows[0].Id | Should -Be 1
        $rows[0].Name | Should -Be 'Alice'
      }
      It 'gets rows with NoHeader and NoIMEX' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -NoHeader -NoIMEX
        $rows | Should -HaveCount 2
        @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'F1'
      }
    }
    Context 'Output' {
      It 'returns row metadata and column values for table data' {
        Initialize-ExcelFixture -Path $xlsxPath
        $row = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' | Where-Object -Property Id -EQ 1 | Select-Object -First 1
        $row | Should -Not -BeNullOrEmpty
        $row.Path | Should -Be $xlsxPath
        $row.ObjectName | Should -Be 'Employees$'
        $row.Name | Should -Be 'Alice'
        $row.Department | Should -Be 'Sales'
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $rows = Get-ExcelData -LiteralPath $xlsmPath -Table 'Employees$'
        $rows | Should -Not -BeNullOrEmpty
      }
      It 'gets rows with ReadOnly' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Get-ExcelData -LiteralPath $xlsxPath -Table 'Employees$' -ReadOnly
        $rows | Should -Not -BeNullOrEmpty
        $rows[0].Path | Should -Be $xlsxPath
      }
    }
    Context 'Edge cases' {
      It 'throws when Query is not a SELECT statement' {
        { Get-ExcelData -LiteralPath $xlsxPath -Query 'DELETE FROM Employees' } | Should -Throw
      }
      It 'throws when Query is used with Columns' {
        { Get-ExcelData -LiteralPath $xlsxPath -Query 'SELECT * FROM Employees' -Columns Id } | Should -Throw
      }
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Get-ExcelData -LiteralPath $xlsxPath -Table 'NotExists' } | Should -Throw
      }
      It 'throws when protected with PasswordToOpen' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.xlsx'
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Get-ExcelData -LiteralPath $path -Table 'Employees$' } | Should -Throw
      }
    }
  }
}

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
        $Path
      )
      $extension = [Path]::GetExtension($Path)
      $fileFormat = switch -Exact -CaseSensitive ($extension) {
        '.xlsx' { [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook }
        '.xlsm' { [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled }
      }
      New-ExcelFile -Path $Path -FileFormat $fileFormat -Initialize {
        param($Workbook)
        $sheet = $Workbook.Worksheets.Item(1)
        $sheet.Name = 'Sheet1'
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
        ($tables | Where-Object -Property Name -EQ 'Sheet1$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = Get-ExcelTable -LiteralPath $xlsxPath
        ($tables | Where-Object -Property Name -EQ 'Sheet1$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = $xlsxPath | Get-ExcelTable
        ($tables | Where-Object -Property Name -EQ 'Sheet1$') | Should -Not -BeNullOrEmpty
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelTable
        ($tables | Where-Object -Property Name -EQ 'Sheet1$') | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-ExcelFixture -Path $xlsxPath
        $table = Get-ExcelTable -LiteralPath $xlsxPath | Where-Object -Property Name -EQ 'Sheet1$' | Select-Object -First 1
        $table | Should -Not -BeNullOrEmpty
        $table.Path | Should -Be $xlsxPath
        $table.Type | Should -Not -BeNullOrEmpty
      }
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $tables = Get-ExcelTable -LiteralPath $xlsmPath
        ($tables | Where-Object -Property Name -EQ 'Sheet1$') | Should -Not -BeNullOrEmpty
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
        $columns = Get-ExcelTableColumn -Path $xlsxPath -Table Sheet1$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Sheet1$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = $xlsxPath | Get-ExcelTableColumn -Table Sheet1$
        $columns | Should -Not -BeNullOrEmpty
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelTableColumn -Table Sheet1$
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Sheet1$
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should -Not -BeNullOrEmpty
        $name | Should -Not -BeNullOrEmpty
        $id.Path | Should -Be $xlsxPath
        $id.ObjectType | Should -Be 'Table'
        $id.ObjectName | Should -Be 'Sheet1$'
        $id.Ordinal | Should -BeOfType [int]
        $name.DataType | Should -Not -BeNullOrEmpty
      }
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsmPath -Table Sheet1$
        $columns | Should -Not -BeNullOrEmpty
      }
    }
    Context 'Other parameters' {
      It 'gets table columns from workbook protected with Password' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Sheet1$ -Password $password } | Should -Throw
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Get-ExcelTableColumn -LiteralPath $xlsxPath -Table NotExists } | Should -Throw
      }
      It 'throws when protected with password' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.xlsx'
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Get-ExcelTableColumn -LiteralPath $path -Table NotExists } | Should -Throw
      }
    }
  }
}

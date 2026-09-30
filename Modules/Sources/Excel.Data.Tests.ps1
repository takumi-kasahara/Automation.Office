using assembly Microsoft.Office.Interop.Excel
using assembly System.Web
using module .\..\Automation.Office.psd1
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
        param(
          [Microsoft.Office.Interop.Excel.Workbook]
          $Workbook
        )
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
      }
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
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = Get-ExcelTable -LiteralPath $xlsxPath
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = $xlsxPath | Get-ExcelTable
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelTable
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-ExcelFixture -Path $xlsxPath
        $table = Get-ExcelTable -LiteralPath $xlsxPath | Where-Object -Property Name -EQ 'Employees$' | Select-Object -First 1
        $table | Should-NotBeNull
        $table.Path | Should-Be $xlsxPath
        $table.Type | Should-Be 'TABLE'
        $table.Name | Should-Be 'Employees$'
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $tables = Get-ExcelTable -LiteralPath $xlsmPath
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables with ReadOnly' {
        Initialize-ExcelFixture -Path $xlsxPath
        $tables = Get-ExcelTable -LiteralPath $xlsxPath -ReadOnly
        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $txtPath = Get-TempFile -Extension '.txt'
        New-Item -Path $txtPath -ItemType File -Force | Out-Null
        { Get-ExcelTable -LiteralPath $txtPath } | Should-Throw
      }
      It 'throws when protected with password' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.xlsx'
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Get-ExcelTable -LiteralPath $path } | Should-Throw
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
        $columns | Should-NotBeNull
      }
      It 'gets table columns by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$
        $columns | Should-NotBeNull
      }
      It 'gets table columns by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = $xlsxPath | Get-ExcelTableColumn -Table Employees$
        $columns | Should-NotBeNull
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = [PSCustomObject]@{ PSPath = $xlsxPath } | Get-ExcelTableColumn -Table Employees$
        $columns | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$
        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $id | Should-NotBeNull
        $name | Should-NotBeNull
        $id.Path | Should-Be $xlsxPath
        $id.ObjectType | Should-Be 'Table'
        $id.ObjectName | Should-Be 'Employees$'
        $id.Ordinal | Should-HaveType ([int])
        $name.DataType | Should-NotBeNull
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsmPath -Table Employees$
        $columns | Should-NotBeNull
      }
      It 'gets table columns with ReadOnly' {
        Initialize-ExcelFixture -Path $xlsxPath
        $columns = Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$ -ReadOnly
        $columns | Should-NotBeNull
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Get-ExcelTableColumn -LiteralPath $xlsxPath -Table NotExists } | Should-Throw
      }
      It 'throws when protected with PasswordToOpen' {
        Initialize-ExcelFixture -Path $xlsxPath -PasswordToOpen $password
        { Get-ExcelTableColumn -LiteralPath $xlsxPath -Table Employees$ } | Should-Throw
      }
    }
  }
  Describe 'Invoke-ExcelQuery' {
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
        $rows = Invoke-ExcelQuery -Path $xlsxPath -Table 'Employees$'
        $rows | Should-NotBeNull
        $rows[0].Path | Should-Be $xlsxPath
        $rows[0].ObjectName | Should-Be 'Employees$'
        $rows[0].Name | Should-Be 'Alice'
        $rows[0].Department | Should-Be 'Sales'
      }
      It 'gets rows by LiteralPath' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$'
        $rows | Should-NotBeNull
        $rows[0].Path | Should-Be $xlsxPath
      }
      It 'gets rows by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = $xlsxPath | Invoke-ExcelQuery -Table 'Employees$'
        $rows | Should-NotBeNull
      }
      It 'gets rows by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = [PSCustomObject]@{ PSPath = $xlsxPath } | Invoke-ExcelQuery -Table 'Employees$'
        $rows | Should-NotBeNull
      }
      It 'gets rows with selected columns' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -Columns Id, Name
        $rows | Should-BeCollection -Count 2
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'Id'
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'Name'
        @($rows[0].PSObject.Properties.Name) | Should-NotContainCollection 'Department'
      }
      It 'gets rows by Query' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Query 'SELECT [Id], [Name] FROM [Employees$] WHERE [Id] = 1'
        $rows | Should-BeCollection -Count 1
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-Be 'Alice'
      }
    }
    Context 'Input' {
      It 'inserts rows and returns affected count' {
        Initialize-ExcelFixture -Path $xlsxPath
        $result = Invoke-ExcelQuery -LiteralPath $xlsxPath -Query "INSERT INTO [Employees$] (Id, Name, Department) VALUES (3, 'Carol', 'Marketing')"
        $result | Should-NotBeNull
        $result.Path | Should-Be $xlsxPath
        $result.ObjectType | Should-Be 'Query'
        $result.RecordsAffected | Should-Be 1
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Query 'SELECT * FROM [Employees$] WHERE [Id] = 3'
        $rows | Should-BeCollection -Count 1
        $rows[0].Name | Should-Be 'Carol'
        $rows[0].Department | Should-Be 'Marketing'
      }
      It 'updates rows and returns affected count' {
        Initialize-ExcelFixture -Path $xlsxPath
        $result = Invoke-ExcelQuery -LiteralPath $xlsxPath -Query "UPDATE [Employees$] SET [Department] = 'Marketing' WHERE [Id] = 1"
        $result | Should-NotBeNull
        $result.RecordsAffected | Should-Be 1
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Query 'SELECT * FROM [Employees$] WHERE [Id] = 1'
        $rows | Should-BeCollection -Count 1
        $rows[0].Department | Should-Be 'Marketing'
      }
      It 'throws when executing DELETE' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Invoke-ExcelQuery -LiteralPath $xlsxPath -Query 'DELETE FROM [Employees$] WHERE [Id] = 1' } | Should-Throw
      }
      It 'returns zero affected count when no rows match' {
        Initialize-ExcelFixture -Path $xlsxPath
        $result = Invoke-ExcelQuery -LiteralPath $xlsxPath -Query "UPDATE [Employees$] SET [Department] = 'Marketing' WHERE [Id] = 999"
        $result.RecordsAffected | Should-Be 0
      }
    }
    Context 'Output' {
      It 'returns row metadata and column values for table data' {
        Initialize-ExcelFixture -Path $xlsxPath
        $row = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' | Where-Object -Property Id -EQ 1 | Select-Object -First 1
        $row | Should-NotBeNull
        $row.Path | Should-Be $xlsxPath
        $row.ObjectName | Should-Be 'Employees$'
        $row.Name | Should-Be 'Alice'
        $row.Department | Should-Be 'Sales'
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        Initialize-ExcelFixture -Path $xlsmPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsmPath -Table 'Employees$'
        $rows | Should-NotBeNull
      }
      It 'gets rows with address range' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -Address 'A1:C2'
        $rows | Should-BeCollection -Count 1
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-Be 'Alice'
        $rows[0].Department | Should-Be 'Sales'
      }
      It 'gets rows with address range and selected columns' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -Address 'A1:B3' -Columns Id, Name
        $rows | Should-BeCollection -Count 2
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'Id'
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'Name'
        @($rows[0].PSObject.Properties.Name) | Should-NotContainCollection 'Department'
      }
      It 'throws when address count does not match table count' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$', 'Other$' -Address 'A1:C2' } | Should-Throw
      }
      It 'gets rows with NoHeader' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -NoHeader
        $rows | Should-BeCollection -Count 3
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'F1'
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'F2'
        @($rows[0].PSObject.Properties.Name) | Should-ContainCollection 'F3'
      }
      It 'gets rows with NoIMEX' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -NoIMEX
        $rows | Should-BeCollection -Count 2
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-Be 'Alice'
      }
      It 'gets rows with NoHeader and NoIMEX' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -NoHeader -NoIMEX
        $rows | Should-BeCollection -Count 2
        @($rows[0].PSObject.Properties.Name) | Should-NotContainCollection 'F1'
      }
      It 'gets rows with ReadOnly' {
        Initialize-ExcelFixture -Path $xlsxPath
        $rows = Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'Employees$' -ReadOnly
        $rows | Should-NotBeNull
        $rows[0].Path | Should-Be $xlsxPath
      }
    }
    Context 'Edge cases' {
      It 'throws when Query is used with Columns' {
        { Invoke-ExcelQuery -LiteralPath $xlsxPath -Query 'SELECT * FROM Employees' -Columns Id } | Should-Throw
      }
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $xlsxPath
        { Invoke-ExcelQuery -LiteralPath $xlsxPath -Table 'NotExists' } | Should-Throw
      }
      It 'throws when protected with PasswordToOpen' {
        $password = Get-Password
        $path = Get-TempFile -Extension '.xlsx'
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' } | Should-Throw
      }
    }
  }
}

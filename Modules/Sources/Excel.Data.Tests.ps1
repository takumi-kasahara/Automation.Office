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
      param (
        [string]
        $Path,
        [SecureString]
        $PasswordToOpen = $null
      )
      $extension = [Path]::GetExtension($Path)
      $fileFormat = switch -Exact -CaseSensitive ($extension) {
        '.xlsx' {
          [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
        }
        '.xlsm' {
          [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled
        }
      }
      New-ExcelFile -Path $Path -FileFormat $fileFormat -PasswordToOpen $PasswordToOpen -Initialize {
        param (
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
      $path = Get-TempFile -Extension '.xlsx'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets tables by Path' {
        Initialize-ExcelFixture -Path $path

        $tables = Get-ExcelTable -Path $path

        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables by LiteralPath' {
        Initialize-ExcelFixture -Path $path

        $tables = Get-ExcelTable -LiteralPath $path

        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $path

        $tables = $path | Get-ExcelTable

        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $path

        $tables = [PSCustomObject]@{ PSPath = $path } | Get-ExcelTable

        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected properties' {
        Initialize-ExcelFixture -Path $path

        $table = Get-ExcelTable -LiteralPath $path | Where-Object -Property Name -EQ 'Employees$' | Select-Object -First 1

        $table | Should-NotBeNull
        $table.Path | Should-BeString $path
        $table.Type | Should-BeString 'TABLE'
        $table.Name | Should-BeString 'Employees$'
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        $path = Get-TempFile -Extension '.xlsm'
        Initialize-ExcelFixture -Path $path

        $tables = Get-ExcelTable -LiteralPath $path

        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
      It 'gets tables with ReadOnly' {
        Initialize-ExcelFixture -Path $path

        $tables = Get-ExcelTable -LiteralPath $path -ReadOnly

        ($tables | Where-Object -Property Name -EQ 'Employees$') | Should-NotBeNull
      }
    }
    Context 'Edge cases' {
      It 'throws for unsupported extensions' {
        $path = Get-TempFile -Extension '.txt'
        New-Item -Path $path -ItemType File -Force | Out-Null

        { Get-ExcelTable -LiteralPath $path } | Should-Throw
      }
    }
  }
  Describe 'Get-ExcelTableColumn' {
    BeforeEach {
      $path = Get-TempFile -Extension '.xlsx'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets table columns by Path' {
        Initialize-ExcelFixture -Path $path

        $columns = Get-ExcelTableColumn -Path $path -Table 'Employees$'

        $columns | Should-NotBeNull
      }
      It 'gets table columns by LiteralPath' {
        Initialize-ExcelFixture -Path $path

        $columns = Get-ExcelTableColumn -LiteralPath $path -Table 'Employees$'

        $columns | Should-NotBeNull
      }
      It 'gets table columns by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $path

        $columns = $path | Get-ExcelTableColumn -Table 'Employees$'

        $columns | Should-NotBeNull
      }
      It 'gets table columns by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $path

        $columns = [PSCustomObject]@{ PSPath = $path } | Get-ExcelTableColumn -Table 'Employees$'

        $columns | Should-NotBeNull
      }
    }
    Context 'Output' {
      It 'returns expected column metadata' {
        Initialize-ExcelFixture -Path $path

        $columns = Get-ExcelTableColumn -LiteralPath $path -Table 'Employees$'

        $id = $columns | Where-Object -Property ColumnName -EQ 'Id' | Select-Object -First 1
        $id | Should-NotBeNull
        $id.Path | Should-BeString $path
        $id.ObjectType | Should-BeString 'Table'
        $id.ObjectName | Should-BeString 'Employees$'
        $id.ColumnName | Should-BeString 'Id'
        $id.Ordinal | Should-HaveType ([int])
        $id.DataType | Should-NotBeNull
        $name = $columns | Where-Object -Property ColumnName -EQ 'Name' | Select-Object -First 1
        $name | Should-NotBeNull
        $name.Path | Should-BeString $path
        $name.ObjectType | Should-BeString 'Table'
        $name.ObjectName | Should-BeString 'Employees$'
        $name.ColumnName | Should-BeString 'Name'
        $name.Ordinal | Should-HaveType ([int])
        $name.DataType | Should-NotBeNull
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        $path = Get-TempFile -Extension '.xlsm'
        Initialize-ExcelFixture -Path $path

        $columns = Get-ExcelTableColumn -LiteralPath $path -Table 'Employees$'

        $columns | Should-NotBeNull
      }
      It 'gets table columns with ReadOnly' {
        Initialize-ExcelFixture -Path $path

        $columns = Get-ExcelTableColumn -LiteralPath $path -Table 'Employees$' -ReadOnly

        $columns | Should-NotBeNull
      }
    }
    Context 'Edge cases' {
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $path

        { Get-ExcelTableColumn -LiteralPath $path -Table 'Dummy' } | Should-Throw
      }
      It 'throws when protected with PasswordToOpen' {
        Initialize-ExcelFixture -Path $path -PasswordToOpen (Get-Password)

        { Get-ExcelTableColumn -LiteralPath $path -Table 'Employees$' } | Should-Throw
      }
    }
  }
  Describe 'Invoke-ExcelQuery' {
    BeforeEach {
      $path = Get-TempFile -Extension '.xlsx'
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'gets rows by Path' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -Path $path -Table 'Employees$'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees$'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets rows by LiteralPath' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees$'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets rows by ValueFromPipeline' {
        Initialize-ExcelFixture -Path $path

        $rows = $path | Invoke-ExcelQuery -Table 'Employees$'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees$'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets rows by ValueFromPipelineByPropertyName' {
        Initialize-ExcelFixture -Path $path

        $rows = [PSCustomObject]@{ PSPath = $path } | Invoke-ExcelQuery -Table 'Employees$'

        $rows | Should-BeCollection -Count 2
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Table'
        $rows[0].ObjectName | Should-BeString 'Employees$'
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Sales'
      }
      It 'gets rows with selected columns' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -Columns Id, Name

        $rows | Should-BeCollection -Count 2
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Id'
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Name'
        $rows[0].PSObject.Properties.Name | Should-NotContainCollection 'Department'
      }
      It 'gets rows by Query' {
        Initialize-ExcelFixture -Path $path

        $select = 'SELECT [Id], [Name] FROM [Employees$] WHERE [Id] = 1'
        $rows = Invoke-ExcelQuery -LiteralPath $path -Query $select

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
        $insert = "INSERT INTO [Employees$] (Id, Name, Department) VALUES (3, 'Carol', 'Marketing')"
        $select = 'SELECT * FROM [Employees$] WHERE [Id] = 3'
        Initialize-ExcelFixture -Path $path

        $result = Invoke-ExcelQuery -LiteralPath $path -Query $insert
        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-BeString $insert
        $result.RecordsAffected | Should-Be 1
        $rows = Invoke-ExcelQuery -LiteralPath $path -Query $select
        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Query'
        $rows[0].ObjectName | Should-BeString $select
        $rows[0].Id | Should-Be 3
        $rows[0].Name | Should-BeString 'Carol'
        $rows[0].Department | Should-BeString 'Marketing'
      }
      It 'updates rows and returns affected count' {
        $update = "UPDATE [Employees$] SET [Department] = 'Marketing' WHERE [Id] = 1"
        $select = 'SELECT * FROM [Employees$] WHERE [Id] = 1'
        Initialize-ExcelFixture -Path $path

        $result = Invoke-ExcelQuery -LiteralPath $path -Query $update
        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-BeString $update
        $result.RecordsAffected | Should-Be 1
        $rows = Invoke-ExcelQuery -LiteralPath $path -Query $select
        $rows | Should-BeCollection -Count 1
        $rows[0].Path | Should-BeString $path
        $rows[0].ObjectType | Should-BeString 'Query'
        $rows[0].ObjectName | Should-BeString $select
        $rows[0].Id | Should-Be 1
        $rows[0].Name | Should-BeString 'Alice'
        $rows[0].Department | Should-BeString 'Marketing'
      }
      It 'throws when executing DELETE' {
        Initialize-ExcelFixture -Path $path

        { Invoke-ExcelQuery -LiteralPath $path -Query 'DELETE FROM [Employees$] WHERE [Id] = 1' } | Should-Throw
      }
      It 'returns zero affected count when no rows match' {
        Initialize-ExcelFixture -Path $path

        $update = "UPDATE [Employees$] SET [Department] = 'Marketing' WHERE [Id] = 999"
        $result = Invoke-ExcelQuery -LiteralPath $path -Query $update

        $result | Should-NotBeNull
        $result.Path | Should-BeString $path
        $result.ObjectType | Should-BeString 'Query'
        $result.ObjectName | Should-BeString $update
        $result.RecordsAffected | Should-Be 0
      }
    }
    Context 'Output' {
      It 'returns row metadata and column values for table data' {
        Initialize-ExcelFixture -Path $path

        $row = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' | Where-Object -Property Id -EQ 1 | Select-Object -First 1

        $row | Should-NotBeNull
        $row.Path | Should-BeString $path
        $row.ObjectType | Should-BeString 'Table'
        $row.ObjectName | Should-BeString 'Employees$'
        $row.Id | Should-Be 1
        $row.Name | Should-BeString 'Alice'
        $row.Department | Should-BeString 'Sales'
      }
    }
    Context 'Other parameters' {
      It 'works with xlsm files' {
        $path = Get-TempFile -Extension '.xlsm'
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$'

        $rows | Should-BeCollection -Count 2
      }
      It 'gets rows with address range' {
        Initialize-ExcelFixture -Path $path
        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -Address 'A1:C2'

        $rows | Should-BeCollection -Count 1
      }
      It 'gets rows with address range and selected columns' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -Address 'A1:B3' -Columns Id, Name

        $rows | Should-BeCollection -Count 2
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Id'
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'Name'
        $rows[0].PSObject.Properties.Name | Should-NotContainCollection 'Department'
      }
      It 'throws when address count does not match table count' {
        Initialize-ExcelFixture -Path $path

        { Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$', 'Other$' -Address 'A1:C2' } | Should-Throw
      }
      It 'gets rows with NoHeader' {
        $path = Get-TempFile -Extension '.xlsm'
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -NoHeader

        $rows | Should-BeCollection -Count 3
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'F1'
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'F2'
        $rows[0].PSObject.Properties.Name | Should-ContainCollection 'F3'
      }
      It 'gets rows with NoIMEX' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -NoIMEX

        $rows | Should-BeCollection -Count 2
      }
      It 'gets rows with NoHeader and NoIMEX' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -NoHeader -NoIMEX

        $rows | Should-BeCollection -Count 2
      }
      It 'gets rows with ReadOnly' {
        Initialize-ExcelFixture -Path $path

        $rows = Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' -ReadOnly

        $rows | Should-BeCollection -Count 2
      }
    }
    Context 'Edge cases' {
      It 'throws when Query is used with Columns' {
        { Invoke-ExcelQuery -LiteralPath $path -Query 'SELECT * FROM Employees' -Columns Id } | Should-Throw
      }
      It 'throws for missing table' {
        Initialize-ExcelFixture -Path $path

        { Invoke-ExcelQuery -LiteralPath $path -Table 'Dummy' } | Should-Throw
      }
      It 'throws when protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen (Get-Password)

        { Invoke-ExcelQuery -LiteralPath $path -Table 'Employees$' } | Should-Throw
      }
    }
  }
}

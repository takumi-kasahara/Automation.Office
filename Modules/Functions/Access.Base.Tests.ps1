using assembly Microsoft.Office.Interop.Access
using assembly Microsoft.Office.Interop.Excel
using module .\..\Automation.Office.psd1
using namespace Microsoft.Office.Interop.Access
using namespace Microsoft.Office.Interop.Access.Dao
using namespace Microsoft.Office.Interop.Excel
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
  }
  Describe 'New-AccessFile' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'creates a file at the requested path' {
        $item = New-AccessFile -Path $path

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $item = $path | New-AccessFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $item = [PSCustomObject]@{ FullName = $path } | New-AccessFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
    }
    Context 'SupportsShouldProcess' {
      It 'overwrites an existing database when Force is specified' {
        New-Item -Path $path -ItemType File | Out-Null

        New-AccessFile -Path $path -Force -Confirm

        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.mdb'
        $item = New-AccessFile -Path $path -FileFormat acNewDatabaseFormatAccess2007

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with Password' {
        New-AccessFile -Path $path -Password $password

        Open-AccessFile -Path $path -Password $password -ActionProject {
          param (
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          $Project.FullName | Should-BeString $path
        }
      }
      It 'creates a file with RemovePersonalInformation' {
        New-AccessFile -Path $path -RemovePersonalInformation

        Open-AccessFile -Path $path -ActionProject {
          param (
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          $Project.RemovePersonalInformation | Should-BeTrue
        }
      }
      It 'creates a file with InitializeDb script block' {
        $item = New-AccessFile -Path $path -InitializeDb {
          param (
            [Microsoft.Office.Interop.Access.Dao.Database]
            $Database
          )
          $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
        }

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'creates a file with InitializeProject script block' {
        $item = New-AccessFile -Path $path -InitializeProject {
          param (
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          $Project.FullName | Should-BeString $path
        }

        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-AccessFile' -Tag 'CI', 'Mock' {
    BeforeEach {
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a file when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith { throw }

        New-AccessFile -Path $path -Force -WhatIf

        Should-NotInvoke -CommandName New-AccessObject
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws when the file exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null

        { New-AccessFile -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Open-AccessFile' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'opens a file by Path' {
        New-AccessFile -Path $path

        Open-AccessFile -Path $path | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipeline' {
        New-AccessFile -Path $path

        $path | Open-AccessFile | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path

        [PSCustomObject]@{ FullName = $path } | Open-AccessFile | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a file with Password' {
        New-AccessFile -Path $path -Password $password

        Open-AccessFile -Path $path -Password $password | Should-BeNull
      }
      It 'opens a file and executes ActionDb script block' {
        New-AccessFile -Path $path
        Open-AccessFile -Path $path -ActionDb {
          param (
            [Microsoft.Office.Interop.Access.Dao.Database]
            $Database
          )
          $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
        }

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'opens a file and executes ActionProject script block' {
        New-AccessFile -Path $path

        Open-AccessFile -Path $path -ActionProject {
          param (
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          return $Project.Connection
        } | Should-NotBeNull
      }
      It 'opens a file using an existing Application object' {
        New-AccessFile -Path $path

        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path | Should-BeNull
        } finally {
          $app.CloseCurrentDataBase()
          $app.Quit()
          Get-Variable |
          Where-Object -Property Value -Is [__ComObject] |
          Clear-Variable -Force -WhatIf:$false -Confirm:$false
          [GC]::Collect()
          [GC]::WaitForPendingFinalizers()
        }
      }
    }
  }
  Describe 'Get-AccessFileProperty' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-AccessFile -Path $path

        $properties = Get-AccessFileProperty -Path $path -Name RemovePersonalInformation, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.RemovePersonalInformation | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by LiteralPath' {
        New-AccessFile -Path $path

        $properties = Get-AccessFileProperty -LiteralPath $path -Name RemovePersonalInformation, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.RemovePersonalInformation | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-AccessFile -Path $path

        $properties = $path | Get-AccessFileProperty -Name RemovePersonalInformation, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.RemovePersonalInformation | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path

        $properties = [PSCustomObject]@{ PSPath = $path } | Get-AccessFileProperty -Name RemovePersonalInformation, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.RemovePersonalInformation | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with Password' {
        New-AccessFile -Path $path -Password $password

        { Get-AccessFileProperty -LiteralPath $path -Name FullName } | Should-Throw
        { Get-AccessFileProperty -LiteralPath $path -Name FullName -Password (Get-Password) } | Should-Throw
        { Get-AccessFileProperty -LiteralPath $path -Name FullName -Password $password } | Should -Not -Throw
      }
      It 'returns the property directly when only one name is provided' {
        New-AccessFile -Path $path

        Get-AccessFileProperty -LiteralPath $path -Name RemovePersonalInformation | Should-BeFalse
      }
      It 'returns file properties with Depth' {
        New-AccessFile -Path $path

        $properties = Get-AccessFileProperty -Path $path -Name Properties -Depth 2

        $properties.Count | Should-Be 0
      }
    }
  }
  Describe 'Get-AccessFileProperty' -Tag 'CI', 'Mock' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Set-AccessFileProperty' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $name = 'RemovePersonalInformation'
      $value = $true
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'updates a file property by Path' {
        New-AccessFile -Path $path

        Set-AccessFileProperty -Path $path -Name $name -Value $value

        Get-AccessFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath' {
        New-AccessFile -Path $path

        Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value
        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeTrue

        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with InputObject' {
        New-AccessFile -Path $path

        $properties = [PSCustomObject]@{ RemovePersonalInformation = $value }
        Set-AccessFileProperty -LiteralPath $path -InputObject $properties

        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with InputObject' {
        New-AccessFile -Path $path

        $properties = [PSCustomObject]@{ RemovePersonalInformation = $value }
        Set-AccessFileProperty -LiteralPath $path -InputObject $properties

        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-AccessFile -Path $path

        $path | Set-AccessFileProperty -Name $name -Value $value

        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path

        [PSCustomObject]@{ PSPath = $path } | Set-AccessFileProperty -Name $name -Value $value

        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update property when WhatIf is specified' {
        New-AccessFile -Path $path

        Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Get-AccessFileProperty -LiteralPath $path -Name $name | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with Password' {
        New-AccessFile -Path $path -Password $password

        { Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value -Password (Get-Password) } | Should-Throw
        { Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value -Password $password } | Should -Not -Throw
        Get-AccessFileProperty -LiteralPath $path -Name $name -Password $password | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-AccessFile -Path $path

        $property = Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value -PassThru

        $property | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-AccessFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        $property = Set-AccessFileProperty -LiteralPath $path -InputObject $properties -PassThru

        $property | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-AccessFile -Path $path

        { Set-AccessFileProperty -LiteralPath $path -Name 'Dummy' -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-AccessFileProperty' -Tag 'CI', 'Mock' {
    BeforeEach {
      $path = Get-TempFile
      $name = 'RemovePersonalInformation'
      $value = $true
      New-Item -Path $path -ItemType File -Force | Out-Null

      $app = [PSCustomObject]@{}
      $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
      $app | Add-Member -MemberType ScriptMethod -Name CloseCurrentDatabase -Value { }

      $database = [PSCustomObject]@{}
      $database | Add-Member -MemberType ScriptMethod -Name Close -Value { }

      Mock -CommandName New-AccessObject -MockWith { $app }
      Mock -CommandName Open-AccessFile -MockWith { $null }
      Mock -CommandName Set-ObjectProperty
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not open the file when WhatIf is specified' {
        Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Should-NotInvoke -CommandName Open-AccessFile
      }
    }
    Context 'Edge cases' {
      It 'throws when the target is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        Should-NotInvoke -CommandName Open-AccessFile
      }
    }
  }
  Describe 'Export-AccessDatabase' {
    BeforeAll {
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
        }
      }
      function Get-Destination {
        [CmdletBinding()]
        [OutputType([string])]
        param (
          [string]
          $Extension = '.txt'
        )
        return $env:TEMP | Join-Path -ChildPath "Export.$([guid]::NewGuid().ToString('N'))$Extension"
      }
    }
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $destination = Get-Destination
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Force
      }
    }
    Context 'ParameterSetName' {
      It 'exports data to text file with default TextSet' {
        Initialize-AccessFixture -Path $path

        $item = Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports data by Path with ValueFromPipeline' {
        Initialize-AccessFixture -Path $path

        $item = $path | Export-AccessDatabase -TableName 'Employees' -Destination $destination

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports data by Path with ValueFromPipelineByPropertyName' {
        Initialize-AccessFixture -Path $path

        $item = [PSCustomObject]@{ FullName = $path } | Export-AccessDatabase -TableName 'Employees' -Destination $destination

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not export when WhatIf is specified' {
        Initialize-AccessFixture -Path $path

        Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -WhatIf

        Test-Path -LiteralPath $destination | Should-BeFalse
      }
      It 'overwrites an existing read-only destination when Force is specified' {
        Initialize-AccessFixture -Path $path
        New-Item -Path $destination -ItemType File -Force | Out-Null
        $readOnlyItem = Get-Item -LiteralPath $destination -Force
        $readOnlyItem.IsReadOnly = $true

        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination } | Should-Throw
        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -Confirm -Force } | Should -Not -Throw
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'throws and keeps the existing destination content unchanged when NoClobber is specified' {
        Initialize-AccessFixture -Path $path
        New-Item -Path $destination -ItemType File | Out-Null

        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -NoClobber } | Should-Throw
        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -Force -NoClobber } | Should-Throw
        (Get-Item -LiteralPath $destination).Length | Should-Be 0
      }
    }
    Context 'Other parameters' {
      It 'exports data from a file protected with Password' {
        Initialize-AccessFixture -Path $path -Password $password

        $item = Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -Password $password

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'export data with TransferType' {
        Initialize-AccessFixture -Path $path

        $item = Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -TransferType acExportHTML

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports data with CodePage' {
        Initialize-AccessFixture -Path $path

        $item = Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -CodePage 65001

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports data as SpreadsheetType' {
        Initialize-AccessFixture -Path $path

        $destination = Get-Destination -Extension '.xlsx'
        $item = Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -SpreadsheetType acSpreadsheetTypeExcel12Xml

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports data as SpreadsheetType with Range' {
        Initialize-AccessFixture -Path $path

        $destination = Get-Destination -Extension '.xlsx'
        $item = Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -SpreadsheetType acSpreadsheetTypeExcel12Xml -Range 'A1:C3'

        $item | Should-HaveType ([System.IO.FileInfo])
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Export-AccessDatabase' -Tag 'CI', 'Mock' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null

      $destination = $env:TEMP | Join-Path -ChildPath "Export.$([guid]::NewGuid().ToString('N')).txt"

      $app = [PSCustomObject]@{}
      $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
      $app | Add-Member -MemberType ScriptMethod -Name CloseCurrentDataBase -Value { }

      Mock -CommandName New-AccessObject -MockWith { $app }
      Mock -CommandName Open-AccessFile -MockWith { $null }
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a file when WhatIf is specified' {
        Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -WhatIf
        Should-NotInvoke -CommandName New-AccessObject
        Test-Path -LiteralPath $destination | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws when destination exists and Force is not specified' {
        New-Item -Path $destination -ItemType File -Force | Out-Null

        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination } | Should-Throw
      }
      It 'throws when destination is an existing directory' {
        New-Item -Path $destination -ItemType Directory | Out-Null

        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination } | Should-Throw
      }
    }
  }
  Describe 'Import-AccessDatabase' {
    BeforeAll {
      function Get-SourceFile {
        [CmdletBinding()]
        [OutputType([string])]
        param (
          [string]
          $Extension = '.txt'
        )
        return $env:TEMP | Join-Path -ChildPath "Source.$([guid]::NewGuid().ToString('N'))$Extension"
      }
      function Initialize-TextFile {
        [CmdletBinding()]
        [OutputType([string])]
        param (
          [string]
          $Path,
          [string]
          $Content = @(
            '"Id","Name","Department"'
            '1,"Alice","Sales"'
            '2,"Bob","Engineering"'
          ) -join [Environment]::NewLine,
          [ValidateSet(1200, 65001)]
          [int]
          $CodePage = 1200
        )
        switch ($CodePage) {
          1200 {
            $Content | Out-File -LiteralPath $Path -Encoding unicode
          }
          65001 {
            $Content | Out-File -LiteralPath $Path -Encoding utf8
          }
        }
        return $Path
      }
      function Initialize-ExcelFile {
        [CmdletBinding()]
        [OutputType([string])]
        param (
          [string]
          $Path,
          [SecureString]
          $PasswordToOpen,
          [SecureString]
          $PasswordToModify
        )
        New-ExcelFile -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Initialize {
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
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $source = Get-SourceFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $source) {
        Remove-Item -LiteralPath $source -Force
      }
    }
    Context 'ParameterSetName' {
      It 'imports data from text file with default TextSet' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees'

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'imports data by Path with ValueFromPipeline' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source

        $path | Import-AccessDatabase -Source $source -TableName 'Employees'

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'imports data by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source

        [PSCustomObject]@{ FullName = $path } | Import-AccessDatabase -Source $source -TableName 'Employees'

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not import when WhatIf is specified' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -WhatIf

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'imports data from a file protected with Password' {
        New-AccessFile -Path $path -Password $password
        Initialize-TextFile -Path $source

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -Password $password

        Get-AccessTable -Path $path -Password $password | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'imports data with TransferType' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -TransferType acImportDelim

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'imports data with CodePage' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source -CodePage 65001

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -CodePage 65001

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'imports data as SpreadsheetType' {
        New-AccessFile -Path $path
        $source = Get-SourceFile -Extension '.xlsx'
        Initialize-ExcelFile -Path $source -PasswordToModify (Get-Password)

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
      It 'imports data as SpreadsheetType with Range' {
        New-AccessFile -Path $path
        $source = Get-SourceFile -Extension '.xlsx'
        Initialize-ExcelFile -Path $source -PasswordToModify (Get-Password)

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml -Range 'A1:C3'

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-NotBeNull
      }
    }
    Context 'Edge cases' {
      It 'throws when the source file does not exist' {
        New-AccessFile -Path $path

        { Import-AccessDatabase -Path $path -Source (Get-TempFile) -TableName 'Employees' } | Should-Throw
      }
      It 'throws when the database is read-only' {
        New-AccessFile -Path $path
        $readOnlyItem = Get-Item -LiteralPath $path -Force
        $readOnlyItem.IsReadOnly = $true

        { Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' } | Should-Throw
      }
      It 'throws when TableName does not exist in source file' {
        New-AccessFile -Path $path

        { Import-AccessDatabase -Path $path -Source $source -TableName 'Dummy' } | Should-Throw
      }
      It 'does not import data when the file is protected with PasswordToOpen' {
        New-AccessFile -Path $path
        $source = Get-SourceFile -Extension '.xlsx'
        Initialize-ExcelFile -Path $source -PasswordToOpen (Get-Password)

        Get-AccessTable -Path $path | Where-Object -Property Name -EQ 'Employees' | Should-BeNull
      }
    }
  }
  Describe 'Import-AccessDatabase' -Tag 'CI', 'Mock' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null

      $source = $env:TEMP | Join-Path -ChildPath "Source.$([guid]::NewGuid().ToString('N')).txt"
      New-Item -Path $source -ItemType File -Force | Out-Null

      $app = [PSCustomObject]@{}
      $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
      $app | Add-Member -MemberType ScriptMethod -Name CloseCurrentDataBase -Value { }

      Mock -CommandName New-AccessObject -MockWith { $app }
      Mock -CommandName Open-AccessFile -MockWith { $null }
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $source) {
        Remove-Item -LiteralPath $source -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call the program when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith { throw }

        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -WhatIf
        Should-NotInvoke -CommandName New-AccessObject
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call the program when the target is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' } | Should-Throw
        Should-NotInvoke -CommandName New-AccessObject
      }
    }
  }
}

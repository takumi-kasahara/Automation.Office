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
    function Get-AccessPassword {
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
    BeforeAll {
      function Invoke-Confirm {
        [CmdletBinding()]
        [OutputType([int])]
        param (
          [Parameter(ValueFromPipeline)]
          [string]
          $Response,
          [Parameter(Mandatory)]
          [string]
          $Path
        )
        process {
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $command = @(
            "Import-Module -Name '$escapedModulePath' -Force"
            "New-AccessFile -Path '$escapedPath' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $path = Get-TempFile
      $password = Get-AccessPassword
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'creates a file at the requested path' {
        $item = New-AccessFile -Path $path
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $path = Get-TempFile
        $item = $path | New-AccessFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $path = Get-TempFile
        $item = [PSCustomObject]@{ FullName = $path } | New-AccessFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a file when WhatIf is specified' {
        $path = Get-TempFile
        New-AccessFile -Path $path -Force -WhatIf
        Test-Path -LiteralPath $path | Should-BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        $path = Get-TempFile
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should-Be 0
        Test-Path -LiteralPath $path | Should-BeFalse
      }
      It 'overwrites an existing database when Force is specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        New-AccessFile -Path $path -Force -Confirm
        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.mdb'
        $item = New-AccessFile -Path $path -FileFormat acNewDatabaseFormatAccess2007
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a database with Password' {
        $path = Get-TempFile
        New-AccessFile -Path $path -Password $password
        Test-Path -LiteralPath $path | Should-BeTrue
        Open-AccessFile -Path $path -Password $password -ActionProject {
          param(
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          $Project.FullName | Should-Be $path
        }
      }
      It 'creates a database with RemovePersonalInformation' {
        $path = Get-TempFile
        New-AccessFile -Path $path -RemovePersonalInformation
        Test-Path -LiteralPath $path | Should-BeTrue
        Open-AccessFile -Path $path -ActionProject {
          param(
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          $Project.RemovePersonalInformation | Should-BeTrue
        }
      }
      It 'creates a database with InitializeDb script block' {
        $path = Get-TempFile
        $item = New-AccessFile -Path $path -InitializeDb {
          param(
            [Microsoft.Office.Interop.Access.Dao.Database]
            $Database
          )
          $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
        }
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'creates a database with InitializeProject script block' {
        $path = Get-TempFile
        $item = New-AccessFile -Path $path -InitializeProject {
          param(
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          $Project.FullName | Should-Be $path
        }
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'fails when the path already exists and Force is not specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        { New-AccessFile -Path $path } | Should-Throw
      }
    }
  }
  Describe 'New-AccessFile.Unit' {
    BeforeEach {
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-AccessObject when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith { throw }

        { New-AccessFile -Path $path -Force -WhatIf } | Should -Not -Throw
        Should-NotInvoke -CommandName New-AccessObject
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-AccessObject when path exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null
        Mock -CommandName New-AccessObject

        { New-AccessFile -Path $path } | Should-Throw
        Should-NotInvoke -CommandName New-AccessObject
      }
    }
  }
  Describe 'Open-AccessFile' {
    BeforeEach {
      $path = Get-TempFile
      $password = Get-AccessPassword
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'opens a database by Path' {
        New-AccessFile -Path $path
        Open-AccessFile -Path $path | Should-BeNull
      }
      It 'opens a database by Path with ValueFromPipeline' {
        New-AccessFile -Path $path
        $path | Open-AccessFile | Should-BeNull
      }
      It 'opens a database by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path
        [PSCustomObject]@{ FullName = $path } | Open-AccessFile | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a database with Password' {
        New-AccessFile -Path $path -Password $password
        Open-AccessFile -Path $path -Password $password | Should-BeNull
      }
      It 'opens a database and executes ActionDb script block' {
        New-AccessFile -Path $path
        Open-AccessFile -Path $path -ActionDb {
          param(
            [Microsoft.Office.Interop.Access.Dao.Database]
            $Database
          )
          $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
        }
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'opens a database and executes ActionProject script block' {
        New-AccessFile -Path $path
        Open-AccessFile -Path $path -ActionProject {
          param(
            [Microsoft.Office.Interop.Access.CurrentProject]
            $Project
          )
          return $Project.Connection
        } | Should-NotBeNull
      }
      It 'opens a database using an existing Application object' {
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
      $path = Get-TempFile
      $password = Get-AccessPassword
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-AccessFile -Path $path
        $properties = Get-AccessFileProperty -Path $path -Name RemovePersonalInformation
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.RemovePersonalInformation | Should-BeFalse
      }
      It 'returns selected file properties by LiteralPath' {
        New-AccessFile -Path $path
        $properties = Get-AccessFileProperty -LiteralPath $path -Name RemovePersonalInformation
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.RemovePersonalInformation | Should-BeFalse
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-AccessFile -Path $path
        $properties = $path | Get-AccessFileProperty -Name RemovePersonalInformation
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.RemovePersonalInformation | Should-BeFalse
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-AccessFileProperty -Name RemovePersonalInformation
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.RemovePersonalInformation | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with Password' {
        New-AccessFile -Path $path -Password $password
        { Get-AccessFileProperty -Path $path } | Should-Throw
        { Get-AccessFileProperty -Path $path -Password (Get-AccessPassword) } | Should-Throw
        { Get-AccessFileProperty -Path $path -Password $password } | Should -Not -Throw
      }
    }
  }
  Describe 'Get-AccessFileProperty.Unit' {
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
      It 'throws when New-AccessObject fails' {
        Mock -CommandName New-AccessObject -MockWith { throw }

        { Get-AccessFileProperty -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Set-AccessFileProperty' {
    BeforeEach {
      $path = Get-TempFile
      $password = Get-AccessPassword
      $name = 'RemovePersonalInformation'
      $value = $true
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'updates a file property by by Path' {
        New-AccessFile -Path $path
        Set-AccessFileProperty -Path $path -Name $name -Value $value
        (Get-AccessFileProperty -Path $path).RemovePersonalInformation | Should-BeTrue
      }
      It 'updates a file property by by LiteralPath' {
        New-AccessFile -Path $path
        Set-AccessFileProperty -LiteralPath $path -Name $name -Value $value
        (Get-AccessFileProperty -LiteralPath $path).RemovePersonalInformation | Should-BeTrue
      }
      It 'updates a file properties by Path with InputObject' {
        New-AccessFile -Path $path
        $properties = [PSCustomObject]@{ RemovePersonalInformation = $value }
        Set-AccessFileProperty -Path $path -InputObject $properties
        (Get-AccessFileProperty -Path $path).RemovePersonalInformation | Should-BeTrue
      }
      It 'updates a file properties by LiteralPath with InputObject' {
        New-AccessFile -Path $path
        $properties = [PSCustomObject]@{ RemovePersonalInformation = $value }
        Set-AccessFileProperty -LiteralPath $path -InputObject $properties
        (Get-AccessFileProperty -LiteralPath $path).RemovePersonalInformation | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-AccessFile -Path $path
        $path | Set-AccessFileProperty -Name $name -Value $value
        (Get-AccessFileProperty -Path $path).RemovePersonalInformation | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-AccessFileProperty -Name $name -Value $value
        (Get-AccessFileProperty -Path $path).RemovePersonalInformation | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update property when WhatIf is specified' {
        New-AccessFile -Path $path
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $value -WhatIf
        (Get-AccessFileProperty -Path $path).RemovePersonalInformation | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with Password' {
        New-AccessFile -Path $path -Password $password
        { Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true } | Should-Throw
        { Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -Password (Get-AccessPassword) } | Should-Throw
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -Password $password
        (Get-AccessFileProperty -Path $path -Password $password).RemovePersonalInformation | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-AccessFile -Path $path
        $result = Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -PassThru
        $result | Should-NotBeNull
        $result.RemovePersonalInformation | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-AccessFile -Path $path
        { Set-AccessFileProperty -Path $path -Name 'NonExistentProperty' -Value 'Value' } | Should-Throw
      }
    }
  }
  Describe 'Set-AccessFileProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
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
      It 'does not call Open-AccessFile when WhatIf is specified' {
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -WhatIf
        Should-NotInvoke -CommandName Open-AccessFile
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        { Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true } | Should-Throw
      }
    }
  }
  Describe 'Export-AccessDatabase' {
    BeforeAll {
      function Invoke-Confirm {
        [CmdletBinding()]
        [OutputType([int])]
        param (
          [Parameter(ValueFromPipeline)]
          [string]
          $Response,
          [Parameter(Mandatory)]
          [string]
          $Path,
          [string]
          $Destination
        )
        process {
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $escapedDestination = $Destination.Replace("'", "''")
          $command = @(
            "Import-Module -Name '$escapedModulePath' -Force"
            "Export-AccessDatabase -Path '$escapedPath' -TableName 'Employees' -Destination '$escapedDestination' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
      function Initialize-AccessFixture {
        param(
          [string]
          $Path,
          [SecureString]
          $Password = $null
        )
        New-AccessFile -Path $Path -Password $Password -Force -InitializeDb {
          param(
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
      $path = Get-TempFile
      $password = Get-AccessPassword
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
        $item.FullName | Should-Be ([Path]::GetFullPath($destination))
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
      It 'asks for confirmation when Confirm is specified' {
        Initialize-AccessFixture -Path $path
        $exitCode = 'N' | Invoke-Confirm -Path $path -Destination $destination
        $exitCode | Should-Be 0
        Test-Path -LiteralPath $destination | Should-BeFalse
      }
      It 'overwrites an existing read-only destination when Force is specified' {
        Initialize-AccessFixture -Path $path
        New-Item -Path $destination -ItemType File -Force | Out-Null
        $readOnlyItem = Get-Item -LiteralPath $destination -Force
        $readOnlyItem.IsReadOnly = $true
        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination } | Should-Throw
        Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -Force | Out-Null
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
      It 'exports data from a database protected with Password' {
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
      It 'throws when the destination is an existing directory' {
        Initialize-AccessFixture -Path $path
        New-Item -Path $destination -ItemType Directory | Out-Null
        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination } | Should-Throw
      }
    }
  }
  Describe 'Export-AccessDatabase.Unit' {
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
      It 'does not call New-AccessObject when WhatIf is specified' {
        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination -WhatIf } | Should -Not -Throw
        Should-NotInvoke -CommandName New-AccessObject
        Test-Path -LiteralPath $destination | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-AccessObject when destination exists and Force is not specified' {
        New-Item -Path $destination -ItemType File -Force | Out-Null
        Mock -CommandName New-AccessObject

        { Export-AccessDatabase -Path $path -TableName 'Employees' -Destination $destination } | Should-Throw
        Should-NotInvoke -CommandName New-AccessObject
      }
    }
  }
  Describe 'Import-AccessDatabase' {
    BeforeAll {
      function Invoke-Confirm {
        [CmdletBinding()]
        [OutputType([int])]
        param (
          [Parameter(ValueFromPipeline)]
          [string]
          $Response,
          [Parameter(Mandatory)]
          [string]
          $Path,
          [string]
          $Source
        )
        process {
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $escapedSource = $Source.Replace("'", "''")
          $command = @(
            "Import-Module -Name '$escapedModulePath' -Force"
            "Import-AccessDatabase -Path '$escapedPath' -Source '$escapedSource' -TableName 'Employees' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
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
          1200 { $Content | Out-File -LiteralPath $Path -Encoding unicode }
          65001 { $Content | Out-File -LiteralPath $Path -Encoding utf8 }
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
    BeforeEach {
      $path = Get-TempFile
      $password = Get-AccessPassword
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
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'imports data by Path with ValueFromPipeline' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source
        $path | Import-AccessDatabase -Source $source -TableName 'Employees'
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'imports data by Path with ValueFromPipelineByPropertyName' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source
        [PSCustomObject]@{ FullName = $path } | Import-AccessDatabase -Source $source -TableName 'Employees'
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not import when WhatIf is specified' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source
        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -WhatIf
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-BeNull
      }
      It 'asks for confirmation when Confirm is specified' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source
        $exitCode = 'N' | Invoke-Confirm -Path $path -Source $source
        $exitCode | Should-Be 0
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'imports data from a database protected with Password' {
        New-AccessFile -Path $path -Password $password
        Initialize-TextFile -Path $source
        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -Password $password
        $tables = Get-AccessTable -Path $path -Password $password
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'imports data with TransferType' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source
        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -TransferType acImportDelim
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'imports data with CodePage' {
        New-AccessFile -Path $path
        Initialize-TextFile -Path $source -CodePage 65001
        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -CodePage 65001
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
      It 'imports data as SpreadsheetType' {
        New-AccessFile -Path $path
        $source = Get-SourceFile -Extension '.xlsx'
        Initialize-ExcelFile -Path $source -PasswordToModify (Get-Password)
        Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
      }
    }
    Context 'Edge cases' {
      It 'throws when the source file does not exist' {
        New-AccessFile -Path $path
        $invalidSource = Get-TempFile -Extension '.txt'
        { Import-AccessDatabase -Path $path -Source $invalidSource -TableName 'Employees' } | Should-Throw
      }
      It 'throws when the database is read-only' {
        New-AccessFile -Path $path
        $readOnlyItem = Get-Item -LiteralPath $path -Force
        $readOnlyItem.IsReadOnly = $true
        { Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' } | Should-Throw
      }
      It 'throws when TableName does not exist in source file' {
        New-AccessFile -Path $path
        { Import-AccessDatabase -Path $path -Source $source -TableName 'NonExistentTable' } | Should-Throw
      }
      It 'does not import data when workbook is protected with PasswordToOpen' {
        New-AccessFile -Path $path
        $source = Get-SourceFile -Extension '.xlsx'
        Initialize-ExcelFile -Path $source -PasswordToOpen (Get-Password)
        $tables = Get-AccessTable -Path $path
        ($tables | Where-Object -Property Name -EQ 'Employees') | Should-BeNull
      }
    }
  }
  Describe 'Import-AccessDatabase.Unit' {
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
      It 'does not call New-AccessObject when WhatIf is specified' {
        { Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' -WhatIf } | Should -Not -Throw
        Should-NotInvoke -CommandName New-AccessObject
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-AccessObject when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        Mock -CommandName New-AccessObject

        { Import-AccessDatabase -Path $path -Source $source -TableName 'Employees' } | Should-Throw
        Should-NotInvoke -CommandName New-AccessObject
      }
    }
  }
}

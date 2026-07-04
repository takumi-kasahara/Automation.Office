using namespace Microsoft.Office.Interop.Excel
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Security

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest

InModuleScope 'Excel.Base' {
  BeforeAll {
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
  }
  Describe 'New-ExcelFile' {
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
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $command = @(
            "Import-Module '$escapedModulePath' -Force"
            "New-ExcelFile -Path '$escapedPath' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'creates a file at the requested path' {
        $path = Get-TempFile
        $item = New-ExcelFile -Path $path
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $path = Get-TempFile
        $item = $path | New-ExcelFile
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $path = Get-TempFile
        $item = [PSCustomObject]@{ FullName = $path } | New-ExcelFile
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a file when WhatIf is specified' {
        $path = Get-TempFile
        New-ExcelFile -Path $path -Force -WhatIf
        Test-Path -LiteralPath $path | Should -BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        $path = Get-TempFile
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should -Be 0
        Test-Path -LiteralPath $path | Should -BeFalse
      }
      It 'overwrites an existing workbook when Force is specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        New-ExcelFile -Path $path -Force -Confirm
        (Get-Item -LiteralPath $path).Length | Should -BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.xlsm'
        $item = New-ExcelFile -Path $path -FileFormat xlOpenXMLWorkbookMacroEnabled
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
      It 'creates a file with PasswordToOpen' {
        $path = Get-TempFile
        New-ExcelFile -Path $path -PasswordToOpen $password
        Test-Path -LiteralPath $path | Should -BeTrue
        $app = New-ExcelObject
        Open-ExcelFile -Path $path -PasswordToOpen $password -Action {
          param (
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.HasPassword | Should -BeTrue
          $Workbook.ReadOnly | Should -BeFalse
        }
      }
      It 'creates a file with PasswordToModify' {
        $path = Get-TempFile
        New-ExcelFile -Path $path -PasswordToModify $password
        Test-Path -LiteralPath $path | Should -BeTrue
        Open-ExcelFile -Path $path -PasswordToModify $password -Action {
          param (
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.HasPassword | Should -BeTrue
          $Workbook.ReadOnly | Should -BeFalse
        }
      }
      It 'creates a file with ReadOnlyRecommended' {
        $path = Get-TempFile
        New-ExcelFile -Path $path -ReadOnlyRecommended
        Test-Path -LiteralPath $path | Should -BeTrue
        Open-ExcelFile -Path $path -Action {
          param (
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.ReadOnlyRecommended | Should -BeTrue
        }
      }
      It 'creates a file with RemovePersonalInformation' {
        $path = Get-TempFile
        New-ExcelFile -Path $path -RemovePersonalInformation
        Test-Path -LiteralPath $path | Should -BeTrue
        Open-ExcelFile -Path $path -Action {
          param (
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.RemovePersonalInformation | Should -BeTrue
        }
      }
      It 'opens a file with Action script block' {
        $path = Get-TempFile
        New-ExcelFile -Path $path
        Open-ExcelFile -Path $path -Action {
          param (
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.Worksheets.Item(1).Name = 'TestData'
        }
        Open-ExcelFile -Path $path -Action {
          param (
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.Worksheets.Item(1).Name | Should -Be 'TestData'
        }
      }
    }
    It 'creates a file with Initialize script block' {
      $path = Get-TempFile
      $item = New-ExcelFile -Path $path -Initialize {
        param (
          [Parameter(Mandatory)]
          [Workbook]
          $Workbook
        )
        $Workbook.Worksheets.Item(1).Name = 'TestData'
      }
      $item | Should -BeOfType [System.IO.FileInfo]
      $item.FullName | Should -Be ([Path]::GetFullPath($path))
      Test-Path -LiteralPath $path | Should -BeTrue
      Open-ExcelFile -Path $path -Action {
        param (
          [Parameter(Mandatory)]
          [Workbook]
          $Workbook
        )
        $Workbook.Worksheets.Item(1).Name | Should -Be 'TestData'
      }
    }
    Context 'Edge cases' {
      It 'fails when the path already exists and Force is not specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        { New-ExcelFile -Path $path } | Should -Throw
      }
    }
  }
  Describe 'New-ExcelFile.Unit' {
    BeforeEach {
      $path = Get-TempFile
      $script:called = 0
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-ExcelObject when WhatIf is specified' {
        Mock -CommandName New-ExcelObject -MockWith {
          $script:called++
          throw 'must not be called'
        }

        { New-ExcelFile -Path $path -Force -WhatIf } | Should -Not -Throw
        $script:called | Should -Be 0
        Test-Path -LiteralPath $path | Should -BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-ExcelObject when path exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null
        Mock -CommandName New-ExcelObject -MockWith {
          $script:called++
          throw 'must not be called'
        }

        { New-ExcelFile -Path $path } | Should -Throw
        $script:called | Should -Be 0
      }
    }
  }
  Describe 'Open-ExcelFile' {
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
      It 'opens a workbook and returns the workbook object' {
        New-ExcelFile -Path $path
        { Open-ExcelFile -Path $path } | Should -Not -Throw
      }
      It 'opens a workbook by Path with ValueFromPipeline' {
        New-ExcelFile -Path $path
        { $path | Open-ExcelFile } | Should -Not -Throw
      }
      It 'opens a workbook by Path with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path
        { [PSCustomObject]@{ FullName = $path } | Open-ExcelFile } | Should -Not -Throw
      }
      It 'opens a workbook using an existing Application object' {
        New-ExcelFile -Path $path
        $app = New-ExcelObject
        try {
          { Open-ExcelFile -Application $app -Path $path } | Should -Not -Throw
        }
        finally {
          $app.Quit()
          Get-Variable |
          Where-Object -Property Value -Is [__ComObject] |
          Clear-Variable -Force -WhatIf:$false -Confirm:$false
          [GC]::Collect()
          [GC]::WaitForPendingFinalizers()
        }
      }
    }
    Context 'Other parameters' {
      It 'opens a workbook with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Open-ExcelFile -Path $path -PasswordToOpen $password } | Should -Not -Throw
      }
      It 'opens a workbook with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password
        { Open-ExcelFile -Path $path -PasswordToModify $password } | Should -Not -Throw
      }
      It 'opens a workbook with ReadOnly' {
        New-ExcelFile -Path $path
        { Open-ExcelFile -Path $path -ReadOnly } | Should -Not -Throw
      }
      It 'opens a workbook with Force' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        { Open-ExcelFile -Path $path -Force } | Should -Not -Throw
      }
      It 'opens a workbook with Action script block' {
        New-ExcelFile -Path $path
        $result = Open-ExcelFile -Path $path -Action {
          param(
            [Parameter(Mandatory)]
            [Workbook]
            $Workbook
          )
          $Workbook.Worksheets.Item(1).Name = 'Sheet1'
          return $Workbook.Worksheets.Item(1).Name
        }
        $result | Should -Be 'Sheet1'
      }
    }
  }
  Describe 'Open-ExcelFile.Unit' {
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
      It 'throws when New-ExcelObject fails and Application is not specified' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

        { Open-ExcelFile -Path $path } | Should -Throw
        Assert-MockCalled -CommandName New-ExcelObject -Times 1 -Exactly
      }
    }
  }
  Describe 'Get-ExcelAppProperty' {
    It 'return application properties' {
      { Get-ExcelAppProperty | Out-Host } | Should -Not -Throw
    }
  }
  Describe 'Get-ExcelAppProperty.Unit' {
    It 'throws when New-ExcelObject fails and requests NoSetup' {
      Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

      { Get-ExcelAppProperty } | Should -Throw
      Assert-MockCalled -CommandName New-ExcelObject -ParameterFilter { $NoSetup } -Times 1 -Exactly
    }
  }
  Describe 'Set-ExcelAppProperty' {
    It 'set application properties' {
      $properties = Get-ExcelAppProperty
      { Set-ExcelAppProperty -Properties $properties } | Should -Not -Throw
    }
    It 'throws when trying to set a property that does not exist.' {
      $properties = [PSCustomObject]@{ NonExistentProperty = 'Value' }
      { Set-ExcelAppProperty -Properties $properties } | Should -Throw
    }
    It 'does not update properties when WhatIf is specified' {
      $properties = Get-ExcelAppProperty
      { Set-ExcelAppProperty -Properties $properties -WhatIf } | Should -Not -Throw
    }
  }
  Describe 'Set-ExcelAppProperty.Unit' {
    BeforeEach {
      $properties = [PSCustomObject]@{ Visible = $false }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-ExcelObject when WhatIf is specified' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'must not be called' }

        { Set-ExcelAppProperty -Properties $properties -WhatIf } | Should -Not -Throw
        Assert-MockCalled -CommandName New-ExcelObject -Times 0 -Exactly
      }
    }
    Context 'ParameterSetName' {
      It 'throws when New-ExcelObject fails' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

        { Set-ExcelAppProperty -Properties $properties } | Should -Throw
      }
    }
  }
  Describe 'Get-ExcelFileProperty' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    It 'returns file properties' {
      New-ExcelFile -Path $path
      { Get-ExcelFileProperty -Path $path | Out-Host } | Should -Not -Throw
    }
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-ExcelFile -Path $path
        $properties = Get-ExcelFileProperty -Path $path -Name Final
        @($properties.PSObject.Properties).Count | Should -Be 1
        $properties.Final | Should -BeFalse
      }
      It 'returns selected file properties by LiteralPath' {
        New-ExcelFile -Path $path
        $properties = Get-ExcelFileProperty -LiteralPath $path -Name Final
        @($properties.PSObject.Properties).Count | Should -Be 1
        $properties.Final | Should -BeFalse
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-ExcelFile -Path $path
        $properties = $path | Get-ExcelFileProperty -Name Final
        @($properties.PSObject.Properties).Count | Should -Be 1
        $properties.Final | Should -BeFalse
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-ExcelFileProperty -Name Final
        @($properties.PSObject.Properties).Count | Should -Be 1
        $properties.Final | Should -BeFalse
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Get-ExcelFileProperty -Path $path -PasswordToOpen (Get-Password) | Out-Null } | Should -Throw
        { Get-ExcelFileProperty -Path $path -PasswordToOpen $password | Out-Null } | Should -Not -Throw
      }
      It 'returns file properties from a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password
        { Get-ExcelFileProperty -Path $path -PasswordToModify (Get-Password) | Out-Null } | Should -Not -Throw
        { Get-ExcelFileProperty -Path $path -PasswordToModify $password | Out-Null } | Should -Not -Throw
      }
    }
  }
  Describe 'Get-ExcelFileProperty.Unit' {
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
      It 'throws when New-ExcelObject fails' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

        { Get-ExcelFileProperty -Path $path } | Should -Throw
      }
    }
  }
  Describe 'Set-ExcelFileProperty' {
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
          [Parameter(Mandatory)]
          [string]
          $Name,
          [Parameter(Mandatory)]
          [bool]
          $Value
        )
        process {
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $escapedName = $Name.Replace("'", "''")
          $escapedValue = if ($Value) {
            '$true'
          }
          else {
            '$false'
          }
          $command = @(
            "Import-Module '$escapedModulePath' -Force"
            "Set-ExcelFileProperty -LiteralPath '$escapedPath' -Name '$escapedName' -Value $escapedValue -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $name = 'Final'
      $value = $true
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'updates a file property by Path with Name and Value' {
        New-ExcelFile -Path $path
        Set-ExcelFileProperty -Path $path -Name $name -Value $value
        (Get-ExcelFileProperty -Path $path).Final | Should -BeTrue
      }
      It 'updates a file property by LiteralPath with Name and Value' {
        New-ExcelFile -Path $path
        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value
        (Get-ExcelFileProperty -LiteralPath $path).Final | Should -BeTrue
      }
      It 'updates a file properties by Path with InputObject' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        Set-ExcelFileProperty -Path $path -InputObject $properties
        $actual = Get-ExcelFileProperty -Path $path
        $actual.Final | Should -BeTrue
      }
      It 'updates a file properties by LiteralPath with InputObject' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        Set-ExcelFileProperty -LiteralPath $path -InputObject $properties
        $actual = Get-ExcelFileProperty -LiteralPath $path
        $actual.Final | Should -BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-ExcelFile -Path $path
        $path | Set-ExcelFileProperty -Name $name -Value $value
        (Get-ExcelFileProperty -Path $path).Final | Should -BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-ExcelFileProperty -Name $name -Value $value
        (Get-ExcelFileProperty -LiteralPath $path).Final | Should -BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-ExcelFile -Path $path
        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -Force -WhatIf
        (Get-ExcelFileProperty -LiteralPath $path).Final | Should -BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        New-ExcelFile -Path $path
        $exitCode = 'N' | Invoke-Confirm -Path $path -Name $name -Value $value
        $exitCode | Should -Be 0
        (Get-ExcelFileProperty -LiteralPath $path).Final | Should -BeFalse
      }
      It 'updates a read-only recommended workbook when Force is specified' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -Force -Confirm
        (Get-ExcelFileProperty -Path $path).Final | Should -BeTrue
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should -Throw
        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        (Get-ExcelFileProperty -LiteralPath $path -PasswordToOpen $password).Final | Should -BeTrue
      }
      It 'updates a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should -Throw
        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        (Get-ExcelFileProperty -LiteralPath $path -PasswordToModify $password).Final | Should -BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-ExcelFile -Path $path
        $property = Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PassThru
        $property | Should -BeOfType [PSCustomObject]
        $property.Final | Should -BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        $property = Set-ExcelFileProperty -LiteralPath $path -InputObject $properties -PassThru
        $property | Should -BeOfType [PSCustomObject]
        $property.Final | Should -BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-ExcelFile -Path $path
        { Set-ExcelFileProperty -Path $path -Name 'NonExistentProperty' -Value 'Value' } | Should -Throw
      }
      It 'throws an error when read-only recommended file without Force' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value } | Should -Throw
      }
    }
  }
  Describe 'Set-ExcelFileProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null

      $script:saveCalled = 0
      $script:called = 0

      $app = [PSCustomObject]@{}
      $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }

      $file = [PSCustomObject]@{ ReadOnly = $false }
      $file | Add-Member -MemberType ScriptMethod -Name Save -Value { $script:saveCalled++ }
      $file | Add-Member -MemberType ScriptMethod -Name Close -Value { }

      Mock -CommandName New-ExcelObject -MockWith { $app }
      Mock -CommandName Open-ExcelFile -MockWith {
        $script:called++
        $file
      }
      Mock -CommandName Set-ObjectProperty -MockWith { }
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call Open-ExcelFile when WhatIf is specified' {
        Set-ExcelFileProperty -LiteralPath $path -Name Final -Value $false -WhatIf
        $script:called | Should -Be 0
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        { Set-ExcelFileProperty -LiteralPath $path -Name Final -Value $false } | Should -Throw
      }
    }
  }
  Describe 'Test-ExcelExtension' {
    BeforeEach {
      Mock -CommandName Test-Path -MockWith { $true }
      Mock -CommandName Get-Item -MockWith {
        param(
          [string]
          $LiteralPath
        )
        [PSCustomObject]@{
          PSIsContainer = $false
          PSPath        = ($env:TEMP | Join-Path -ChildPath $LiteralPath)
          FullName      = ($env:TEMP | Join-Path -ChildPath $LiteralPath)
        }
      }
    }
    It 'returns true for *.xlsx files by Path' {
      Mock -CommandName Get-Item -ParameterFilter { $Path -eq '*.txt' } -MockWith {
        1, 2 |
        ForEach-Object {
          [PSCustomObject]@{
            PSIsContainer = $false
            PSPath        = ($env:TEMP | Join-Path -ChildPath "$_.txt")
            FullName      = ($env:TEMP | Join-Path -ChildPath "$_.txt")
          }
        }
      }
      Mock -CommandName Get-Item -ParameterFilter { $Path -eq '*.xlsx' } -MockWith {
        1, 2 |
        ForEach-Object {
          [PSCustomObject]@{
            PSIsContainer = $false
            PSPath        = ($env:TEMP | Join-Path -ChildPath "$_.xlsx")
            FullName      = ($env:TEMP | Join-Path -ChildPath "$_.xlsx")
          }
        }
      }
      Test-ExcelExtension -Path '*.txt' | Should -BeFalse
      Test-ExcelExtension -Path '*.xlsx' | Should -BeTrue
    }
    It 'returns true for *.xlsx files by LiteralPath' {
      Test-ExcelExtension -LiteralPath 'Book.txt' | Should -BeFalse
      Test-ExcelExtension -LiteralPath 'Book.xlsx' | Should -BeTrue
    }
  }
  Describe 'Test-ExcelExtension.Unit' {
    Context 'Edge cases' {
      It 'returns false when Get-Item throws ItemNotFoundException' {
        Mock -CommandName Get-Item -MockWith {
          throw [ItemNotFoundException]::new('not found')
        }

        Test-ExcelExtension -Path '*.xlsx' | Should -BeFalse
      }
      It 'returns false when matched items are not leaf paths' {
        Mock -CommandName Get-Item -MockWith {
          [PSCustomObject]@{
            FullName = ($env:TEMP | Join-Path -ChildPath 'Book.xlsx')
          }
        }
        Mock -CommandName Test-Path -ParameterFilter { $IsValid } -MockWith { $true }
        Mock -CommandName Test-Path -ParameterFilter { $PathType -eq 'Leaf' } -MockWith { $false }

        Test-ExcelExtension -LiteralPath 'Book.xlsx' | Should -BeFalse
      }
    }
  }
}

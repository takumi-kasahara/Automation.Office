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
      $password = Get-Password
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
      $password = Get-Password
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
      $password = Get-Password
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
        { Get-AccessFileProperty -Path $path -Password (Get-Password) } | Should-Throw
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
      $password = Get-Password
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
        { Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -Password (Get-Password) } | Should-Throw
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
}

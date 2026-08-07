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
      $password = Get-Password
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'creates a database by Path' {
        $path = Get-TempFile
        $item = New-AccessFile -Path $path
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a database by Path with ValueFromPipeline' {
        $path = Get-TempFile
        $item = $path | New-AccessFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a database by Path with ValueFromPipelineByPropertyName' {
        $path = Get-TempFile
        $item = [PSCustomObject]@{ FullName = $path } | New-AccessFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a database when WhatIf is specified' {
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
      It 'creates a database using the requested file format' {
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
        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path -Password $password
          try {
            $app.CurrentProject.FullName | Should-Be $path
          } finally {
            $app.CloseCurrentDatabase()
          }
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'creates a database with RemovePersonalInformation' {
        $path = Get-TempFile
        New-AccessFile -Path $path -RemovePersonalInformation
        Test-Path -LiteralPath $path | Should-BeTrue
        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path
          try {
            $app.CurrentProject.RemovePersonalInformation | Should-BeTrue
          } finally {
            $app.CloseCurrentDatabase()
          }
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'creates a database with InitializeDb script block' {
        $path = Get-TempFile
        $item = New-AccessFile -Path $path -InitializeDb {
          param(
            [Parameter(Mandatory)]
            [Microsoft.Office.Interop.Access.Dao.Database]
            $database
          )
          $database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
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
            [Parameter(Mandatory)]
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
        { New-AccessFile -Path $path } | Should-Throw -ExceptionType [System.IO.IOException]
      }
    }
  }
  Describe 'New-AccessFile.Unit' {
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
      It 'does not call New-AccessObject when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith {
          $script:called++
          throw 'must not be called'
        }

        { New-AccessFile -Path $path -Force -WhatIf } | Should -Not -Throw
        $script:called | Should-Be 0
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-AccessObject when path exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null
        Mock -CommandName New-AccessObject -MockWith {
          $script:called++
          throw 'must not be called'
        }

        { New-AccessFile -Path $path } | Should-Throw -ExceptionType [System.IO.IOException]
        $script:called | Should-Be 0
      }
    }
  }
  Describe 'Get-AccessAppProperty' {
    It 'returns application properties' {
      $properties = Get-AccessAppProperty

      $properties.Visible | Should-NotBeNull
    }
  }
  Describe 'Get-AccessAppProperty.Unit' {
    It 'throws when New-AccessObject fails and requests NoSetup' {
      Mock -CommandName New-AccessObject -MockWith { throw 'new access object failed' }

      { Get-AccessAppProperty } | Should-Throw -ExceptionType [System.InvalidOperationException]
      Should -Invoke -CommandName New-AccessObject -ParameterFilter { $NoSetup } -Times 1 -Exactly
    }
  }
  Describe 'Set-AccessAppProperty' {
    It 'sets application properties' {
      $properties = [PSCustomObject]@{ Visible = $false }
      { Set-AccessAppProperty -Properties $properties } | Should -Not -Throw
    }
    It 'does not set properties when WhatIf is specified' {
      $properties = [PSCustomObject]@{ Visible = $false }
      { Set-AccessAppProperty -Properties $properties -WhatIf } | Should -Not -Throw
    }
    It 'throws when trying to set a property that does not exist.' {
      $properties = [PSCustomObject]@{ NonExistentProperty = 'Value' }
      { Set-AccessAppProperty -Properties $properties } | Should-Throw -ExceptionType [System.Management.Automation.RuntimeException]
    }
  }
  Describe 'Set-AccessAppProperty.Unit' {
    BeforeEach {
      $properties = [PSCustomObject]@{ Visible = $false }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-AccessObject when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith { throw 'must not be called' }

        { Set-AccessAppProperty -Properties $properties -WhatIf } | Should -Not -Throw
        Should -Invoke -CommandName New-AccessObject -Times 0 -Exactly
      }
    }
    Context 'ParameterSetName' {
      It 'throws when New-AccessObject fails' {
        Mock -CommandName New-AccessObject -MockWith { throw 'new access object failed' }

        { Set-AccessAppProperty -Properties $properties } | Should-Throw -ExceptionType [System.InvalidOperationException]
      }
    }
  }
  Describe 'Get-AccessFileProperty' {
    BeforeEach {
      $ConfirmPreference = 'None'
      $path = Get-TempFile
      New-AccessFile -Path $path -RemovePersonalInformation
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'retrieves file properties by Path' {
        $properties = Get-AccessFileProperty -Path $path
        $properties | Should-NotBeNull
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'retrieves file properties by LiteralPath' {
        $properties = Get-AccessFileProperty -LiteralPath $path
        $properties | Should-NotBeNull
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'retrieves file properties by Path with ValueFromPipeline' {
        $properties = $path | Get-AccessFileProperty -Name RemovePersonalInformation
        $properties | Should-NotBeNull
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'retrieves file properties by Path with ValueFromPipelineByPropertyName' {
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-AccessFileProperty -Name RemovePersonalInformation
        $properties | Should-NotBeNull
        $properties.RemovePersonalInformation | Should-Be $true
      }
    }
    Context 'Other parameters' {
      It 'retrieves file properties with -Name filter' {
        $properties = Get-AccessFileProperty -Path $path -Name RemovePersonalInformation
        $properties.RemovePersonalInformation | Should-Be $true
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
      It 'opens a file protected with Password' {
        $password = Get-Password
        $protectedPath = Get-TempFile
        try {
          New-AccessFile -Path $protectedPath -Password $password
          $properties = Get-AccessFileProperty -Path $protectedPath -Password $password
          $properties | Should-NotBeNull
        } finally {
          if (Test-Path -LiteralPath $protectedPath) {
            Remove-Item -LiteralPath $protectedPath -Force
          }
        }
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
        Mock -CommandName New-AccessObject -MockWith { throw 'new access object failed' }

        { Get-AccessFileProperty -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Set-AccessFileProperty' {
    BeforeEach {
      $ConfirmPreference = 'None'
      $password = Get-Password
      $path = Get-TempFile
      New-AccessFile -Path $path
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'sets file property by Path' {
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'sets file property by LiteralPath' {
        Set-AccessFileProperty -LiteralPath $path -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -LiteralPath $path
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'sets file property by Path with ValueFromPipeline' {
        $path | Set-AccessFileProperty -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'sets file property by Path with ValueFromPipelineByPropertyName' {
        [PSCustomObject]@{ PSPath = $path } | Set-AccessFileProperty -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -LiteralPath $path
        $properties.RemovePersonalInformation | Should-Be $true
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update property when WhatIf is specified' {
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -WhatIf
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should-Be $false
      }
    }
    Context 'Other parameters' {
      It 'sets file property with -InputObject' {
        $inputObject = [PSCustomObject]@{ RemovePersonalInformation = $true }
        Set-AccessFileProperty -Path $path -InputObject $inputObject
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should-Be $true
      }
      It 'returns updated file properties when PassThru is specified' {
        $result = Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -PassThru
        $result | Should-NotBeNull
        $result.RemovePersonalInformation | Should-Be $true
      }
      It 'updates a file protected with Password' {
        $password = Get-Password
        $protectedPath = Get-TempFile
        try {
          New-AccessFile -Path $protectedPath -Password $password
          Set-AccessFileProperty -Path $protectedPath -Name RemovePersonalInformation -Value $true -Password $password
          $properties = Get-AccessFileProperty -Path $protectedPath -Password $password
          $properties.RemovePersonalInformation | Should-Be $true
        } finally {
          if (Test-Path -LiteralPath $protectedPath) {
            Remove-Item -LiteralPath $protectedPath -Force
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist' {
        { Set-AccessFileProperty -Path $path -Name NonExistentProperty -Value 'test' } | Should-Throw
      }
    }
  }
  Describe 'Open-AccessFile' {
    BeforeAll {
      # Reuse existing helper functions from parent scope
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
      It 'opens a database by Path and returns nothing when no action specified' {
        $path = Get-TempFile
        New-AccessFile -Path $path | Out-Null
        $app = New-AccessObject
        try {
          $result = Open-AccessFile -Application $app -Path $path
          $result | Should-BeNull
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'opens a database by Path with ValueFromPipeline and returns nothing when no action specified' {
        $path = Get-TempFile
        New-AccessFile -Path $path | Out-Null
        $app = New-AccessObject
        try {
          $result = $path | Open-AccessFile -Application $app
          $result | Should-BeNull
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'opens a database by Path with ValueFromPipelineByPropertyName and returns nothing when no action specified' {
        $path = Get-TempFile
        New-AccessFile -Path $path | Out-Null
        $app = New-AccessObject
        try {
          $result = [PSCustomObject]@{ FullName = $path } | Open-AccessFile -Application $app
          $result | Should-BeNull
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
    }
    Context 'Other parameters' {
      It 'opens a database with Password' {
        $path = Get-TempFile
        New-AccessFile -Path $path -Password $password | Out-Null
        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path -Password $password | Out-Null
          Test-Path -LiteralPath $path | Should-BeTrue
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'opens a database and executes ActionDb script block' {
        $path = Get-TempFile
        New-AccessFile -Path $path | Out-Null
        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path -ActionDb {
            param($db)
            $db.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
          }
          # Close the database after ActionDb completes
          $app.CloseCurrentDatabase()
          # Now verify the table was created
          $tables = Get-AccessTable -Path $path
          ($tables | Where-Object -Property Name -EQ 'Employees') | Should-NotBeNull
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'opens a database and executes ActionProject script block' {
        $path = Get-TempFile
        New-AccessFile -Path $path | Out-Null
        $app = New-AccessObject
        try {
          $connection = Open-AccessFile -Application $app -Path $path -ActionProject {
            param($project)
            return $project.Connection
          }
          $connection | Should-NotBeNull
        } finally {
          try {
            if ($app) {
              $app.Quit()
            }
          } finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      It 'uses the provided Application object' {
        $path = Get-TempFile
        New-AccessFile -Path $path | Out-Null
        $app1 = New-AccessObject
        $app2 = New-AccessObject
        try {
          Open-AccessFile -Application $app1 -Path $path -ActionProject {
            param($project)
            $script:projectApp1 = $project.Application
          }
          Open-AccessFile -Application $app2 -Path $path -ActionProject {
            param($project)
            $script:projectApp2 = $project.Application
          }
          $script:projectApp1 | Should-Be $app1
          $script:projectApp2 | Should-Be $app2
        } finally {
          try {
            if ($app1) {
              $app1.Quit()
            }
          } finally {
            try {
              if ($app2) {
                $app2.Quit()
              }
            } finally {
              Get-Variable |
              Where-Object -Property Value -Is [__ComObject] |
              Clear-Variable -Force -WhatIf:$false -Confirm:$false
              [GC]::Collect()
              [GC]::WaitForPendingFinalizers()
            }
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'fails when the path does not exist' {
        $path = [IO.Path]::GetTempFileName()
        Remove-Item -LiteralPath $path -Force
        { Open-AccessFile -Path $path } | Should-Throw
      }
    }
  }
}

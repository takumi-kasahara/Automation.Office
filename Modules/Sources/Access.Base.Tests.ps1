using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Security

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest

InModuleScope 'Automation.Office' {
  BeforeAll {
    Add-Type -AssemblyName System.Web
    function Get-Password {
      [CmdletBinding()]
      [OutputType([SecureString])]
      [SuppressMessage('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Used in tests to generate random passwords for verification purposes')]
      param ()
      $value = [guid]::NewGuid().ToString('N').Substring(0, 15)
      return ConvertTo-SecureString -String $value -AsPlainText -Force
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
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $command = @(
            "Import-Module '$escapedModulePath' -Force"
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
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
      It 'creates a database by Path with ValueFromPipeline' {
        $path = Get-TempFile
        $item = $path | New-AccessFile
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
      It 'creates a database by Path with ValueFromPipelineByPropertyName' {
        $path = Get-TempFile
        $item = [PSCustomObject]@{ FullName = $path } | New-AccessFile
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a database when WhatIf is specified' {
        $path = Get-TempFile
        New-AccessFile -Path $path -Force -WhatIf
        Test-Path -LiteralPath $path | Should -BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        $path = Get-TempFile
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should -Be 0
        Test-Path -LiteralPath $path | Should -BeFalse
      }
      It 'overwrites an existing database when Force is specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        New-AccessFile -Path $path -Force -Confirm
        (Get-Item -LiteralPath $path).Length | Should -BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a database using the requested file format' {
        $path = Get-TempFile -Extension '.mdb'
        $item = New-AccessFile -Path $path -FileFormat acNewDatabaseFormatAccess2007
        $item | Should -BeOfType [System.IO.FileInfo]
        $item.FullName | Should -Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should -BeTrue
      }
      It 'creates a database with Password' {
        $path = Get-TempFile
        New-AccessFile -Path $path -Password $password
        Test-Path -LiteralPath $path | Should -BeTrue
        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path -Password $password
          try {
            $app.CurrentProject.FullName | Should -Be $path
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
        Test-Path -LiteralPath $path | Should -BeTrue
        $app = New-AccessObject
        try {
          Open-AccessFile -Application $app -Path $path
          try {
            $app.CurrentProject.RemovePersonalInformation | Should -BeTrue
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
    }
    Context 'Edge cases' {
      It 'fails when the path already exists and Force is not specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        { New-AccessFile -Path $path } | Should -Throw
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
          $script:called | Should -Be 0
          Test-Path -LiteralPath $path | Should -BeFalse
        }
      }
      Context 'Edge cases' {
        It 'throws and does not call New-AccessObject when path exists and Force is not specified' {
          New-Item -Path $path -ItemType File -Force | Out-Null
          Mock -CommandName New-AccessObject -MockWith {
            $script:called++
            throw 'must not be called'
          }

          { New-AccessFile -Path $path } | Should -Throw
          $script:called | Should -Be 0
        }
      }
    }
  }
  Describe 'Get-AccessAppProperty' {
    It 'returns application properties' {
      $properties = Get-AccessAppProperty

      $properties.Visible | Should -Not -BeNullOrEmpty
    }
  }
  Describe 'Get-AccessAppProperty.Unit' {
    It 'throws when New-AccessObject fails and requests NoSetup' {
      Mock -CommandName New-AccessObject -MockWith { throw 'new access object failed' }

      { Get-AccessAppProperty } | Should -Throw
      Assert-MockCalled -CommandName New-AccessObject -ParameterFilter { $NoSetup } -Times 1 -Exactly
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
      { Set-AccessAppProperty -Properties $properties } | Should -Throw
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
        Assert-MockCalled -CommandName New-AccessObject -Times 0 -Exactly
      }
    }
    Context 'ParameterSetName' {
      It 'throws when New-AccessObject fails' {
        Mock -CommandName New-AccessObject -MockWith { throw 'new access object failed' }

        { Set-AccessAppProperty -Properties $properties } | Should -Throw
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
        $properties | Should -Not -BeNullOrEmpty
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'retrieves file properties by LiteralPath' {
        $properties = Get-AccessFileProperty -LiteralPath $path
        $properties | Should -Not -BeNullOrEmpty
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'retrieves file properties by Path with ValueFromPipeline' {
        $properties = $path | Get-AccessFileProperty -Name RemovePersonalInformation
        $properties | Should -Not -BeNullOrEmpty
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'retrieves file properties by Path with ValueFromPipelineByPropertyName' {
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-AccessFileProperty -Name RemovePersonalInformation
        $properties | Should -Not -BeNullOrEmpty
        $properties.RemovePersonalInformation | Should -Be $true
      }
    }
    Context 'Other parameters' {
      It 'retrieves file properties with -Name filter' {
        $properties = Get-AccessFileProperty -Path $path -Name RemovePersonalInformation
        $properties.RemovePersonalInformation | Should -Be $true
        @($properties.PSObject.Properties).Count | Should -Be 1
      }
      It 'opens a password-protected database' {
        $password = Get-Password
        $protectedPath = Get-TempFile
        try {
          New-AccessFile -Path $protectedPath -Password $password
          $properties = Get-AccessFileProperty -Path $protectedPath -Password $password
          $properties | Should -Not -BeNullOrEmpty
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

        { Get-AccessFileProperty -Path $path } | Should -Throw
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
      It 'sets file property by Path and -Name -Value' {
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'sets file property by LiteralPath and -Name -Value' {
        Set-AccessFileProperty -LiteralPath $path -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -LiteralPath $path
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'sets file property by Path with ValueFromPipeline' {
        $path | Set-AccessFileProperty -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'sets file property by Path with ValueFromPipelineByPropertyName' {
        [PSCustomObject]@{ PSPath = $path } | Set-AccessFileProperty -Name RemovePersonalInformation -Value $true
        $properties = Get-AccessFileProperty -LiteralPath $path
        $properties.RemovePersonalInformation | Should -Be $true
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update property when WhatIf is specified' {
        Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -WhatIf
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should -Be $false
      }
    }
    Context 'Other parameters' {
      It 'sets file property with -InputObject' {
        $inputObject = [PSCustomObject]@{ RemovePersonalInformation = $true }
        Set-AccessFileProperty -Path $path -InputObject $inputObject
        $properties = Get-AccessFileProperty -Path $path
        $properties.RemovePersonalInformation | Should -Be $true
      }
      It 'sets file property with -PassThru' {
        $result = Set-AccessFileProperty -Path $path -Name RemovePersonalInformation -Value $true -PassThru
        $result | Should -Not -BeNullOrEmpty
        $result.RemovePersonalInformation | Should -Be $true
      }
      It 'updates a password-protected database' {
        $password = Get-Password
        $protectedPath = Get-TempFile
        try {
          New-AccessFile -Path $protectedPath -Password $password
          Set-AccessFileProperty -Path $protectedPath -Name RemovePersonalInformation -Value $true -Password $password
          $properties = Get-AccessFileProperty -Path $protectedPath -Password $password
          $properties.RemovePersonalInformation | Should -Be $true
        } finally {
          if (Test-Path -LiteralPath $protectedPath) {
            Remove-Item -LiteralPath $protectedPath -Force
          }
        }
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist' {
        { Set-AccessFileProperty -Path $path -Name NonExistentProperty -Value 'test' } | Should -Throw
      }
    }
  }
}

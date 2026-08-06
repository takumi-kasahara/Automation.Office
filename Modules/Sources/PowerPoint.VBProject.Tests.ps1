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
        $Extension = '.pptm'
      )
      return $env:TEMP | Join-Path -ChildPath "Presentation.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function Get-Destination {
      [CmdletBinding()]
      [OutputType([string])]
      param ()
      return $env:TEMP | Join-Path -ChildPath "VBProject.$([guid]::NewGuid().ToString('N')).json"
    }
    function Get-Fixture {
      [CmdletBinding()]
      [OutputType([string])]
      param ()
      return $PSScriptRoot | Join-Path -ChildPath '..\..\Tests\Bin\Presentation.pptm'
    }
    function Get-FixtureSource {
      [CmdletBinding()]
      [OutputType([string])]
      param ()
      return $PSScriptRoot | Join-Path -ChildPath '..\..\Tests\Modules\Presentation\VBProject.json'
    }
    function Get-ComparableVBProjectFromJson {
      [CmdletBinding()]
      [OutputType([PSCustomObject])]
      param (
        [Parameter(Mandatory)]
        [string]
        $LiteralPath
      )
      return Get-Content -LiteralPath $LiteralPath -Encoding UTF8 | ConvertFrom-Json
    }
    function Get-ComparableVBProjectFromFile {
      [CmdletBinding()]
      [OutputType([PSCustomObject])]
      param (
        [Parameter(Mandatory)]
        [string]
        $LiteralPath,
        [SecureString]
        $PasswordToOpen = $null,
        [SecureString]
        $PasswordToModify = $null
      )
      $tempDir = $env:TEMP | Join-Path -ChildPath ([guid]::NewGuid().ToString('N'))
      $destination = $tempDir | Join-Path -ChildPath 'VBProject.json'
      try {
        New-Item -Path $tempDir -ItemType Directory | Out-Null
        Export-PowerPointVBProject -Path $LiteralPath -Destination $destination -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify | Out-Null
        return Get-ComparableVBProjectFromJson -LiteralPath $destination
      } finally {
        if (Test-Path -LiteralPath $tempDir) {
          Remove-Item -LiteralPath $tempDir -Recurse -Force
        }
      }
    }
    $root = $PSScriptRoot | Join-Path -ChildPath '..\..\Tests\Bin'
    $bin = Get-Fixture
    New-PowerPointFile -Path $bin -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -Force
    try {
      Import-PowerPointVBProject -Path $bin -Source (Get-FixtureSource) -Hidden
    } catch {
      Remove-Item -LiteralPath $bin -Force
      throw $_
    }
  }
  AfterAll {
    $destination = Get-FixtureSource
    $componentRoot = [Path]::ChangeExtension($destination, $null)
    if (-not (Test-Path -LiteralPath $componentRoot)) {
      New-Item -Path $componentRoot -ItemType Directory | Out-Null
    }
    Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination
  }
  Describe 'Export-PowerPointVBProject' {
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
          $Destination
        )
        process {
          $resolvedModulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $resolvedModulePath.Replace("'", "''")
          $escapedFilePath = $Path.Replace("'", "''")
          $escapedDestination = $Destination.Replace("'", "''")
          $command = @(
            "Import-Module -Name '$escapedModulePath' -Force"
            "Export-PowerPointVBProject -Path '$escapedFilePath' -Destination '$escapedDestination' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $destination = Get-Destination
      $componentRoot = [Path]::ChangeExtension($destination, $null)
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Force
      }
      if (Test-Path -LiteralPath $componentRoot) {
        Remove-Item -LiteralPath $componentRoot -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'exports VBProject by FilePath as JSON' {
        $item = Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination
        $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
        $item.FullName | Should -Be ([Path]::GetFullPath($destination))
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        @($exported.VBComponents).Count | Should -BeGreaterThan 0
        @($exported.References).Count | Should -BeGreaterThan 0
      }
      It 'exports VBProject by Path with ValueFromPipeline' {
        $item = (Get-Fixture) | Export-PowerPointVBProject -Destination $destination
        $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
        $item.FullName | Should -Be ([Path]::GetFullPath($destination))
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        @($exported.VBComponents).Count | Should -BeGreaterThan 0
        @($exported.References).Count | Should -BeGreaterThan 0
      }
      It 'exports VBProject by Path with ValueFromPipelineByPropertyName' {
        $item = [PSCustomObject]@{ Path = (Get-Fixture) } | Export-PowerPointVBProject -Destination $destination
        $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
        $item.FullName | Should -Be ([Path]::GetFullPath($destination))
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        @($exported.VBComponents).Count | Should -BeGreaterThan 0
        @($exported.References).Count | Should -BeGreaterThan 0
      }
      It 'exports VBProject components to the specified ComponentRoot' {
        $customRoot = $env:TEMP | Join-Path -ChildPath "VBComponents.$([guid]::NewGuid().ToString('N'))"
        try {
          $item = Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -ComponentRoot $customRoot
          $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
          $item.FullName | Should -Be ([Path]::GetFullPath($destination))
          Test-Path -LiteralPath $customRoot -PathType Container | Should -BeTrue
          Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
          @($exported.VBComponents).Count | Should -BeGreaterThan 0
          @($exported.References).Count | Should -BeGreaterThan 0
        } finally {
          if (Test-Path -LiteralPath $customRoot) {
            Remove-Item -LiteralPath $customRoot -Recurse -Force
          }
        }
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create outputs when WhatIf is specified' {
        Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force -WhatIf
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeFalse
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        $exitCode = 'N' | Invoke-Confirm -Path (Get-Fixture) -Destination $destination
        $exitCode | Should -Be 0
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeFalse
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'overwrites an existing read-only destination when Force is specified' {
        (New-Item -Path $destination -ItemType File).IsReadOnly = $true
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination } | Should -Throw
        Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force -Confirm
        (Get-Item -LiteralPath $destination -Force).IsReadOnly | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
      It 'overwrites an existing read-only component when Force is specified' {
        New-Item -Path $componentRoot -ItemType Directory | Out-Null
        $bas = New-Item -Path ($componentRoot | Join-Path -ChildPath 'Module1.bas') -ItemType File
        $bas.IsReadOnly = $true
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination } | Should -Throw
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeFalse
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force -Confirm
        $bas.IsReadOnly | Should -BeTrue
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
      It 'throws and keeps the existing destination content unchanged when NoClobber is specified' {
        New-Item -Path $destination -ItemType File | Out-Null
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -NoClobber } | Should -Throw
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force -NoClobber } | Should -Throw
        (Get-Item -LiteralPath $destination).Length | Should -Be 0
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'throws and keeps the existing component content unchanged when NoClobber is specified' {
        New-Item -Path $componentRoot -ItemType Directory | Out-Null
        New-Item -Path ($componentRoot | Join-Path -ChildPath 'Module1.bas') -ItemType File | Out-Null
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -NoClobber } | Should -Throw
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force -NoClobber } | Should -Throw
        Test-Path -LiteralPath ($componentRoot | Join-Path -ChildPath 'Module1.bas') -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
    }
    Context 'Other parameters' {
      It 'exports VBProject from a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -PasswordToOpen $password
        { Export-PowerPointVBProject -Path $path -Destination $destination -PasswordToOpen (Get-Password) } | Should -Throw
        Export-PowerPointVBProject -Path $path -Destination $destination -PasswordToOpen $password
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
      It 'exports VBProject from a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -PasswordToModify $password
        Export-PowerPointVBProject -Path $path -Destination $destination -PasswordToModify $password
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when the destination is an existing directory' {
        New-Item -Path $destination -ItemType Directory | Out-Null
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force } | Should -Throw
        Test-Path -LiteralPath $destination -PathType Container | Should -BeTrue
      }
      It 'throws when the destination path already exists as a directory' {
        New-Item -Path $destination -ItemType Directory | Out-Null
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -NoClobber } | Should -Throw
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force -NoClobber } | Should -Throw
        Test-Path -LiteralPath $destination -PathType Container | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'throws when the component directory path already exists as a file' {
        New-Item -Path $componentRoot -ItemType File | Out-Null
        { Export-PowerPointVBProject -Path (Get-Fixture) -Destination $destination -Force } | Should -Throw
        Test-Path -LiteralPath $componentRoot -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $destination | Should -BeFalse
      }
    }
  }
  Describe 'Export-PowerPointVBProject.Unit' {
    BeforeAll {
      Mock -CommandName Test-Path -MockWith { $true }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-PowerPointObject when WhatIf is specified' {
        Mock -CommandName New-PowerPointObject -MockWith { throw 'must not be called' }

        { Export-PowerPointVBProject -Path 'dummy.pptx' -Destination 'dummy' -WhatIf } | Should -Not -Throw
        Assert-MockCalled -CommandName New-PowerPointObject -Times 0 -Exactly
      }
    }
    Context 'Edge cases' {
      It 'throws when New-PowerPointObject fails' {
        Mock -CommandName New-PowerPointObject -MockWith { throw 'new powerpoint object failed' }

        { Export-PowerPointVBProject -Path 'dummy.pptx' -Destination 'dummy' } | Should -Throw
      }
    }
  }
  Describe 'Import-PowerPointVBProject' {
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
          $Source
        )
        process {
          $resolvedModulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $resolvedModulePath.Replace("'", "''")
          $escapedFilePath = $Path.Replace("'", "''")
          $escapedSource = $Source.Replace("'", "''")
          $command = @(
            "Import-Module -Name '$escapedModulePath' -Force"
            "Import-PowerPointVBProject -Path '$escapedFilePath' -Source '$escapedSource' -Hidden -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $source = Get-FixtureSource
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'ParameterSetName' {
      It 'imports VBProject by FilePath from JSON' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled
        Import-PowerPointVBProject -Path $path -Source $source -Hidden
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
      It 'imports VBProject by Path with ValueFromPipeline' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled
        $path | Import-PowerPointVBProject -Source $source -Hidden
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
      It 'imports VBProject by Path with ValueFromPipelineByPropertyName' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled
        [PSCustomObject]@{ Path = $path } | Import-PowerPointVBProject -Source $source -Hidden
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not import VBProject when WhatIf is specified' {
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled
        $ante = Get-ComparableVBProjectFromFile -LiteralPath $path
        Import-PowerPointVBProject -Path $path -Source $source -Hidden -Force -WhatIf
        $post = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($post | ConvertTo-Json) | Should -Be ($ante | ConvertTo-Json)
      }
      It 'asks for confirmation when Confirm is specified' {
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled
        $ante = Get-ComparableVBProjectFromFile -LiteralPath $path
        $exitCode = 'N' | Invoke-Confirm -Path $path -Source $source
        $exitCode | Should -Be 0
        $post = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($post | ConvertTo-Json) | Should -Be ($ante | ConvertTo-Json)
      }
    }
    Context 'Other parameters' {
      It 'imports VBProject into a file protected with PasswordToOpen' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -PasswordToOpen $password
        { Import-PowerPointVBProject -Path $path -Source $source -Hidden -PasswordToOpen (Get-Password) } | Should -Throw
        Import-PowerPointVBProject -Path $path -Source $source -Hidden -PasswordToOpen $password
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path -PasswordToOpen $password
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
      It 'imports VBProject into a file protected with PasswordToModify' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -PasswordToModify $password
        { Import-PowerPointVBProject -Path $path -Source $source -Hidden -PasswordToModify (Get-Password) } | Should -Throw
        Import-PowerPointVBProject -Path $path -Source $source -Hidden -PasswordToModify $password
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path -PasswordToModify $password
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read-only recommended presentation' {
        New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -ReadOnlyRecommended
        { Import-PowerPointVBProject -Path $path -Source $source -Hidden } | Should -Throw
      }
    }
  }
  Describe 'Import-PowerPointVBProject.Unit' {
    BeforeEach {
      $path = Get-TempFile
      $source = Get-FixtureSource
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-PowerPointObject when WhatIf is specified' {
        Mock -CommandName New-PowerPointObject -MockWith { throw 'must not be called' }

        { Import-PowerPointVBProject -Path $path -Source $source -Hidden -WhatIf } | Should -Not -Throw
        Assert-MockCalled -CommandName New-PowerPointObject -Times 0 -Exactly
      }
    }
    Context 'Edge cases' {
      It 'throws when New-PowerPointObject fails' {
        Mock -CommandName New-PowerPointObject -MockWith {
          throw 'failed to create PowerPoint object'
        }
        { Import-PowerPointVBProject -Path $path -Source $source -Hidden } | Should -Throw 'failed to create PowerPoint object'
      }
    }
  }
}

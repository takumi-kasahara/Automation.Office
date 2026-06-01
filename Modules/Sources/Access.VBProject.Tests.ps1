using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Security

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest

InModuleScope 'Access.VBProject' {
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
        $Extension = '.accdb'
      )
      return $env:TEMP | Join-Path -ChildPath "Database.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function Get-Destination {
      [CmdletBinding()]
      [OutputType([string])]
      param ()
      return $env:TEMP | Join-Path -ChildPath "VBProject.$([guid]::NewGuid().ToString('N')).json"
    }
    function Get-FixtureWorkbook {
      [CmdletBinding()]
      [OutputType([string])]
      param ()
      return $PSScriptRoot | Join-Path -ChildPath '..\..\Tests\Bin\Database.accdb'
    }
    function Get-FixtureSource {
      [CmdletBinding()]
      [OutputType([string])]
      param ()
      return $PSScriptRoot | Join-Path -ChildPath '..\..\Tests\Modules\Database\VBProject.json'
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
        $Password = $null
      )
      $tempDir = $env:TEMP | Join-Path -ChildPath ([guid]::NewGuid().ToString('N'))
      $destination = $tempDir | Join-Path -ChildPath 'VBProject.json'
      try {
        New-Item -Path $tempDir -ItemType Directory | Out-Null
        Export-AccessVBProject -Path $LiteralPath -Destination $destination -Password $Password | Out-Null
        return Get-ComparableVBProjectFromJson -LiteralPath $destination
      }
      finally {
        if (Test-Path -LiteralPath $tempDir) {
          Remove-Item -LiteralPath $tempDir -Recurse -Force
        }
      }
    }
    $root = $PSScriptRoot | Join-Path -ChildPath '..\..\Tests\Bin'
    $bin = Get-FixtureWorkbook
    New-AccessFile -Path $bin -RemovePersonalInformation -Force
    try {
      Import-AccessVBProject -Path $bin -Source (Get-FixtureSource)
    }
    catch {
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
    Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination
  }
  Describe 'Export-AccessVBProject' {
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
          $resolvedModulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'))
          $escapedModulePath = $resolvedModulePath.Replace("'", "''")
          $escapedFilePath = $Path.Replace("'", "''")
          $escapedDestination = $Destination.Replace("'", "''")
          $command = @(
            "Import-Module '$escapedModulePath' -Force"
            "Export-AccessVBProject -Path '$escapedFilePath' -Destination '$escapedDestination' -Confirm"
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
        $item = Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination
        $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
        $item.FullName | Should -Be ([Path]::GetFullPath($destination))
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        @($exported.VBComponents).Count | Should -BeGreaterThan 0
        @($exported.References).Count | Should -BeGreaterThan 0
      }
      It 'exports VBProject by Path with ValueFromPipeline' {
        $item = (Get-FixtureWorkbook) | Export-AccessVBProject -Destination $destination
        $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
        $item.FullName | Should -Be ([Path]::GetFullPath($destination))
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        @($exported.VBComponents).Count | Should -BeGreaterThan 0
        @($exported.References).Count | Should -BeGreaterThan 0
      }
      It 'exports VBProject by Path with ValueFromPipelineByPropertyName' {
        $item = [PSCustomObject]@{ Path = (Get-FixtureWorkbook) } | Export-AccessVBProject -Destination $destination
        $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
        $item.FullName | Should -Be ([Path]::GetFullPath($destination))
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        @($exported.VBComponents).Count | Should -BeGreaterThan 0
        @($exported.References).Count | Should -BeGreaterThan 0
      }
      It 'exports VBProject components to the specified ComponentRoot' {
        $customRoot = $env:TEMP | Join-Path -ChildPath "VBComponents.$([guid]::NewGuid().ToString('N'))"
        try {
          $item = Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -ComponentRoot $customRoot
          $exported = Get-Content -LiteralPath $destination -Encoding UTF8 | ConvertFrom-Json
          $item.FullName | Should -Be ([Path]::GetFullPath($destination))
          Test-Path -LiteralPath $customRoot -PathType Container | Should -BeTrue
          Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
          @($exported.VBComponents).Count | Should -BeGreaterThan 0
          @($exported.References).Count | Should -BeGreaterThan 0
        }
        finally {
          if (Test-Path -LiteralPath $customRoot) {
            Remove-Item -LiteralPath $customRoot -Recurse -Force
          }
        }
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create outputs when WhatIf is specified' {
        Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force -WhatIf
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeFalse
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        $exitCode = 'N' | Invoke-Confirm -Path (Get-FixtureWorkbook) -Destination $destination
        $exitCode | Should -Be 0
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeFalse
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'overwrites an existing read-only destination when Force is specified' {
        (New-Item -Path $destination -ItemType File).IsReadOnly = $true
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination } | Should -Throw
        Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force -Confirm
        (Get-Item -LiteralPath $destination -Force).IsReadOnly | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
      It 'overwrites an existing read-only component when Force is specified' {
        New-Item -Path $componentRoot -ItemType Directory | Out-Null
        $bas = New-Item -Path ($componentRoot | Join-Path -ChildPath 'Module1.bas') -ItemType File
        $bas.IsReadOnly = $true
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination } | Should -Throw
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeFalse
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
        Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force -Confirm
        $bas.IsReadOnly | Should -BeTrue
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
      It 'throws and keeps the existing destination content unchanged when NoClobber is specified' {
        New-Item -Path $destination -ItemType File | Out-Null
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -NoClobber } | Should -Throw
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force -NoClobber } | Should -Throw
        (Get-Item -LiteralPath $destination).Length | Should -Be 0
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'throws and keeps the existing component content unchanged when NoClobber is specified' {
        New-Item -Path $componentRoot -ItemType Directory | Out-Null
        New-Item -Path ($componentRoot | Join-Path -ChildPath 'Module1.bas') -ItemType File | Out-Null
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -NoClobber } | Should -Throw
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force -NoClobber } | Should -Throw
        Test-Path -LiteralPath ($componentRoot | Join-Path -ChildPath 'Module1.bas') -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
    }
    Context 'Other parameters' {
      It 'exports VBProject from a database protected with Password' {
        New-AccessFile -Path $path -Password $password
        { Export-AccessVBProject -Path $path -Destination $destination -Password (Get-Password) } | Should -Throw
        Export-AccessVBProject -Path $path -Destination $destination -Password $password
        Test-Path -LiteralPath $destination -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when the destination is an existing directory' {
        New-Item -Path $destination -ItemType Directory | Out-Null
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force } | Should -Throw
        Test-Path -LiteralPath $destination -PathType Container | Should -BeTrue
      }
      It 'throws when the destination path already exists as a directory' {
        New-Item -Path $destination -ItemType Directory | Out-Null
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -NoClobber } | Should -Throw
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force -NoClobber } | Should -Throw
        Test-Path -LiteralPath $destination -PathType Container | Should -BeTrue
        Test-Path -LiteralPath $componentRoot -PathType Container | Should -BeFalse
      }
      It 'throws when the component directory path already exists as a file' {
        New-Item -Path $componentRoot -ItemType File | Out-Null
        { Export-AccessVBProject -Path (Get-FixtureWorkbook) -Destination $destination -Force } | Should -Throw
        Test-Path -LiteralPath $componentRoot -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath $destination | Should -BeFalse
      }
    }
  }
  Describe 'Export-AccessVBProject.Unit' {
    BeforeEach {
      $path = Get-FixtureWorkbook
      $destination = Get-Destination
    }
    AfterEach {
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Force
      }
      $componentRoot = [Path]::ChangeExtension($destination, $null)
      if (Test-Path -LiteralPath $componentRoot) {
        Remove-Item -LiteralPath $componentRoot -Recurse -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-AccessObject when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith { throw 'must not be called' }

        { Export-AccessVBProject -Path $path -Destination $destination -WhatIf } | Should -Not -Throw
        Assert-MockCalled -CommandName New-AccessObject -Times 0 -Exactly
        Test-Path -LiteralPath $destination | Should -BeFalse
      }
    }
  }
  Describe 'Import-AccessVBProject.Unit' {
    BeforeEach {
      $path = Get-TempFile
      $source = Get-FixtureSource
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'Failure propagation' {
      It 'throws when New-AccessObject fails' {
        Mock -CommandName New-AccessObject -MockWith {
          throw 'failed to create Access object'
        }
        { Import-AccessVBProject -Path $path -Source $source } | Should -Throw 'failed to create Access object'
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-AccessObject when WhatIf is specified' {
        Mock -CommandName New-AccessObject -MockWith { throw 'must not be called' }

        { Import-AccessVBProject -Path $path -Source $source -WhatIf } | Should -Not -Throw
        Assert-MockCalled -CommandName New-AccessObject -Times 0 -Exactly
      }
    }
  }
  Describe 'Import-AccessVBProject' {
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
          $resolvedModulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'))
          $escapedModulePath = $resolvedModulePath.Replace("'", "''")
          $escapedFilePath = $Path.Replace("'", "''")
          $escapedSource = $Source.Replace("'", "''")
          $command = @(
            "Import-Module '$escapedModulePath' -Force"
            "Import-AccessVBProject -Path '$escapedFilePath' -Source '$escapedSource' -Confirm"
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
        New-AccessFile -Path $path -RemovePersonalInformation
        Import-AccessVBProject -Path $path -Source $source
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
      It 'imports VBProject by Path with ValueFromPipeline' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-AccessFile -Path $path -RemovePersonalInformation
        $path | Import-AccessVBProject -Source $source
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
      It 'imports VBProject by Path with ValueFromPipelineByPropertyName' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-AccessFile -Path $path -RemovePersonalInformation
        [PSCustomObject]@{ Path = $path } | Import-AccessVBProject -Source $source
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not import VBProject when WhatIf is specified' {
        New-AccessFile -Path $path -RemovePersonalInformation
        $ante = Get-ComparableVBProjectFromFile -LiteralPath $path
        Import-AccessVBProject -Path $path -Source $source -WhatIf
        $post = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($post | ConvertTo-Json) | Should -Be ($ante | ConvertTo-Json)
      }
      It 'asks for confirmation when Confirm is specified' {
        New-AccessFile -Path $path -RemovePersonalInformation
        $ante = Get-ComparableVBProjectFromFile -LiteralPath $path
        $exitCode = 'N' | Invoke-Confirm -Path $path -Source $source
        $exitCode | Should -Be 0
        $post = Get-ComparableVBProjectFromFile -LiteralPath $path
        ($post | ConvertTo-Json) | Should -Be ($ante | ConvertTo-Json)
      }
    }
    Context 'Other parameters' {
      It 'imports VBProject into a database protected with Password' {
        $expected = Get-ComparableVBProjectFromJson -LiteralPath $source
        New-AccessFile -Path $path -Password $password
        { Import-AccessVBProject -Path $path -Source $source -Password (Get-Password) } | Should -Throw
        Import-AccessVBProject -Path $path -Source $source -Password $password
        $actual = Get-ComparableVBProjectFromFile -LiteralPath $path -Password $password
        ($actual | ConvertTo-Json) | Should -Be ($expected | ConvertTo-Json)
      }
    }
    Context 'Edge cases' {
      It 'throws an error when the database is read-only' {
        New-AccessFile -Path $path -RemovePersonalInformation
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        { Import-AccessVBProject -Path $path -Source $source } | Should -Throw
      }
    }
  }
}

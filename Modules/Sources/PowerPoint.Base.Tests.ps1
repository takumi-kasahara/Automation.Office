using assembly Microsoft.Office.Interop.PowerPoint
using assembly System.Web
using module .\..\Automation.Office.psd1
using namespace Microsoft.Office.Interop.PowerPoint
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation

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
        $Extension = '.pptx'
      )
      return $env:TEMP | Join-Path -ChildPath "Presentation.$([guid]::NewGuid().ToString('N'))$Extension"
    }
  }
  Describe 'New-PowerPointFile' {
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
            "New-PowerPointFile -Path '$escapedPath' -Confirm"
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
        $item = New-PowerPointFile -Path $path
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $path = Get-TempFile
        $item = $path | New-PowerPointFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $path = Get-TempFile
        $item = [PSCustomObject]@{ FullName = $path } | New-PowerPointFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a file when WhatIf is specified' {
        $path = Get-TempFile
        New-PowerPointFile -Path $path -Force -WhatIf
        Test-Path -LiteralPath $path | Should-BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        $path = Get-TempFile
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should-Be 0
        Test-Path -LiteralPath $path | Should-BeFalse
      }
      It 'overwrites an existing presentation when Force is specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        New-PowerPointFile -Path $path -Force -Confirm
        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.pptm'
        $item = New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file with PasswordToOpen' {
        $path = Get-TempFile
        New-PowerPointFile -Path $path -PasswordToOpen $password
        Test-Path -LiteralPath $path | Should-BeTrue
        Open-PowerPointFile -Path $path -PasswordToOpen $password -Action {
          param(
            [Microsoft.Office.Interop.PowerPoint.Presentation]
            $Presentation
          )
          $Presentation.Password | Should-NotBeEmptyString
        }
      }
      It 'creates a file with PasswordToModify' {
        $path = Get-TempFile
        New-PowerPointFile -Path $path -PasswordToModify $password
        Test-Path -LiteralPath $path | Should-BeTrue
        Open-PowerPointFile -Path $path -PasswordToModify $password -Action {
          param(
            [Microsoft.Office.Interop.PowerPoint.Presentation]
            $Presentation
          )
          $Presentation.WritePassword | Should-NotBeEmptyString
        }
      }
      It 'creates a file with ReadOnlyRecommended' {
        $path = Get-TempFile
        New-PowerPointFile -Path $path -ReadOnlyRecommended
        Test-Path -LiteralPath $path | Should-BeTrue
        Open-PowerPointFile -Path $path -Action {
          param(
            [Microsoft.Office.Interop.PowerPoint.Presentation]
            $Presentation
          )
          $Presentation.ReadOnlyRecommended | Should-BeTrue
        }
      }
      It 'creates a file with Initialize script block' {
        $path = Get-TempFile
        $item = New-PowerPointFile -Path $path -Initialize {
          param (
            [Microsoft.Office.Interop.PowerPoint.Presentation]
            $Presentation
          )
          $Presentation.Slides.Add(1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutTitleOnly) | Out-Null
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
        { New-PowerPointFile -Path $path } | Should-Throw
      }
    }
  }
  Describe 'New-PowerPointFile.Unit' {
    BeforeEach {
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-PowerPointObject when WhatIf is specified' {
        Mock -CommandName New-PowerPointObject -MockWith { throw }

        { New-PowerPointFile -Path $path -Force -WhatIf } | Should -Not -Throw
        Should-NotInvoke -CommandName New-PowerPointObject
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-PowerPointObject when path exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null
        Mock -CommandName New-PowerPointObject

        { New-PowerPointFile -Path $path } | Should-Throw
        Should-NotInvoke -CommandName New-PowerPointObject
      }
    }
  }
  Describe 'Open-PowerPointFile' {
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
      It 'opens a presentation by Path and returns the presentation object' {
        New-PowerPointFile -Path $path
        Open-PowerPointFile -Path $path | Should-NotBeNull
      }
      It 'opens a presentation by Path with ValueFromPipeline' {
        New-PowerPointFile -Path $path
        $path | Open-PowerPointFile | Should-NotBeNull
      }
      It 'opens a presentation by Path with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path
        [PSCustomObject]@{ FullName = $path } | Open-PowerPointFile | Should-NotBeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a presentation with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password
        Open-PowerPointFile -Path $path -PasswordToOpen $password | Should-NotBeNull
      }
      It 'opens a presentation with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password
        Open-PowerPointFile -Path $path -PasswordToModify $password | Should-NotBeNull
      }
      It 'opens a presentation with ReadOnly' {
        New-PowerPointFile -Path $path
        Open-PowerPointFile -Path $path -ReadOnly | Should-NotBeNull
      }
      It 'opens a presentation with Action script block' {
        New-PowerPointFile -Path $path
        $result = Open-PowerPointFile -Path $path -Action {
          param(
            [Microsoft.Office.Interop.PowerPoint.Presentation]
            $Presentation
          )
          $Presentation.Slides.Add(1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutTitleOnly) | Out-Null
          return $Presentation.Slides.Count
        }
        $result | Should-Be 1
      }
      It 'opens a presentation using an existing Application object' {
        New-PowerPointFile -Path $path
        $app = New-PowerPointObject
        try {
          Open-PowerPointFile -Path $path -Application $app | Should-NotBeNull
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
  Describe 'Open-PowerPointFile.Unit' {
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
      It 'throws when New-PowerPointObject fails' {
        Mock -CommandName New-PowerPointObject -MockWith { throw }

        { Open-PowerPointFile -Path $path } | Should-Throw
        Should-Invoke -CommandName New-PowerPointObject -Times 1 -Exactly
      }
    }
  }
  Describe 'Get-PowerPointFileProperty' {
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
      New-PowerPointFile -Path $path
      { Get-PowerPointFileProperty -Path $path | Out-Host } | Should -Not -Throw
    }
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-PowerPointFile -Path $path
        $properties = Get-PowerPointFileProperty -Path $path -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.Final | Should-BeFalse
      }
      It 'returns selected file properties by LiteralPath' {
        New-PowerPointFile -Path $path
        $properties = Get-PowerPointFileProperty -LiteralPath $path -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.Final | Should-BeFalse
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-PowerPointFile -Path $path
        $properties = $path | Get-PowerPointFileProperty -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.Final | Should-BeFalse
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-PowerPointFileProperty -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
        $properties.Final | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password
        { Get-PowerPointFileProperty -Path $path } | Should-Throw
        { Get-PowerPointFileProperty -Path $path -PasswordToOpen (Get-Password) } | Should-Throw
        { Get-PowerPointFileProperty -Path $path -PasswordToOpen $password } | Should -Not -Throw
      }
      It 'returns file properties from a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password
        { Get-PowerPointFileProperty -Path $path } | Should -Not -Throw
        { Get-PowerPointFileProperty -Path $path -PasswordToModify (Get-Password) } | Should -Not -Throw
        { Get-PowerPointFileProperty -Path $path -PasswordToModify $password } | Should -Not -Throw
      }
      It 'returns file properties with Recurse and Depth' {
        New-PowerPointFile -Path $path
        { Get-PowerPointFileProperty -Path $path -Recurse -Depth 1 } | Should -Not -Throw
      }
    }
  }
  Describe 'Get-PowerPointFileProperty.Unit' {
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
      It 'throws when New-PowerPointObject fails' {
        Mock -CommandName New-PowerPointObject -MockWith { throw }

        { Get-PowerPointFileProperty -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Set-PowerPointFileProperty' {
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
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $escapedName = $Name.Replace("'", "''")
          $escapedValue = if ($Value) {
            '$true'
          } else {
            '$false'
          }
          $command = @(
            "Import-Module -Name '$escapedModulePath' -Force"
            "Set-PowerPointFileProperty -LiteralPath '$escapedPath' -Name '$escapedName' -Value $escapedValue -Confirm"
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
        New-PowerPointFile -Path $path
        Set-PowerPointFileProperty -Path $path -Name $name -Value $value
        (Get-PowerPointFileProperty -Path $path).Final | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with Name and Value' {
        New-PowerPointFile -Path $path
        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value
        (Get-PowerPointFileProperty -LiteralPath $path).Final | Should-BeTrue
      }
      It 'updates a file properties by Path with InputObject' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        Set-PowerPointFileProperty -Path $path -InputObject $properties
        (Get-PowerPointFileProperty -Path $path).Final | Should-BeTrue
      }
      It 'updates a file properties by LiteralPath with InputObject' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        Set-PowerPointFileProperty -LiteralPath $path -InputObject $properties
        (Get-PowerPointFileProperty -LiteralPath $path).Final | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-PowerPointFile -Path $path
        $path | Set-PowerPointFileProperty -Name $name -Value $value
        (Get-PowerPointFileProperty -Path $path).Final | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-PowerPointFileProperty -Name $name -Value $value
        (Get-PowerPointFileProperty -LiteralPath $path).Final | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-PowerPointFile -Path $path
        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf
        (Get-PowerPointFileProperty -LiteralPath $path).Final | Should-BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        New-PowerPointFile -Path $path
        $exitCode = 'N' | Invoke-Confirm -Path $path -Name $name -Value $value
        $exitCode | Should-Be 0
        (Get-PowerPointFileProperty -LiteralPath $path).Final | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        (Get-PowerPointFileProperty -LiteralPath $path -PasswordToOpen $password).Final | Should-BeTrue
      }
      It 'updates a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        (Get-PowerPointFileProperty -LiteralPath $path -PasswordToModify $password).Final | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-PowerPointFile -Path $path
        $property = Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.Final | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        $property = Set-PowerPointFileProperty -LiteralPath $path -InputObject $properties -PassThru
        $property | Should-NotBeNull
        $property.Final | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-PowerPointFile -Path $path
        { Set-PowerPointFileProperty -Path $path -Name 'NonExistentProperty' -Value 'Value' } | Should-Throw
      }
      It 'throws an error when read-only recommended file' {
        New-PowerPointFile -Path $path -ReadOnlyRecommended
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-PowerPointFileProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null

      $app = [PSCustomObject]@{}
      $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }

      $file = [PSCustomObject]@{ ReadOnly = $false }
      $file | Add-Member -MemberType ScriptMethod -Name Save -Value { }
      $file | Add-Member -MemberType ScriptMethod -Name Close -Value { }

      Mock -CommandName New-PowerPointObject -MockWith { $app }
      Mock -CommandName Open-PowerPointFile -MockWith { $file }
      Mock -CommandName Set-ObjectProperty
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call Open-PowerPointFile when WhatIf is specified' {
        Set-PowerPointFileProperty -LiteralPath $path -Name Final -Value $false -WhatIf
        Should-NotInvoke -CommandName Open-PowerPointFile
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        { Set-PowerPointFileProperty -LiteralPath $path -Name Final -Value $false } | Should-Throw
      }
    }
  }
  Describe 'Test-PowerPointExtension' {
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
    It 'returns true for *.pptx files by Path' {
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
      Mock -CommandName Get-Item -ParameterFilter { $Path -eq '*.pptx' } -MockWith {
        1, 2 |
        ForEach-Object {
          [PSCustomObject]@{
            PSIsContainer = $false
            PSPath        = ($env:TEMP | Join-Path -ChildPath "$_.pptx")
            FullName      = ($env:TEMP | Join-Path -ChildPath "$_.pptx")
          }
        }
      }
      Test-PowerPointExtension -Path '*.txt' | Should-BeFalse
      Test-PowerPointExtension -Path '*.pptx' | Should-BeTrue
    }
    It 'returns true for *.pptx files by LiteralPath' {
      Test-PowerPointExtension -LiteralPath 'Presentation.txt' | Should-BeFalse
      Test-PowerPointExtension -LiteralPath 'Presentation.pptx' | Should-BeTrue
    }
  }
  Describe 'Test-PowerPointExtension.Unit' {
    Context 'Edge cases' {
      It 'returns false when Get-Item throws ItemNotFoundException' {
        Mock -CommandName Get-Item -MockWith {
          throw [ItemNotFoundException]::new('not found')
        }

        Test-PowerPointExtension -Path '*.pptx' | Should-BeFalse
      }
      It 'returns false when matched items are not leaf paths' {
        Mock -CommandName Get-Item -MockWith {
          [PSCustomObject]@{
            FullName = ($env:TEMP | Join-Path -ChildPath 'Presentation.pptx')
          }
        }
        Mock -CommandName Test-Path -ParameterFilter { $IsValid } -MockWith { $true }
        Mock -CommandName Test-Path -ParameterFilter { $PathType -eq 'Leaf' } -MockWith { $false }

        Test-PowerPointExtension -LiteralPath 'Presentation.pptx' | Should-BeFalse
      }
    }
  }
}

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
        $item = New-PowerPointFile -Path $path

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $item = $path | New-PowerPointFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $item = [PSCustomObject]@{ FullName = $path } | New-PowerPointFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
    }
    Context 'SupportsShouldProcess' {
      It 'overwrites an existing presentation when Force is specified' {
        New-Item -Path $path -ItemType File | Out-Null

        New-PowerPointFile -Path $path -Force -Confirm

        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.pptm'
        $item = New-PowerPointFile -Path $path -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with PasswordToOpen' {
        $item = New-PowerPointFile -Path $path -PasswordToOpen $password

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with PasswordToModify' {
        $item = New-PowerPointFile -Path $path -PasswordToModify $password

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with ReadOnlyRecommended' {
        New-PowerPointFile -Path $path -ReadOnlyRecommended

        Get-PowerPointFileProperty -LiteralPath $path -Name ReadOnlyRecommended | Should-BeTrue
      }
      It 'creates a file with Initialize script block' {
        {
          New-PowerPointFile -Path $path -Initialize {
            param (
              [Microsoft.Office.Interop.PowerPoint.Presentation]
              $Presentation
            )
            $Presentation.Slides.Add(1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutTitleOnly) | Out-Null
          }
        } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-PowerPointFile' -Tag 'CI', 'Mock' {
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
        Mock -CommandName New-PowerPointObject -MockWith { throw }

        New-PowerPointFile -Path $path -Force -WhatIf

        Should-NotInvoke -CommandName New-PowerPointObject
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws when the file exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null

        { New-PowerPointFile -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Open-PowerPointFile' {
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
        New-PowerPointFile -Path $path

        Open-PowerPointFile -Path $path | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipeline' {
        New-PowerPointFile -Path $path

        $path | Open-PowerPointFile | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path

        [PSCustomObject]@{ FullName = $path } | Open-PowerPointFile | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a file with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password

        Open-PowerPointFile -Path $path -PasswordToOpen $password | Should-BeNull
      }
      It 'opens a file with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password

        Open-PowerPointFile -Path $path -PasswordToModify $password | Should-BeNull
      }
      It 'opens a file with ReadOnly' {
        New-PowerPointFile -Path $path

        Open-PowerPointFile -Path $path -ReadOnly | Should-BeNull
      }
      It 'opens a file with Action script block' {
        New-PowerPointFile -Path $path

        $value = Open-PowerPointFile -Path $path -Action {
          param (
            [Microsoft.Office.Interop.PowerPoint.Presentation]
            $Presentation
          )
          $Presentation.Slides.Add(1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutTitleOnly) | Out-Null
          return $Presentation.Slides.Count
        }

        $value | Should-Be 1
      }
      It 'opens a file using an existing Application object' {
        New-PowerPointFile -Path $path

        $app = New-PowerPointObject
        try {
          Open-PowerPointFile -Path $path -Application $app | Should-BeNull
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
  Describe 'Open-PowerPointFile' -Tag 'CI', 'Mock' {
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
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-PowerPointFile -Path $path

        $properties = Get-PowerPointFileProperty -Path $path -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by LiteralPath' {
        New-PowerPointFile -Path $path

        $properties = Get-PowerPointFileProperty -LiteralPath $path -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-PowerPointFile -Path $path

        $properties = $path | Get-PowerPointFileProperty -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path

        $properties = [PSCustomObject]@{ PSPath = $path } | Get-PowerPointFileProperty -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password

        { Get-PowerPointFileProperty -LiteralPath $path -Name FullName } | Should-Throw
        { Get-PowerPointFileProperty -LiteralPath $path -Name FullName -PasswordToOpen (Get-Password) } | Should-Throw
        { Get-PowerPointFileProperty -LiteralPath $path -Name FullName -PasswordToOpen $password } | Should -Not -Throw
      }
      It 'returns file properties from a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password

        { Get-PowerPointFileProperty -LiteralPath $path -Name FullName } | Should -Not -Throw
        { Get-PowerPointFileProperty -LiteralPath $path -Name FullName -PasswordToModify (Get-Password) } | Should -Not -Throw
        { Get-PowerPointFileProperty -LiteralPath $path -Name FullName -PasswordToModify $password } | Should -Not -Throw
      }
      It 'returns the property directly when only one name is provided' {
        New-PowerPointFile -Path $path

        Get-PowerPointFileProperty -LiteralPath $path -Name Final | Should-BeFalse
      }
      It 'returns file properties with Depth' {
        New-PowerPointFile -Path $path

        $properties = Get-PowerPointFileProperty -Path $path -Name CustomXMLParts -Depth 2

        $properties._NewEnum | Should-BeCollection -Count 3
      }
    }
  }
  Describe 'Get-PowerPointFileProperty' -Tag 'CI', 'Mock' {
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
  Describe 'Set-PowerPointFileProperty' {
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

        Get-PowerPointFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with Name and Value' {
        New-PowerPointFile -Path $path

        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value

        Get-PowerPointFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with InputObject' {
        New-PowerPointFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        Set-PowerPointFileProperty -Path $path -InputObject $properties

        Get-PowerPointFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with InputObject' {
        New-PowerPointFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        Set-PowerPointFileProperty -LiteralPath $path -InputObject $properties

        Get-PowerPointFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-PowerPointFile -Path $path

        $path | Set-PowerPointFileProperty -Name $name -Value $value

        Get-PowerPointFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path

        [PSCustomObject]@{ PSPath = $path } | Set-PowerPointFileProperty -Name $name -Value $value

        Get-PowerPointFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-PowerPointFile -Path $path

        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Get-PowerPointFileProperty -LiteralPath $path -Name $name | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password

        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password } | Should -Not -Throw
        Get-PowerPointFileProperty -LiteralPath $path -Name $name -PasswordToOpen $password | Should-BeTrue
      }
      It 'updates a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password

        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password } | Should -Not -Throw
        Get-PowerPointFileProperty -LiteralPath $path -Name $name -PasswordToModify $password | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-PowerPointFile -Path $path

        $property = Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -PassThru

        $property | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-PowerPointFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        $property = Set-PowerPointFileProperty -LiteralPath $path -InputObject $properties -PassThru

        $property | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-PowerPointFile -Path $path

        { Set-PowerPointFileProperty -LiteralPath $path -Name 'Dummy' -Value $value } | Should-Throw
      }
      It 'throws an error when read-only recommended file' {
        New-PowerPointFile -Path $path -ReadOnlyRecommended

        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-PowerPointFileProperty' -Tag 'CI', 'Mock' {
    BeforeEach {
      $path = Get-TempFile
      $name = 'Final'
      $value = $true
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
      It 'does not open the file when WhatIf is specified' {
        Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Should-NotInvoke -CommandName Open-PowerPointFile
      }
    }
    Context 'Edge cases' {
      It 'throws when the target is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-PowerPointFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        Should-NotInvoke -CommandName Open-PowerPointFile
      }
    }
  }
  Describe 'Test-PowerPointExtension' -Tag 'CI', 'Mock' {
    BeforeEach {
      Mock -CommandName Test-Path -MockWith { $true }
      Mock -CommandName Get-Item -MockWith {
        param (
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
    Context 'ParameterSetName' {
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

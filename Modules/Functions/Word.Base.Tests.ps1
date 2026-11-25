using assembly Microsoft.Office.Interop.Word
using assembly System.Web
using module .\..\Automation.Office.psd1
using namespace Microsoft.Office.Interop.Word
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
        $Extension = '.docx'
      )
      return $env:TEMP | Join-Path -ChildPath "Document.$([guid]::NewGuid().ToString('N'))$Extension"
    }
  }
  Describe 'New-WordFile' {
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
        $item = New-WordFile -Path $path

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $item = $path | New-WordFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $item = [PSCustomObject]@{ FullName = $path } | New-WordFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
    }
    Context 'SupportsShouldProcess' {
      It 'overwrites an existing document when Force is specified' {
        New-Item -Path $path -ItemType File | Out-Null

        New-WordFile -Path $path -Force -Confirm

        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.docm'
        $item = New-WordFile -Path $path -FileFormat wdFormatXMLDocumentMacroEnabled

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with PasswordToOpen' {
        $item = New-WordFile -Path $path -PasswordToOpen $password

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with PasswordToModify' {
        $item = New-WordFile -Path $path -PasswordToModify $password

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with ReadOnlyRecommended' {
        New-WordFile -Path $path -ReadOnlyRecommended

        Get-WordFileProperty -LiteralPath $path -Name ReadOnlyRecommended | Should-BeTrue
      }
      It 'creates a file with RemovePersonalInformation' {
        New-WordFile -Path $path -RemovePersonalInformation

        Get-WordFileProperty -LiteralPath $path -Name RemovePersonalInformation | Should-BeTrue
      }
      It 'creates a file with Initialize script block' {
        { New-WordFile -Path $path -Initialize {
            param (
              [Microsoft.Office.Interop.Word.Document]
              $Document
            )
            $Document.Range().Text = 'TestData'
          }
        } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-WordFile' -Tag 'CI', 'Mock' {
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
        Mock -CommandName New-WordObject -MockWith { throw }

        New-WordFile -Path $path -Force -WhatIf

        Should-NotInvoke -CommandName New-WordObject
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws when the file exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null

        { New-WordFile -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Open-WordFile' {
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
        New-WordFile -Path $path

        Open-WordFile -Path $path | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipeline' {
        New-WordFile -Path $path

        $path | Open-WordFile | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path

        [PSCustomObject]@{ FullName = $path } | Open-WordFile | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a file with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password

        Open-WordFile -Path $path -PasswordToOpen $password | Should-BeNull
      }
      It 'opens a file with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password

        Open-WordFile -Path $path -PasswordToModify $password | Should-BeNull
      }
      It 'opens a file with ReadOnly' {
        New-WordFile -Path $path

        Open-WordFile -Path $path -ReadOnly | Should-BeNull
      }
      It 'opens a file with Action script block' {
        New-WordFile -Path $path

        $value = Open-WordFile -Path $path -Action {
          param (
            [Microsoft.Office.Interop.Word.Document]
            $Document
          )
          $Document.Range().Text = 'TestContent'
          return $Document.Range().Text
        }

        $value | Should-BeString 'TestContent'
      }
      It 'opens a file using an existing Application object' {
        New-WordFile -Path $path

        $app = New-WordObject
        try {
          Open-WordFile -Path $path -Application $app | Should-BeNull
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
  Describe 'Open-WordFile' -Tag 'CI', 'Mock' {
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
  Describe 'Get-WordFileProperty' {
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
        New-WordFile -Path $path

        $properties = Get-WordFileProperty -Path $path -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by LiteralPath' {
        New-WordFile -Path $path

        $properties = Get-WordFileProperty -LiteralPath $path -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-WordFile -Path $path

        $properties = $path | Get-WordFileProperty -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path

        $properties = [PSCustomObject]@{ PSPath = $path } | Get-WordFileProperty -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password

        { Get-WordFileProperty -LiteralPath $path -Name FullName } | Should-Throw
        { Get-WordFileProperty -LiteralPath $path -Name FullName -PasswordToOpen (Get-Password) } | Should-Throw
        { Get-WordFileProperty -LiteralPath $path -Name FullName -PasswordToOpen $password } | Should -Not -Throw
      }
      It 'returns file properties from a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password

        { Get-WordFileProperty -LiteralPath $path -Name FullName } | Should -Not -Throw
        { Get-WordFileProperty -LiteralPath $path -Name FullName -PasswordToModify (Get-Password) } | Should -Not -Throw
        { Get-WordFileProperty -LiteralPath $path -Name FullName -PasswordToModify $password } | Should -Not -Throw
      }
      It 'returns the property directly when only one name is provided' {
        New-WordFile -Path $path

        Get-WordFileProperty -LiteralPath $path -Name Final | Should-BeFalse
      }
      It 'returns file properties with Depth' {
        New-WordFile -Path $path

        $properties = Get-WordFileProperty -Path $path -Name CustomXMLParts -Depth 2

        $properties._NewEnum | Should-BeCollection -Count 4
      }
    }
  }
  Describe 'Get-WordFileProperty' -Tag 'CI', 'Mock' {
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
  Describe 'Set-WordFileProperty' {
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
        New-WordFile -Path $path

        Set-WordFileProperty -Path $path -Name $name -Value $value

        Get-WordFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with Name and Value' {
        New-WordFile -Path $path

        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value

        Get-WordFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with InputObject' {
        New-WordFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        Set-WordFileProperty -Path $path -InputObject $properties

        Get-WordFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with InputObject' {
        New-WordFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        Set-WordFileProperty -LiteralPath $path -InputObject $properties

        Get-WordFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-WordFile -Path $path

        $path | Set-WordFileProperty -Name $name -Value $value

        Get-WordFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path

        [PSCustomObject]@{ PSPath = $path } | Set-WordFileProperty -Name $name -Value $value

        Get-WordFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-WordFile -Path $path

        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Get-WordFileProperty -LiteralPath $path -Name $name | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password

        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password } | Should -Not -Throw
        Get-WordFileProperty -LiteralPath $path -Name $name -PasswordToOpen $password | Should-BeTrue
      }
      It 'updates a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password

        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password } | Should -Not -Throw
        Get-WordFileProperty -LiteralPath $path -Name $name -PasswordToModify $password | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-WordFile -Path $path

        $property = Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PassThru

        $property | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-WordFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        $property = Set-WordFileProperty -LiteralPath $path -InputObject $properties -PassThru

        $property | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-WordFile -Path $path

        { Set-WordFileProperty -LiteralPath $path -Name 'Dummy' -Value $value } | Should-Throw
      }
      It 'throws an error when read-only recommended file' {
        New-WordFile -Path $path -ReadOnlyRecommended

        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-WordFileProperty' -Tag 'CI', 'Mock' {
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

      Mock -CommandName New-WordObject -MockWith { $app }
      Mock -CommandName Open-WordFile -MockWith { $file }
      Mock -CommandName Set-ObjectProperty
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not open the file when WhatIf is specified' {
        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Should-NotInvoke -CommandName Open-WordFile
      }
    }
    Context 'Edge cases' {
      It 'throws when the target is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        Should-NotInvoke -CommandName Open-WordFile
      }
    }
  }
  Describe 'Test-WordExtension' -Tag 'CI', 'Mock' {
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
      It 'returns true for *.docx files by Path' {
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
        Mock -CommandName Get-Item -ParameterFilter { $Path -eq '*.docx' } -MockWith {
          1, 2 |
          ForEach-Object {
            [PSCustomObject]@{
              PSIsContainer = $false
              PSPath        = ($env:TEMP | Join-Path -ChildPath "$_.docx")
              FullName      = ($env:TEMP | Join-Path -ChildPath "$_.docx")
            }
          }
        }

        Test-WordExtension -Path '*.txt' | Should-BeFalse
        Test-WordExtension -Path '*.docx' | Should-BeTrue
      }
      It 'returns true for *.docx files by LiteralPath' {
        Test-WordExtension -LiteralPath 'Document.txt' | Should-BeFalse
        Test-WordExtension -LiteralPath 'Document.docx' | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'returns false when Get-Item throws ItemNotFoundException' {
        Mock -CommandName Get-Item -MockWith {
          throw [ItemNotFoundException]::new('not found')
        }

        Test-WordExtension -Path '*.docx' | Should-BeFalse
      }
      It 'returns false when matched items are not leaf paths' {
        Mock -CommandName Get-Item -MockWith {
          [PSCustomObject]@{
            FullName = ($env:TEMP | Join-Path -ChildPath 'Document.docx')
          }
        }
        Mock -CommandName Test-Path -ParameterFilter { $IsValid } -MockWith { $true }
        Mock -CommandName Test-Path -ParameterFilter { $PathType -eq 'Leaf' } -MockWith { $false }

        Test-WordExtension -LiteralPath 'Document.docx' | Should-BeFalse
      }
    }
  }
}

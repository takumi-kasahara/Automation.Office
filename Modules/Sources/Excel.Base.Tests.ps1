using assembly Microsoft.Office.Interop.Excel
using assembly System.Web
using module .\..\Automation.Office.psd1
using namespace Microsoft.Office.Interop.Excel
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
        $Extension = '.xlsx'
      )
      return $env:TEMP | Join-Path -ChildPath "Book.$([guid]::NewGuid().ToString('N'))$Extension"
    }
  }
  Describe 'New-ExcelFile' {
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
        $item = New-ExcelFile -Path $path

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $item = $path | New-ExcelFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $item = [PSCustomObject]@{ FullName = $path } | New-ExcelFile

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
    }
    Context 'SupportsShouldProcess' {
      It 'overwrites an existing workbook when Force is specified' {
        New-Item -Path $path -ItemType File | Out-Null

        New-ExcelFile -Path $path -Force -Confirm

        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.xlsm'
        $item = New-ExcelFile -Path $path -FileFormat xlOpenXMLWorkbookMacroEnabled

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with PasswordToOpen' {
        $item = New-ExcelFile -Path $path -PasswordToOpen $password

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with PasswordToModify' {
        $item = New-ExcelFile -Path $path -PasswordToModify $password

        Test-Path -LiteralPath $path | Should-BeTrue
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
      }
      It 'creates a file with ReadOnlyRecommended' {
        New-ExcelFile -Path $path -ReadOnlyRecommended

        Get-ExcelFileProperty -LiteralPath $path -Name ReadOnlyRecommended | Should-BeTrue
      }
      It 'creates a file with RemovePersonalInformation' {
        $item = New-ExcelFile -Path $path -RemovePersonalInformation

        Get-ExcelFileProperty -LiteralPath $path -Name RemovePersonalInformation | Should-BeTrue
      }
      It 'creates a file with Initialize script block' {
        {
          New-ExcelFile -Path $path -Initialize {
            param (
              [Microsoft.Office.Interop.Excel.Workbook]
              $Workbook
            )
            $Workbook.Worksheets.Item(1).Name = 'TestData'
          }
        } | Should -Not -Throw
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-ExcelFile' -Tag 'CI', 'Mock' {
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
        Mock -CommandName New-ExcelObject -MockWith { throw }

        New-ExcelFile -Path $path -Force -WhatIf

        Should-NotInvoke -CommandName New-ExcelObject
        Test-Path -LiteralPath $path | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws when the file exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null

        { New-ExcelFile -Path $path } | Should-Throw
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
      It 'opens a file by Path' {
        New-ExcelFile -Path $path

        Open-ExcelFile -Path $path | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipeline' {
        New-ExcelFile -Path $path

        $path | Open-ExcelFile | Should-BeNull
      }
      It 'opens a file by Path with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path

        [PSCustomObject]@{ FullName = $path } | Open-ExcelFile | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a file with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password

        Open-ExcelFile -Path $path -PasswordToOpen $password | Should-BeNull
      }
      It 'opens a file with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password

        Open-ExcelFile -Path $path -PasswordToModify $password | Should-BeNull
      }
      It 'opens a file with ReadOnly' {
        New-ExcelFile -Path $path

        Open-ExcelFile -Path $path -ReadOnly | Should-BeNull
      }
      It 'opens a file with Force' {
        New-ExcelFile -Path $path -ReadOnlyRecommended

        Open-ExcelFile -Path $path -Force | Should-BeNull
      }
      It 'opens a file with Action script block' {
        New-ExcelFile -Path $path

        $value = Open-ExcelFile -Path $path -Action {
          param (
            [Microsoft.Office.Interop.Excel.Workbook]
            $Workbook
          )
          $Workbook.Worksheets.Item(1).Name = 'Sheet1'
          return $Workbook.Worksheets.Item(1).Name
        }

        $value | Should-BeString 'Sheet1'
      }
      It 'opens a file using an existing Application object' {
        New-ExcelFile -Path $path

        $app = New-ExcelObject
        try {
          Open-ExcelFile -Application $app -Path $path | Should-BeNull
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
  Describe 'Open-ExcelFile' -Tag 'CI', 'Mock' {
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
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-ExcelFile -Path $path

        $properties = Get-ExcelFileProperty -Path $path -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by LiteralPath' {
        New-ExcelFile -Path $path

        $properties = Get-ExcelFileProperty -LiteralPath $path -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-ExcelFile -Path $path

        $properties = $path | Get-ExcelFileProperty -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path

        $properties = [PSCustomObject]@{ PSPath = $path } | Get-ExcelFileProperty -Name Final, FullName

        @($properties.PSObject.Properties).Count | Should-Be 2
        $properties.Final | Should-BeFalse
        $properties.FullName | Should-BeString $path
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password

        { Get-ExcelFileProperty -LiteralPath $path -Name FullName } | Should-Throw
        { Get-ExcelFileProperty -LiteralPath $path -Name FullName -PasswordToOpen (Get-Password) } | Should-Throw
        { Get-ExcelFileProperty -LiteralPath $path -Name FullName -PasswordToOpen $password } | Should -Not -Throw
      }
      It 'returns file properties from a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password

        { Get-ExcelFileProperty -LiteralPath $path -Name FullName } | Should -Not -Throw
        { Get-ExcelFileProperty -LiteralPath $path -Name FullName -PasswordToModify (Get-Password) } | Should -Not -Throw
        { Get-ExcelFileProperty -LiteralPath $path -Name FullName -PasswordToModify $password } | Should -Not -Throw
      }
      It 'returns the property directly when only one name is provided' {
        New-ExcelFile -Path $path

        Get-ExcelFileProperty -LiteralPath $path -Name Final | Should-BeFalse
      }
      It 'returns file properties with Depth' {
        New-ExcelFile -Path $path

        $properties = Get-ExcelFileProperty -Path $path -Name Worksheets -Depth 2

        $properties._NewEnum | Should-HaveType ([PSCustomObject])
      }
    }
  }
  Describe 'Get-ExcelFileProperty' -Tag 'CI', 'Mock' {
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
  Describe 'Set-ExcelFileProperty' {
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

        Get-ExcelFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with Name and Value' {
        New-ExcelFile -Path $path

        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value

        Get-ExcelFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with InputObject' {
        New-ExcelFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        Set-ExcelFileProperty -Path $path -InputObject $properties

        Get-ExcelFileProperty -Path $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with InputObject' {
        New-ExcelFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        Set-ExcelFileProperty -LiteralPath $path -InputObject $properties

        Get-ExcelFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-ExcelFile -Path $path

        $path | Set-ExcelFileProperty -Name $name -Value $value

        Get-ExcelFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path

        [PSCustomObject]@{ PSPath = $path } | Set-ExcelFileProperty -Name $name -Value $value

        Get-ExcelFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-ExcelFile -Path $path

        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -Force -WhatIf

        Get-ExcelFileProperty -LiteralPath $path -Name $name | Should-BeFalse
      }
      It 'updates a read-only recommended workbook when Force is specified' {
        New-ExcelFile -Path $path -ReadOnlyRecommended

        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -Force -Confirm

        Get-ExcelFileProperty -LiteralPath $path -Name $name | Should-BeTrue
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password

        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password } | Should -Not -Throw
        Get-ExcelFileProperty -LiteralPath $path -Name $name -PasswordToOpen $password | Should-BeTrue
      }
      It 'updates a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password

        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password } | Should -Not -Throw
        Get-ExcelFileProperty -LiteralPath $path -Name $name -PasswordToModify $password | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-ExcelFile -Path $path

        $property = Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -PassThru

        $property | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-ExcelFile -Path $path

        $properties = [PSCustomObject]@{ $name = $value }
        $property = Set-ExcelFileProperty -LiteralPath $path -InputObject $properties -PassThru

        $property | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-ExcelFile -Path $path

        { Set-ExcelFileProperty -LiteralPath $path -Name 'Dummy' -Value $value } | Should-Throw
      }
      It 'throws an error when read-only recommended file without Force' {
        New-ExcelFile -Path $path -ReadOnlyRecommended

        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-ExcelFileProperty' -Tag 'CI', 'Mock' {
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

      Mock -CommandName New-ExcelObject -MockWith { $app }
      Mock -CommandName Open-ExcelFile -MockWith { $file }
      Mock -CommandName Set-ObjectProperty
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not open the file when WhatIf is specified' {
        Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf

        Should-NotInvoke -CommandName Open-ExcelFile
      }
    }
    Context 'Edge cases' {
      It 'throws when the target is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-ExcelFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        Should-NotInvoke -CommandName Open-ExcelFile
      }
    }
  }
  Describe 'Test-ExcelExtension' -Tag 'CI', 'Mock' {
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

        Test-ExcelExtension -Path '*.txt' | Should-BeFalse
        Test-ExcelExtension -Path '*.xlsx' | Should-BeTrue
      }
      It 'returns true for *.xlsx files by LiteralPath' {
        Test-ExcelExtension -LiteralPath 'Book.txt' | Should-BeFalse
        Test-ExcelExtension -LiteralPath 'Book.xlsx' | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'returns false when Get-Item throws ItemNotFoundException' {
        Mock -CommandName Get-Item -MockWith {
          throw [ItemNotFoundException]::new('not found')
        }

        Test-ExcelExtension -Path '*.xlsx' | Should-BeFalse
      }
      It 'returns false when matched items are not leaf paths' {
        Mock -CommandName Get-Item -MockWith {
          [PSCustomObject]@{
            FullName = ($env:TEMP | Join-Path -ChildPath 'Book.xlsx')
          }
        }
        Mock -CommandName Test-Path -ParameterFilter { $IsValid } -MockWith { $true }
        Mock -CommandName Test-Path -ParameterFilter { $PathType -eq 'Leaf' } -MockWith { $false }

        Test-ExcelExtension -LiteralPath 'Book.xlsx' | Should-BeFalse
      }
    }
  }
}

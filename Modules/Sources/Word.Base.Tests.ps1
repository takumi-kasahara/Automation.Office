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
            "New-WordFile -Path '$escapedPath' -Confirm"
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
        $item = New-WordFile -Path $path
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file by Path with ValueFromPipeline' {
        $path = Get-TempFile
        $item = $path | New-WordFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file by Path with ValueFromPipelineByPropertyName' {
        $path = Get-TempFile
        $item = [PSCustomObject]@{ FullName = $path } | New-WordFile
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create a file when WhatIf is specified' {
        $path = Get-TempFile
        New-WordFile -Path $path -Force -WhatIf
      }
      It 'asks for confirmation when Confirm is specified' {
        $path = Get-TempFile
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should-Be 0
      }
      It 'overwrites an existing document when Force is specified' {
        $path = Get-TempFile
        New-Item -Path $path -ItemType File | Out-Null
        New-WordFile -Path $path -Force -Confirm
        (Get-Item -LiteralPath $path).Length | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'creates a file using the requested file format' {
        $path = Get-TempFile -Extension '.docm'
        $item = New-WordFile -Path $path -FileFormat wdFormatXMLDocumentMacroEnabled
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
      }
      It 'creates a file with PasswordToOpen' {
        $path = Get-TempFile
        New-WordFile -Path $path -PasswordToOpen $password
        Test-Path -LiteralPath $path | Should-BeTrue
        $app = New-WordObject
        try {
          Open-WordFile -Application $app -Path $path -PasswordToOpen $password -Action {
            param (
              [Microsoft.Office.Interop.Word.Document]
              $Document
            )
            $Document.HasPassword | Should-BeTrue
            $Document.Password | Should-NotBeEmptyString
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
      It 'creates a file with PasswordToModify' {
        $path = Get-TempFile
        New-WordFile -Path $path -PasswordToModify $password
        Test-Path -LiteralPath $path | Should-BeTrue
        $app = New-WordObject
        try {
          Open-WordFile -Application $app -Path $path -PasswordToModify $password -Action {
            param (
              [Microsoft.Office.Interop.Word.Document]
              $Document
            )
            $Document.WritePassword | Should-NotBeEmptyString
            $Document.WriteReserved | Should-BeTrue
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
      It 'creates a file with ReadOnlyRecommended' {
        $path = Get-TempFile
        New-WordFile -Path $path -ReadOnlyRecommended
        Test-Path -LiteralPath $path | Should-BeTrue
        $app = New-WordObject
        try {
          Open-WordFile -Application $app -Path $path -Action {
            param (
              [Microsoft.Office.Interop.Word.Document]
              $Document
            )
            $Document.ReadOnlyRecommended | Should-BeTrue
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
      It 'creates a file with RemovePersonalInformation' {
        $path = Get-TempFile
        New-WordFile -Path $path -RemovePersonalInformation
        Test-Path -LiteralPath $path | Should-BeTrue
        $app = New-WordObject
        try {
          Open-WordFile -Application $app -Path $path -Action {
            param (
              [Microsoft.Office.Interop.Word.Document]
              $Document
            )
            $Document.RemovePersonalInformation | Should-BeTrue
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
      It 'creates a file with Initialize script block' {
        $path = Get-TempFile
        $item = New-WordFile -Path $path -Initialize {
          param (
            [Microsoft.Office.Interop.Word.Document]
            $Document
          )
          $Document.Range().Text = 'TestData'
        }
        $item | Should-HaveType ([System.IO.FileInfo])
        $item.FullName | Should-Be ([Path]::GetFullPath($path))
        Test-Path -LiteralPath $path | Should-BeTrue
        $app = New-WordObject
        try {
          Open-WordFile -Application $app -Path $path -Action {
            param (
              [Microsoft.Office.Interop.Word.Document]
              $Document
            )
            $Document.Range().Text | Should-Be 'TestData'
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
        { New-WordFile -Path $path } | Should-Throw
      }
    }
  }
  Describe 'New-WordFile.Unit' {
    BeforeEach {
      $path = Get-TempFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call New-WordObject when WhatIf is specified' {
        Mock -CommandName New-WordObject

        { New-WordFile -Path $path -Force -WhatIf } | Should -Not -Throw
        Should-NotInvoke -CommandName New-WordObject
      }
    }
    Context 'Edge cases' {
      It 'throws and does not call New-WordObject when path exists and Force is not specified' {
        New-Item -Path $path -ItemType File -Force | Out-Null
        Mock -CommandName New-WordObject

        { New-WordFile -Path $path } | Should-Throw
        Should-NotInvoke -CommandName New-WordObject
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
      It 'opens a document by Path and returns the document object' {
        New-WordFile -Path $path
        Open-WordFile -Path $path | Should-NotBeNull
      }
      It 'opens a document by Path with ValueFromPipeline' {
        New-WordFile -Path $path
        $path | Open-WordFile | Should-NotBeNull
      }
      It 'opens a document by Path with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path
        [PSCustomObject]@{ FullName = $path } | Open-WordFile | Should-NotBeNull
      }
    }
    Context 'Other parameters' {
      It 'opens a document with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password
        Open-WordFile -Path $path -PasswordToOpen $password | Should-NotBeNull
      }
      It 'opens a document with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password
        Open-WordFile -Path $path -PasswordToModify $password | Should-NotBeNull
      }
      It 'opens a document with ReadOnly' {
        New-WordFile -Path $path
        Open-WordFile -Path $path -ReadOnly | Should-NotBeNull
      }
      It 'opens a document with Action script block' {
        New-WordFile -Path $path
        $result = Open-WordFile -Path $path -Action {
          param(
            [Microsoft.Office.Interop.Word.Document]
            $Document
          )
          $Document.Range().Text = 'TestContent'
          return $Document.Range().Text
        }
        $result | Should-Be 'TestContent'
      }
      It 'opens a document using an existing Application object' {
        New-WordFile -Path $path
        $app = New-WordObject
        try {
          Open-WordFile -Path $path -Application $app | Should-NotBeNull
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
  Describe 'Open-WordFile.Unit' {
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
      It 'throws when New-WordObject fails' {
        Mock -CommandName New-WordObject -MockWith { throw }

        { Open-WordFile -Path $path } | Should-Throw
        Should-Invoke -CommandName New-WordObject -Times 1 -Exactly
      }
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
    It 'returns file properties' {
      New-WordFile -Path $path
      { Get-WordFileProperty -Path $path | Out-Host } | Should -Not -Throw
    }
    Context 'ParameterSetName' {
      It 'returns selected file properties by Path' {
        New-WordFile -Path $path
        $properties = Get-WordFileProperty -Path $path -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
      It 'returns selected file properties by LiteralPath' {
        New-WordFile -Path $path
        $properties = Get-WordFileProperty -LiteralPath $path -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
      It 'returns selected file properties by Path with ValueFromPipeline' {
        New-WordFile -Path $path
        $properties = $path | Get-WordFileProperty -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
      It 'returns selected file properties by Path with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-WordFileProperty -Name Final
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
    }
    Context 'Other parameters' {
      It 'returns file properties from a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password
        { Get-WordFileProperty -Path $path } | Should-Throw
        { Get-WordFileProperty -Path $path -PasswordToOpen (Get-Password) } | Should-Throw
        { Get-WordFileProperty -Path $path -PasswordToOpen $password } | Should -Not -Throw
      }
      It 'returns file properties from a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password
        { Get-WordFileProperty -Path $path } | Should -Not -Throw
        { Get-WordFileProperty -Path $path -PasswordToModify (Get-Password) } | Should -Not -Throw
        { Get-WordFileProperty -Path $path -PasswordToModify $password } | Should -Not -Throw
      }
    }
  }
  Describe 'Get-WordFileProperty.Unit' {
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
      It 'throws when New-WordObject fails' {
        Mock -CommandName New-WordObject -MockWith { throw }

        { Get-WordFileProperty -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Set-WordFileProperty' {
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
            "Set-WordFileProperty -LiteralPath '$escapedPath' -Name '$escapedName' -Value $escapedValue -Confirm"
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
        New-WordFile -Path $path
        Set-WordFileProperty -Path $path -Name $name -Value $value
        (Get-WordFileProperty -Path $path).Final | Should-BeTrue
      }
      It 'updates a file property by LiteralPath with Name and Value' {
        New-WordFile -Path $path
        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value
        (Get-WordFileProperty -LiteralPath $path).Final | Should-BeTrue
      }
      It 'updates a file properties by Path with InputObject' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        Set-WordFileProperty -Path $path -InputObject $properties
        (Get-WordFileProperty -Path $path).Final | Should-BeTrue
      }
      It 'updates a file properties by LiteralPath with InputObject' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        Set-WordFileProperty -LiteralPath $path -InputObject $properties
        (Get-WordFileProperty -LiteralPath $path).Final | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipeline' {
        New-WordFile -Path $path
        $path | Set-WordFileProperty -Name $name -Value $value
        (Get-WordFileProperty -Path $path).Final | Should-BeTrue
      }
      It 'updates a file property by Path with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-WordFileProperty -Name $name -Value $value
        (Get-WordFileProperty -LiteralPath $path).Final | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-WordFile -Path $path
        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -WhatIf
        (Get-WordFileProperty -LiteralPath $path).Final | Should-BeFalse
      }
      It 'asks for confirmation when Confirm is specified' {
        New-WordFile -Path $path
        $exitCode = 'N' | Invoke-Confirm -Path $path -Name $name -Value $value
        $exitCode | Should-Be 0
        (Get-WordFileProperty -LiteralPath $path).Final | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        (Get-WordFileProperty -LiteralPath $path -PasswordToOpen $password).Final | Should-BeTrue
      }
      It 'updates a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        (Get-WordFileProperty -LiteralPath $path -PasswordToModify $password).Final | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified' {
        New-WordFile -Path $path
        $property = Set-WordFileProperty -LiteralPath $path -Name $name -Value $value -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.Final | Should-BeTrue
      }
      It 'returns updated file properties when PassThru is specified with InputObject' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{ Final = $true }
        $property = Set-WordFileProperty -LiteralPath $path -InputObject $properties -PassThru
        $property | Should-NotBeNull
        $property.Final | Should-BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when trying to set a property that does not exist.' {
        New-WordFile -Path $path
        { Set-WordFileProperty -Path $path -Name 'NonExistentProperty' -Value 'Value' } | Should-Throw
      }
      It 'throws an error when read-only recommended file' {
        New-WordFile -Path $path -ReadOnlyRecommended
        { Set-WordFileProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-WordFileProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
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
      It 'does not call Open-WordFile when WhatIf is specified' {
        Set-WordFileProperty -LiteralPath $path -Name Final -Value $true -WhatIf
        Should-NotInvoke -CommandName Open-WordFile
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        { Set-WordFileProperty -LiteralPath $path -Name Final -Value $true } | Should-Throw
      }
    }
  }
  Describe 'Test-WordExtension' {
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
      Test-WordExtension -Path '*.docx' | Should-BeTrue
    }
    It 'returns true for *.docx files by LiteralPath' {
      Test-WordExtension -LiteralPath 'Document.docx' | Should-BeTrue
    }
  }
  Describe 'Test-WordExtension.Unit' {
    Context 'Edge cases' {
      It 'returns false when Get-Item throws ItemNotFoundException' {
        Mock -CommandName Get-Item -MockWith {
          throw [ItemNotFoundException]::new('not found')
        }

      }
      It 'returns false when matched items are not leaf paths' {
        Mock -CommandName Get-Item -MockWith {
          [PSCustomObject]@{
            FullName = ($env:TEMP | Join-Path -ChildPath 'Document.docx')
          }
        }
        Mock -CommandName Test-Path -ParameterFilter { $IsValid } -MockWith { $true }
        Mock -CommandName Test-Path -ParameterFilter { $PathType -eq 'Leaf' } -MockWith { $false }

      }
    }
  }
}

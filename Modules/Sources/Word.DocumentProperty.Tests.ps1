using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Security

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'
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
    function Get-TempFileSet {
      [CmdletBinding()]
      [OutputType([PSCustomObject])]
      param ()
      $parent = $env:TEMP | Join-Path -ChildPath ("Document.$([guid]::NewGuid().ToString('N'))")
      New-Item -Path $parent -ItemType Directory | Out-Null
      return [PSCustomObject]@{
        Parent    = $parent
        ChildItem = @(
          ($parent | Join-Path -ChildPath '1.docx')
          ($parent | Join-Path -ChildPath '2.docx')
        )
        Path      = ($parent | Join-Path -ChildPath '*.docx')
      }
    }
  }
  Describe 'Get-WordDocumentProperty' {
    BeforeEach {
      $password = Get-Password
      $path = Get-TempFile
      $pathSet = Get-TempFileSet
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $pathSet.Parent) {
        Remove-Item -LiteralPath $pathSet.Parent -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'reads built-in properties by Path' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        $properties = @(Get-WordDocumentProperty -Path $pathSet.Path)
        $properties.Count | Should -Be 2
        $properties | ForEach-Object { @($_.PSObject.Properties).Count | Should -BeGreaterThan 0 }
      }
      It 'reads built-in properties by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        $properties = @($pathSet.Path | Get-WordDocumentProperty)
        $properties.Count | Should -Be 2
        $properties | ForEach-Object { @($_.PSObject.Properties).Count | Should -BeGreaterThan 0 }
      }
      It 'reads built-in properties by LiteralPath' {
        New-WordFile -Path $path
        $properties = Get-WordDocumentProperty -LiteralPath $path
        @($properties.PSObject.Properties).Count | Should -BeGreaterThan 0
      }
      It 'reads built-in properties by LiteralPath with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-WordDocumentProperty
        @($properties.PSObject.Properties).Count | Should -BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'filters properties by Name' {
        New-WordFile -Path $path
        $allProperties = Get-WordDocumentProperty -Path $path
        $name = @($allProperties.PSObject.Properties.Name) | Select-Object -First 1
        $properties = Get-WordDocumentProperty -Path $path -Name $name
        @($properties.PSObject.Properties.Name) | Should -Contain $name
        @($properties.PSObject.Properties).Count | Should -Be 1
      }
      It 'reads properties from a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password
        { Get-WordDocumentProperty -LiteralPath $path -PasswordToOpen (Get-Password) } | Should -Throw
        $properties = Get-WordDocumentProperty -LiteralPath $path -PasswordToOpen $password
        @($properties.PSObject.Properties).Count | Should -BeGreaterThan 0
      }
      It 'reads properties from a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify (Get-Password)
        $properties = Get-WordDocumentProperty -LiteralPath $path -PasswordToModify (Get-Password)
        @($properties.PSObject.Properties).Count | Should -BeGreaterThan 0
      }
      It 'does not return built-in properties when Custom is specified' {
        New-WordFile -Path $path
        $properties = Get-WordDocumentProperty -LiteralPath $path -Custom
        @($properties.PSObject.Properties).Count | Should -Be 0
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Get-WordDocumentProperty.Unit' {
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
        Mock -CommandName New-WordObject -MockWith { throw 'new word object failed' }

        { Get-WordDocumentProperty -Path $path } | Should -Throw
      }
    }
  }
  Describe 'Get-WordPropertyValue.Unit' {
    BeforeAll {
      Mock -CommandName Test-Path -MockWith { $true }
    }
    Context 'ParameterSetName' {
      It 'returns null when Get-WordDocumentProperty returns null' {
        Mock -CommandName Get-WordDocumentProperty

        $result = Get-WordPropertyValue -LiteralPath 'dummy.docx' -Name Title
        $result | Should -BeNullOrEmpty
      }
      It 'returns property values from Get-WordDocumentProperty output' {
        Mock -CommandName Get-WordDocumentProperty -MockWith {
          [PSCustomObject]@{
            Title = 'A'
          },
          [PSCustomObject]@{
            Title = 'B'
          }
        }

        $result = @(Get-WordPropertyValue -LiteralPath 'dummy.docx' -Name Title)
        $result.Count | Should -Be 2
        $result | Should -Be @('A', 'B')
      }
    }
  }
  Describe 'Set-WordDocumentProperty' {
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
          [string]
          $Value
        )
        process {
          $modulePath = [Path]::GetFullPath(($PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psd1'))
          $escapedModulePath = $modulePath.Replace("'", "''")
          $escapedPath = $Path.Replace("'", "''")
          $escapedName = $Name.Replace("'", "''")
          $escapedValue = $Value.Replace("'", "''")
          $command = @(
            "Import-Module '$escapedModulePath' -Force"
            "Set-WordDocumentProperty -LiteralPath '$escapedPath' -Name '$escapedName' -Value '$escapedValue' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $name = 'Title'
      $value = [guid]::NewGuid().ToString('N')
      $password = Get-Password
      $path = Get-TempFile
      $pathSet = Get-TempFileSet
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $pathSet.Parent) {
        Remove-Item -LiteralPath $pathSet.Parent -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'sets a built-in property by Path' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        Set-WordDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        $values = @(Get-WordPropertyValue -Path $pathSet.Path -Name $name)
        $values.Count | Should -Be $pathSet.ChildItem.Count
        $values | Should -Be (@($value) * $pathSet.ChildItem.Count)
      }
      It 'sets a built-in property by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        $pathSet.Path | Set-WordDocumentProperty -Name $name -Value $value
        $values = @(Get-WordPropertyValue -Path $pathSet.Path -Name $name)
        $values.Count | Should -Be $pathSet.ChildItem.Count
        $values | Should -Be (@($value) * $pathSet.ChildItem.Count)
      }
      It 'sets multiple built-in properties by Path with InputObject' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        Set-WordDocumentProperty -Path $pathSet.Path -InputObject $properties
        @(Get-WordPropertyValue -Path $pathSet.Path -Name 'Title') | Should -Be (@($properties.Title) * $pathSet.ChildItem.Count)
        @(Get-WordPropertyValue -Path $pathSet.Path -Name 'Subject') | Should -Be (@($properties.Subject) * $pathSet.ChildItem.Count)
      }
      It 'sets a built-in property by LiteralPath' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -Be $value
      }
      It 'sets a built-in property by LiteralPath with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-WordDocumentProperty -Name $name -Value $value
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -Be $value
      }
      It 'sets multiple built-in properties by LiteralPath with InputObject' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        Set-WordDocumentProperty -LiteralPath $path -InputObject $properties
        Get-WordPropertyValue -LiteralPath $path -Name 'Title' | Should -Be $properties.Title
        Get-WordPropertyValue -LiteralPath $path -Name 'Subject' | Should -Be $properties.Subject
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value 'ante value'
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value 'post value' -Force -WhatIf
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -Be 'ante value'
      }
      It 'asks for confirmation when Confirm is specified' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value 'ante confirm'
        $exitCode = 'N' | Invoke-Confirm -Path $path -Name $name -Value 'post confirm'
        $exitCode | Should -Be 0
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -Be 'ante confirm'
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password
        { Set-WordDocumentProperty -LiteralPath $path -Name $name -Value ([guid]::NewGuid().ToString('N')) -PasswordToOpen (Get-Password) } | Should -Throw
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        Get-WordPropertyValue -LiteralPath $path -Name $name -PasswordToOpen $password | Should -Be $value
      }
      It 'updates a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password
        { Set-WordDocumentProperty -LiteralPath $path -Name $name -Value ([guid]::NewGuid().ToString('N')) -PasswordToModify (Get-Password) } | Should -Throw
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        Get-WordPropertyValue -LiteralPath $path -Name $name -PasswordToModify $password | Should -Be $value
      }
      It 'returns the updated built-in property when PassThru is specified' {
        New-WordFile -Path $path
        $property = Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -PassThru
        $property | Should -BeOfType [PSCustomObject]
        $property.$name | Should -Be $value
      }
      It 'returns updated properties when PassThru is specified with InputObject' {
        New-WordFile -Path $path
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        $property = Set-WordDocumentProperty -LiteralPath $path -InputObject $properties -PassThru
        $property | Should -BeOfType [PSCustomObject]
        $property.Title | Should -Be $properties.Title
        $property.Subject | Should -Be $properties.Subject
      }
      It 'sets a custom property' {
        New-WordFile -Path $path
        $name = 'MyProperty'
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -Custom
        Get-WordPropertyValue -LiteralPath $path -Name $name -Custom | Should -Be $value
      }
      It 'returns the updated custom property when PassThru is specified' {
        New-WordFile -Path $path
        $name = 'MyProperty'
        $property = Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -Custom -PassThru
        $property | Should -BeOfType [PSCustomObject]
        $property.$name | Should -Be $value
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read-only recommended document' {
        New-WordFile -Path $path -ReadOnlyRecommended
        { Set-WordDocumentProperty -LiteralPath $path -Name 'Title' -Value ([guid]::NewGuid().ToString('N')) } | Should -Throw
      }
    }
  }
  Describe 'Set-WordDocumentProperty.Unit' {
    BeforeEach {
      $name = 'Title'
      $value = [guid]::NewGuid().ToString('N')
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null
      $script:called = 0

      Mock -CommandName New-WordObject -MockWith {
        $app = [PSCustomObject]@{}
        $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
        return $app
      }
      Mock -CommandName Open-WordFile -MockWith {
        $script:called++
        throw 'must not be called in this test'
      }
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call Open-WordFile when WhatIf is specified' {
        { Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -WhatIf } | Should -Not -Throw
        $script:called | Should -Be 0
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should -Throw
        $script:called | Should -Be 0
      }
      It 'throws when New-WordObject fails' {
        Mock -CommandName New-WordObject -MockWith { throw 'new word object failed' }

        { Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should -Throw
      }
    }
  }
  Describe 'Remove-WordDocumentProperty' {
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
            "Import-Module '$escapedModulePath' -Force"
            "Remove-WordDocumentProperty -LiteralPath '$escapedPath' -Confirm"
          ) -join '; '
          @($Response) | & powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -Command $command | Out-Host
          return $LASTEXITCODE
        }
      }
    }
    BeforeEach {
      $name = 'Title'
      $value = [guid]::NewGuid().ToString('N')
      $password = Get-Password
      $path = Get-TempFile
      $pathSet = Get-TempFileSet
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $pathSet.Parent) {
        Remove-Item -LiteralPath $pathSet.Parent -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'removes built-in properties by Path' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        Set-WordDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        Remove-WordDocumentProperty -Path $pathSet.Path
        $values = Get-WordPropertyValue -Path $pathSet.Path -Name $name
        $values.Count | Should -Be $pathSet.ChildItem.Count
        $values | ForEach-Object { $_ | Should -BeNullOrEmpty }
      }
      It 'removes built-in properties by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-WordFile -Path $_ }
        Set-WordDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        $pathSet.Path | Remove-WordDocumentProperty
        $values = Get-WordPropertyValue -Path $pathSet.Path -Name $name
        $values.Count | Should -Be $pathSet.ChildItem.Count
        $values | ForEach-Object { $_ | Should -BeNullOrEmpty }
      }
      It 'removes built-in properties by LiteralPath' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-WordDocumentProperty -LiteralPath $path
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -BeNullOrEmpty
      }
      It 'removes built-in properties by LiteralPath with ValueFromPipelineByPropertyName' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value
        [PSCustomObject]@{ PSPath = $path } | Remove-WordDocumentProperty
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -BeNullOrEmpty
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not remove properties when WhatIf is specified' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-WordDocumentProperty -LiteralPath $path -Force -WhatIf
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -Be $value
      }
      It 'asks for confirmation when Confirm is specified' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should -Be 0
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -Be $value
      }
    }
    Context 'Other parameters' {
      It 'removes properties from a file protected with PasswordToOpen' {
        New-WordFile -Path $path -PasswordToOpen $password
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        { Remove-WordDocumentProperty -LiteralPath $path -PasswordToOpen (Get-Password) } | Should -Throw
        Remove-WordDocumentProperty -LiteralPath $path -PasswordToOpen $password
        Get-WordPropertyValue -LiteralPath $path -Name $name -PasswordToOpen $password | Should -BeNullOrEmpty
      }
      It 'removes properties from a file protected with PasswordToModify' {
        New-WordFile -Path $path -PasswordToModify $password
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        { Remove-WordDocumentProperty -LiteralPath $path -PasswordToModify (Get-Password) } | Should -Throw
        Remove-WordDocumentProperty -LiteralPath $path -PasswordToModify $password
        Get-WordPropertyValue -LiteralPath $path -Name $name -PasswordToModify $password | Should -BeNullOrEmpty
      }
      It 'uses RemoveDocInfoType when removing properties' {
        New-WordFile -Path $path
        Set-WordDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-WordDocumentProperty -LiteralPath $path -RemoveDocInfoType wdRDIAll
        Get-WordPropertyValue -LiteralPath $path -Name $name | Should -BeNullOrEmpty
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read- la-only recommended document' {
        New-WordFile -Path $path -ReadOnlyRecommended
        { Remove-WordDocumentProperty -LiteralPath $path } | Should -Throw
      }
    }
  }
  Describe 'Remove-WordDocumentProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null
      $script:called = 0

      Mock -CommandName New-WordObject -MockWith {
        $app = [PSCustomObject]@{}
        $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
        return $app
      }
      Mock -CommandName Open-WordFile -MockWith {
        $script:called++
        throw 'must not be called in this test'
      }
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call Open-WordFile when WhatIf is specified' {
        { Remove-WordDocumentProperty -LiteralPath $path -WhatIf } | Should -Not -Throw
        $script:called | Should -Be 0
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Remove-WordDocumentProperty -LiteralPath $path } | Should -Throw
        $script:called | Should -Be 0
      }
      It 'throws when New-WordObject fails' {
        Mock -CommandName New-WordObject -MockWith { throw 'new word object failed' }

        { Remove-WordDocumentProperty -LiteralPath $path } | Should -Throw
      }
    }
  }
}

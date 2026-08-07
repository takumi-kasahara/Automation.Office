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
        $Extension = '.xlsx'
      )
      return $env:TEMP | Join-Path -ChildPath "Book.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function Get-TempFileSet {
      [CmdletBinding()]
      [OutputType([PSCustomObject])]
      param ()
      $parent = $env:TEMP | Join-Path -ChildPath ("Book.$([guid]::NewGuid().ToString('N'))")
      New-Item -Path $parent -ItemType Directory | Out-Null
      return [PSCustomObject]@{
        Parent    = $parent
        ChildItem = @(
          ($parent | Join-Path -ChildPath '1.xlsx')
          ($parent | Join-Path -ChildPath '2.xlsx')
        )
        Path      = ($parent | Join-Path -ChildPath '*.xlsx')
      }
    }
  }
  Describe 'Get-ExcelDocumentProperty' {
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
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        $properties = @(Get-ExcelDocumentProperty -Path $pathSet.Path)
        $properties.Count | Should-Be 2
        $properties | ForEach-Object { @($_.PSObject.Properties).Count | Should-BeGreaterThan 0 }
      }
      It 'reads built-in properties by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        $properties = @($pathSet.Path | Get-ExcelDocumentProperty)
        $properties.Count | Should-Be 2
        $properties | ForEach-Object { @($_.PSObject.Properties).Count | Should-BeGreaterThan 0 }
      }
      It 'reads built-in properties by LiteralPath' {
        New-ExcelFile -Path $path
        $properties = Get-ExcelDocumentProperty -LiteralPath $path
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
      It 'reads built-in properties by LiteralPath with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-ExcelDocumentProperty
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'filters properties by Name' {
        New-ExcelFile -Path $path
        $allProperties = Get-ExcelDocumentProperty -Path $path
        $name = @($allProperties.PSObject.Properties.Name) | Select-Object -First 1
        $properties = Get-ExcelDocumentProperty -Path $path -Name $name
        @($properties.PSObject.Properties.Name) | Should-ContainCollection $name
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
      It 'reads properties from a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Get-ExcelDocumentProperty -LiteralPath $path -PasswordToOpen (Get-Password) } | Should-Throw
        $properties = Get-ExcelDocumentProperty -LiteralPath $path -PasswordToOpen $password
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
      It 'reads properties from a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify (Get-Password)
        $properties = Get-ExcelDocumentProperty -LiteralPath $path -PasswordToModify (Get-Password)
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
      It 'does not return built-in properties when Custom is specified' {
        New-ExcelFile -Path $path
        $properties = Get-ExcelDocumentProperty -LiteralPath $path -Custom
        @($properties.PSObject.Properties).Count | Should-Be 0
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Get-ExcelDocumentProperty.Unit' {
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
      It 'throws when New-ExcelObject fails' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

        { Get-ExcelDocumentProperty -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Get-ExcelPropertyValue.Unit' {
    BeforeAll {
      Mock -CommandName Test-Path -MockWith { $true }
    }
    Context 'ParameterSetName' {
      It 'returns null when Get-ExcelDocumentProperty returns null' {
        Mock -CommandName Get-ExcelDocumentProperty

        $result = Get-ExcelPropertyValue -LiteralPath $path -Name Title
        $result | Should-BeNull
      }
      It 'returns property values from Get-ExcelDocumentProperty output' {
        Mock -CommandName Get-ExcelDocumentProperty -MockWith {
          [PSCustomObject]@{
            Title = 'A'
          },
          [PSCustomObject]@{
            Title = 'B'
          }
        }

        $result = @(Get-ExcelPropertyValue -LiteralPath 'dummy.xlsx' -Name Title)
        $result.Count | Should-Be 2
        $result | Should-BeCollection @('A', 'B')
      }
    }
  }
  Describe 'Set-ExcelDocumentProperty' {
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
            "Import-Module -Name '$escapedModulePath' -Force"
            "Set-ExcelDocumentProperty -LiteralPath '$escapedPath' -Name '$escapedName' -Value '$escapedValue' -Confirm"
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
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        Set-ExcelDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        $values = @(Get-ExcelPropertyValue -Path $pathSet.Path -Name $name)
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | Should-Be (@($value) * $pathSet.ChildItem.Count)
      }
      It 'sets a built-in property by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        $pathSet.Path | Set-ExcelDocumentProperty -Name $name -Value $value
        $values = @(Get-ExcelPropertyValue -Path $pathSet.Path -Name $name)
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | Should-Be (@($value) * $pathSet.ChildItem.Count)
      }
      It 'sets multiple built-in properties by Path with InputObject' {
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        Set-ExcelDocumentProperty -Path $pathSet.Path -InputObject $properties
        @(Get-ExcelPropertyValue -Path $pathSet.Path -Name 'Title') | Should-Be (@($properties.Title) * $pathSet.ChildItem.Count)
        @(Get-ExcelPropertyValue -Path $pathSet.Path -Name 'Subject') | Should-Be (@($properties.Subject) * $pathSet.ChildItem.Count)
      }
      It 'sets a built-in property by LiteralPath' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'sets a built-in property by LiteralPath with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-ExcelDocumentProperty -Name $name -Value $value
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'sets multiple built-in properties by LiteralPath with InputObject' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        Set-ExcelDocumentProperty -LiteralPath $path -InputObject $properties
        Get-ExcelPropertyValue -LiteralPath $path -Name 'Title' | Should-Be $properties.Title
        Get-ExcelPropertyValue -LiteralPath $path -Name 'Subject' | Should-Be $properties.Subject
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value 'ante value'
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value 'post value' -Force -WhatIf
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be 'ante value'
      }
      It 'asks for confirmation when Confirm is specified' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value 'ante confirm'
        $exitCode = 'N' | Invoke-Confirm -Path $path -Name $name -Value 'post confirm'
        $exitCode | Should-Be 0
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be 'ante confirm'
      }
      It 'updates a read-only recommended workbook when Force is specified' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -Force -Confirm
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password
        { Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        Get-ExcelPropertyValue -LiteralPath $path -Name $name -PasswordToOpen $password | Should-Be $value
      }
      It 'updates a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password
        { Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        Get-ExcelPropertyValue -LiteralPath $path -Name $name -PasswordToModify $password | Should-Be $value
      }
      It 'returns the updated built-in property when PassThru is specified' {
        New-ExcelFile -Path $path
        $property = Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.$name | Should-Be $value
      }
      It 'returns updated properties when PassThru is specified with InputObject' {
        New-ExcelFile -Path $path
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        $property = Set-ExcelDocumentProperty -LiteralPath $path -InputObject $properties -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.Title | Should-Be $properties.Title
        $property.Subject | Should-Be $properties.Subject
      }
      It 'sets a custom property' {
        New-ExcelFile -Path $path
        $name = 'MyProperty'
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -Custom
        Get-ExcelPropertyValue -LiteralPath $path -Name $name -Custom | Should-Be $value
      }
      It 'returns the updated custom property when PassThru is specified' {
        New-ExcelFile -Path $path
        $name = 'MyProperty'
        $property = Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -Custom -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.$name | Should-Be $value
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read-only recommended workbook without Force' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        { Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-ExcelDocumentProperty.Unit' {
    BeforeEach {
      $name = 'Title'
      $value = [guid]::NewGuid().ToString('N')
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null
      $script:called = 0

      Mock -CommandName New-ExcelObject -MockWith {
        $app = [PSCustomObject]@{}
        $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
        return $app
      }
      Mock -CommandName Open-ExcelFile -MockWith {
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
      It 'does not call Open-ExcelFile when WhatIf is specified' {
        { Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -WhatIf } | Should -Not -Throw
        $script:called | Should-Be 0
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        $script:called | Should-Be 0
      }
      It 'throws when New-ExcelObject fails' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

        { Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Remove-ExcelDocumentProperty' {
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
            "Remove-ExcelDocumentProperty -LiteralPath '$escapedPath' -Confirm"
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
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        Set-ExcelDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        Remove-ExcelDocumentProperty -Path $pathSet.Path
        $values = Get-ExcelPropertyValue -Path $pathSet.Path -Name $name
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | ForEach-Object { $_ | Should-BeNull }
      }
      It 'removes built-in properties by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-ExcelFile -Path $_ }
        Set-ExcelDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        $pathSet.Path | Remove-ExcelDocumentProperty
        $values = Get-ExcelPropertyValue -Path $pathSet.Path -Name $name
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | ForEach-Object { $_ | Should-BeNull }
      }
      It 'removes built-in properties by LiteralPath' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-ExcelDocumentProperty -LiteralPath $path
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
      It 'removes built-in properties by LiteralPath with ValueFromPipelineByPropertyName' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value
        [PSCustomObject]@{ PSPath = $path } | Remove-ExcelDocumentProperty
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not remove properties when WhatIf is specified' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-ExcelDocumentProperty -LiteralPath $path -Force -WhatIf
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'asks for confirmation when Confirm is specified' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should-Be 0
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'removes properties from a read-only recommended workbook when Force is specified' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -Force -Confirm
        Remove-ExcelDocumentProperty -LiteralPath $path -Force -Confirm
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
    }
    Context 'Other parameters' {
      It 'removes properties from a file protected with PasswordToOpen' {
        New-ExcelFile -Path $path -PasswordToOpen $password
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        { Remove-ExcelDocumentProperty -LiteralPath $path -PasswordToOpen (Get-Password) } | Should-Throw
        Remove-ExcelDocumentProperty -LiteralPath $path -PasswordToOpen $password
        Get-ExcelPropertyValue -LiteralPath $path -Name $name -PasswordToOpen $password | Should-BeNull
      }
      It 'removes properties from a file protected with PasswordToModify' {
        New-ExcelFile -Path $path -PasswordToModify $password
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        { Remove-ExcelDocumentProperty -LiteralPath $path -PasswordToModify (Get-Password) } | Should-Throw
        Remove-ExcelDocumentProperty -LiteralPath $path -PasswordToModify $password
        Get-ExcelPropertyValue -LiteralPath $path -Name $name -PasswordToModify $password | Should-BeNull
      }
      It 'uses RemoveDocInfoType when removing properties' {
        New-ExcelFile -Path $path
        Set-ExcelDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-ExcelDocumentProperty -LiteralPath $path -RemoveDocInfoType xlRDIAll
        Get-ExcelPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read-only recommended workbook without Force' {
        New-ExcelFile -Path $path -ReadOnlyRecommended
        { Remove-ExcelDocumentProperty -LiteralPath $path } | Should-Throw
      }
    }
  }
  Describe 'Remove-ExcelDocumentProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null
      $script:called = 0

      Mock -CommandName New-ExcelObject -MockWith {
        $app = [PSCustomObject]@{}
        $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
        return $app
      }
      Mock -CommandName Open-ExcelFile -MockWith {
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
      It 'does not call Open-ExcelFile when WhatIf is specified' {
        { Remove-ExcelDocumentProperty -LiteralPath $path -WhatIf } | Should -Not -Throw
        $script:called | Should-Be 0
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Remove-ExcelDocumentProperty -LiteralPath $path } | Should-Throw
        $script:called | Should-Be 0
      }
      It 'throws when New-ExcelObject fails' {
        Mock -CommandName New-ExcelObject -MockWith { throw 'new excel object failed' }

        { Remove-ExcelDocumentProperty -LiteralPath $path } | Should-Throw
      }
    }
  }
}

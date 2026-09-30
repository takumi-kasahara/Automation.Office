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
        $Extension = '.pptx'
      )
      return $env:TEMP | Join-Path -ChildPath "Presentation.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function Get-TempFileSet {
      [CmdletBinding()]
      [OutputType([PSCustomObject])]
      param ()
      $parent = $env:TEMP | Join-Path -ChildPath ("Presentation.$([guid]::NewGuid().ToString('N'))")
      New-Item -Path $parent -ItemType Directory | Out-Null
      return [PSCustomObject]@{
        Parent    = $parent
        ChildItem = @(
          ($parent | Join-Path -ChildPath '1.pptx')
          ($parent | Join-Path -ChildPath '2.pptx')
        )
        Path      = ($parent | Join-Path -ChildPath '*.pptx')
      }
    }
  }
  Describe 'Get-PowerPointDocumentProperty' {
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
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        $properties = @(Get-PowerPointDocumentProperty -Path $pathSet.Path)
        $properties.Count | Should-Be 2
        $properties | ForEach-Object { @($_.PSObject.Properties).Count | Should-BeGreaterThan 0 }
      }
      It 'reads built-in properties by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        $properties = @($pathSet.Path | Get-PowerPointDocumentProperty)
        $properties.Count | Should-Be 2
        $properties | ForEach-Object { @($_.PSObject.Properties).Count | Should-BeGreaterThan 0 }
      }
      It 'reads built-in properties by LiteralPath' {
        New-PowerPointFile -Path $path
        $properties = Get-PowerPointDocumentProperty -LiteralPath $path
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
      It 'reads built-in properties by LiteralPath with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{ PSPath = $path } | Get-PowerPointDocumentProperty
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
    }
    Context 'Other parameters' {
      It 'filters properties by Name' {
        New-PowerPointFile -Path $path
        $allProperties = Get-PowerPointDocumentProperty -Path $path
        $name = @($allProperties.PSObject.Properties.Name) | Select-Object -First 1
        $properties = Get-PowerPointDocumentProperty -Path $path -Name $name
        @($properties.PSObject.Properties.Name) | Should-ContainCollection $name
        @($properties.PSObject.Properties).Count | Should-Be 1
      }
      It 'reads properties from a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password
        { Get-PowerPointDocumentProperty -LiteralPath $path -PasswordToOpen (Get-Password) } | Should-Throw
        $properties = Get-PowerPointDocumentProperty -LiteralPath $path -PasswordToOpen $password
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
      It 'reads properties from a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify (Get-Password)
        $properties = Get-PowerPointDocumentProperty -LiteralPath $path -PasswordToModify (Get-Password)
        @($properties.PSObject.Properties).Count | Should-BeGreaterThan 0
      }
      It 'does not return built-in properties when Custom is specified' {
        New-PowerPointFile -Path $path
        $properties = Get-PowerPointDocumentProperty -LiteralPath $path -Custom
        @($properties.PSObject.Properties).Count | Should-Be 0
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Get-PowerPointDocumentProperty.Unit' {
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

        { Get-PowerPointDocumentProperty -Path $path } | Should-Throw
      }
    }
  }
  Describe 'Get-PowerPointPropertyValue.Unit' {
    BeforeAll {
      Mock -CommandName Test-Path -MockWith { $true }
    }
    Context 'ParameterSetName' {
      It 'returns null when Get-PowerPointDocumentProperty returns null' {
        Mock -CommandName Get-PowerPointDocumentProperty

        $result = Get-PowerPointPropertyValue -LiteralPath 'dummy.pptx' -Name Title
        $result | Should-BeNull
      }
      It 'returns property values from Get-PowerPointDocumentProperty output' {
        Mock -CommandName Get-PowerPointDocumentProperty -MockWith {
          [PSCustomObject]@{
            Title = 'A'
          },
          [PSCustomObject]@{
            Title = 'B'
          }
        }

        $result = @(Get-PowerPointPropertyValue -LiteralPath 'dummy.pptx' -Name Title)
        $result.Count | Should-Be 2
        $result | Should-BeCollection @('A', 'B')
      }
    }
  }
  Describe 'Set-PowerPointDocumentProperty' {
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
            "Set-PowerPointDocumentProperty -LiteralPath '$escapedPath' -Name '$escapedName' -Value '$escapedValue' -Confirm"
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
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        Set-PowerPointDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        $values = @(Get-PowerPointPropertyValue -Path $pathSet.Path -Name $name)
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | Should-Be (@($value) * $pathSet.ChildItem.Count)
      }
      It 'sets a built-in property by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        $pathSet.Path | Set-PowerPointDocumentProperty -Name $name -Value $value
        $values = @(Get-PowerPointPropertyValue -Path $pathSet.Path -Name $name)
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | Should-Be (@($value) * $pathSet.ChildItem.Count)
      }
      It 'sets multiple built-in properties by Path with InputObject' {
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        Set-PowerPointDocumentProperty -Path $pathSet.Path -InputObject $properties
        @(Get-PowerPointPropertyValue -Path $pathSet.Path -Name 'Title') | Should-Be (@($properties.Title) * $pathSet.ChildItem.Count)
        @(Get-PowerPointPropertyValue -Path $pathSet.Path -Name 'Subject') | Should-Be (@($properties.Subject) * $pathSet.ChildItem.Count)
      }
      It 'sets a built-in property by LiteralPath' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'sets a built-in property by LiteralPath with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path
        [PSCustomObject]@{ PSPath = $path } | Set-PowerPointDocumentProperty -Name $name -Value $value
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'sets multiple built-in properties by LiteralPath with InputObject' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        Set-PowerPointDocumentProperty -LiteralPath $path -InputObject $properties
        Get-PowerPointPropertyValue -LiteralPath $path -Name 'Title' | Should-Be $properties.Title
        Get-PowerPointPropertyValue -LiteralPath $path -Name 'Subject' | Should-Be $properties.Subject
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update properties when WhatIf is specified' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value 'ante value'
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value 'post value' -Force -WhatIf
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-Be 'ante value'
      }
      It 'asks for confirmation when Confirm is specified' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value 'ante confirm'
        $exitCode = 'N' | Invoke-Confirm -Path $path -Name $name -Value 'post confirm'
        $exitCode | Should-Be 0
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-Be 'ante confirm'
      }
    }
    Context 'Other parameters' {
      It 'updates a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password
        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen (Get-Password) } | Should-Throw
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name -PasswordToOpen $password | Should-Be $value
      }
      It 'updates a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password
        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify (Get-Password) } | Should-Throw
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name -PasswordToModify $password | Should-Be $value
      }
      It 'returns the updated built-in property when PassThru is specified' {
        New-PowerPointFile -Path $path
        $property = Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.$name | Should-Be $value
      }
      It 'returns updated properties when PassThru is specified with InputObject' {
        New-PowerPointFile -Path $path
        $properties = [PSCustomObject]@{
          Title   = [guid]::NewGuid().ToString('N')
          Subject = [guid]::NewGuid().ToString('N')
        }
        $property = Set-PowerPointDocumentProperty -LiteralPath $path -InputObject $properties -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.Title | Should-Be $properties.Title
        $property.Subject | Should-Be $properties.Subject
      }
      It 'sets a custom property' {
        New-PowerPointFile -Path $path
        $name = 'MyProperty'
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -Custom
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name -Custom | Should-Be $value
      }
      It 'returns the updated custom property when PassThru is specified' {
        New-PowerPointFile -Path $path
        $name = 'MyProperty'
        $property = Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -Custom -PassThru
        $property | Should-HaveType ([PSCustomObject])
        $property.$name | Should-Be $value
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read-only recommended presentation' {
        New-PowerPointFile -Path $path -ReadOnlyRecommended
        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Set-PowerPointDocumentProperty.Unit' {
    BeforeEach {
      $name = 'Title'
      $value = [guid]::NewGuid().ToString('N')
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null

      Mock -CommandName New-PowerPointObject -MockWith {
        $app = [PSCustomObject]@{}
        $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
        return $app
      }
      Mock -CommandName Open-PowerPointFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call Open-PowerPointFile when WhatIf is specified' {
        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -WhatIf } | Should -Not -Throw

        Should-NotInvoke -CommandName Open-PowerPointFile
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
        Should-NotInvoke -CommandName Open-PowerPointFile
      }
      It 'throws when New-PowerPointObject fails' {
        Mock -CommandName New-PowerPointObject -MockWith { throw }

        { Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value } | Should-Throw
      }
    }
  }
  Describe 'Remove-PowerPointDocumentProperty' {
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
            "Remove-PowerPointDocumentProperty -LiteralPath '$escapedPath' -Confirm"
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
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        Set-PowerPointDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        Remove-PowerPointDocumentProperty -Path $pathSet.Path
        $values = Get-PowerPointPropertyValue -Path $pathSet.Path -Name $name
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | ForEach-Object { $_ | Should-BeNull }
      }
      It 'removes built-in properties by Path with ValueFromPipeline' {
        $pathSet.ChildItem | ForEach-Object { New-PowerPointFile -Path $_ }
        Set-PowerPointDocumentProperty -Path $pathSet.Path -Name $name -Value $value
        $pathSet.Path | Remove-PowerPointDocumentProperty
        $values = Get-PowerPointPropertyValue -Path $pathSet.Path -Name $name
        $values.Count | Should-Be $pathSet.ChildItem.Count
        $values | ForEach-Object { $_ | Should-BeNull }
      }
      It 'removes built-in properties by LiteralPath' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-PowerPointDocumentProperty -LiteralPath $path
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
      It 'removes built-in properties by LiteralPath with ValueFromPipelineByPropertyName' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value
        [PSCustomObject]@{ PSPath = $path } | Remove-PowerPointDocumentProperty
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not remove properties when WhatIf is specified' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-PowerPointDocumentProperty -LiteralPath $path -Force -WhatIf
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
      It 'asks for confirmation when Confirm is specified' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value
        $exitCode = 'N' | Invoke-Confirm -Path $path
        $exitCode | Should-Be 0
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-Be $value
      }
    }
    Context 'Other parameters' {
      It 'removes properties from a file protected with PasswordToOpen' {
        New-PowerPointFile -Path $path -PasswordToOpen $password
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToOpen $password
        { Remove-PowerPointDocumentProperty -LiteralPath $path } | Should-Throw
        { Remove-PowerPointDocumentProperty -LiteralPath $path -PasswordToOpen (Get-Password) } | Should-Throw
        Remove-PowerPointDocumentProperty -LiteralPath $path -PasswordToOpen $password
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name -PasswordToOpen $password | Should-BeNull
      }
      It 'removes properties from a file protected with PasswordToModify' {
        New-PowerPointFile -Path $path -PasswordToModify $password
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value -PasswordToModify $password
        { Remove-PowerPointDocumentProperty -LiteralPath $path } | Should-Throw
        { Remove-PowerPointDocumentProperty -LiteralPath $path -PasswordToModify (Get-Password) } | Should-Throw
        Remove-PowerPointDocumentProperty -LiteralPath $path -PasswordToModify $password
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name -PasswordToModify $password | Should-BeNull
      }
      It 'uses RemoveDocInfoType when removing properties' {
        New-PowerPointFile -Path $path
        Set-PowerPointDocumentProperty -LiteralPath $path -Name $name -Value $value
        Remove-PowerPointDocumentProperty -LiteralPath $path -RemoveDocInfoType ppRDIAll
        Get-PowerPointPropertyValue -LiteralPath $path -Name $name | Should-BeNull
      }
    }
    Context 'Edge cases' {
      It 'throws an error when read-only recommended presentation' {
        New-PowerPointFile -Path $path -ReadOnlyRecommended
        { Remove-PowerPointDocumentProperty -LiteralPath $path } | Should-Throw
      }
    }
  }
  Describe 'Remove-PowerPointDocumentProperty.Unit' {
    BeforeEach {
      $path = Get-TempFile
      New-Item -Path $path -ItemType File -Force | Out-Null

      Mock -CommandName New-PowerPointObject -MockWith {
        $app = [PSCustomObject]@{}
        $app | Add-Member -MemberType ScriptMethod -Name Quit -Value { }
        return $app
      }
      Mock -CommandName Open-PowerPointFile
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call Open-PowerPointFile when WhatIf is specified' {
        { Remove-PowerPointDocumentProperty -LiteralPath $path -WhatIf } | Should -Not -Throw
        Should-NotInvoke -CommandName Open-PowerPointFile
      }
    }
    Context 'Edge cases' {
      It 'throws when target item is read-only' {
        (Get-Item -LiteralPath $path -Force).IsReadOnly = $true

        { Remove-PowerPointDocumentProperty -LiteralPath $path } | Should-Throw
        Should-NotInvoke -CommandName Open-PowerPointFile
      }
      It 'throws when New-PowerPointObject fails' {
        Mock -CommandName New-PowerPointObject -MockWith { throw }

        { Remove-PowerPointDocumentProperty -LiteralPath $path } | Should-Throw
      }
    }
  }
}

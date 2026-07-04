using namespace Microsoft.Office.Core
using namespace Microsoft.Office.Interop.PowerPoint
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Security

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()

$modulePath = $PSScriptRoot | Join-Path -ChildPath '..\Automation.Office.psm1'
Import-Module -Name $modulePath -Force
Set-StrictMode -Version Latest

InModuleScope 'PowerPoint.Automation' {
  BeforeAll {
    Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
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
        $Extension = '.pptx'
      )
      return $env:TEMP | Join-Path -ChildPath "Presentation.$([guid]::NewGuid().ToString('N'))$Extension"
    }
    function New-PowerPointFileWithSpeakerNote {
      [CmdletBinding()]
      [OutputType([void])]
      [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for PowerPoint application and does not change persistent state')]
      param (
        [Parameter(Mandatory)]
        [string]
        $Path,
        [SecureString]
        $PasswordToOpen,
        [SecureString]
        $PasswordToModify,
        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $SlideCount
      )
      New-PowerPointFile -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Initialize {
        param (
          [Parameter(Mandatory)]
          [Presentation]
          $Presentation
        )
        foreach ($index in 1..$SlideCount) {
          $slide = $Presentation.Slides.Add($index, [PpSlideLayout]::ppLayoutText)
          $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "page $index note"
        }
      } | Out-Null
    }
    function New-PowerPointFileWithHiddenSpeakerNote {
      [CmdletBinding()]
      [OutputType([void])]
      [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for PowerPoint application and does not change persistent state')]
      param (
        [Parameter(Mandatory)]
        [string]
        $Path,
        [SecureString]
        $PasswordToOpen,
        [SecureString]
        $PasswordToModify
      )
      New-PowerPointFile @PSBoundParameters -Initialize {
        param (
          [Parameter(Mandatory)]
          [Presentation]
          $Presentation
        )
        $slide1 = $Presentation.Slides.Add(1, [PpSlideLayout]::ppLayoutText)
        $slide2 = $Presentation.Slides.Add(2, [PpSlideLayout]::ppLayoutText)

        $slide1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'visible note'
        $slide2.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'hidden note'
        $slide2.SlideShowTransition.Hidden = [MsoTriState]::msoTrue
      } | Out-Null
    }
  }
  Describe 'Get-PowerPointSpeakerNote' {
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
      It 'returns speaker notes by Path' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2

        $result = Get-PowerPointSpeakerNote -Path $path

        $result | Should -Not -BeNullOrEmpty
        $result.Path | Should -Be ([Path]::GetFullPath($path))
        $result.Items -is [SpeakerNote[]] | Should -BeTrue
        $result.Items[0].Page | Should -Be 1
        $result.Items[0].Text | Should -Be 'page 1 note'
      }
      It 'returns speaker notes by LiteralPath' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2

        $result = Get-PowerPointSpeakerNote -LiteralPath $path

        $result | Should -Not -BeNullOrEmpty
        $result.Items[1].Page | Should -Be 2
        $result.Items[1].Text | Should -Be 'page 2 note'
      }
      It 'returns speaker notes by ValueFromPipelineByPropertyName' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2

        $result = [PSCustomObject]@{ PSPath = $path } | Get-PowerPointSpeakerNote

        $result | Should -Not -BeNullOrEmpty
        $result.Items[0].Page | Should -Be 1
        $result.Items[0].Text | Should -Be 'page 1 note'
      }
    }
    Context 'Other parameters' {
      It 'does not return notes from hidden slides by default' {
        New-PowerPointFileWithHiddenSpeakerNote -Path $path

        $result = Get-PowerPointSpeakerNote -Path $path

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Count | Should -Be 1
        $result.Items[0].Page | Should -Be 1
        $result.Items[0].Text | Should -Be 'visible note'
      }
      It 'returns notes from hidden slides when Force is specified' {
        New-PowerPointFileWithHiddenSpeakerNote -Path $path

        $result = Get-PowerPointSpeakerNote -Path $path -Force

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Count | Should -Be 2
        $result.Items[1].Page | Should -Be 2
        $result.Items[1].Text | Should -Be 'hidden note'
      }
      It 'throws when PasswordToOpen is incorrect' {
        New-PowerPointFile -Path $path -PasswordToOpen $password | Out-Null

        { Get-PowerPointSpeakerNote -Path $path -PasswordToOpen (Get-Password) } | Should -Throw
      }
      It 'returns notes only for requested discrete and range mix' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 8

        $result = Get-PowerPointSpeakerNote -Path $path -Range '1-3, 5, 7-8'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(1, 2, 3, 5, 7, 8)
        $result.Items.Text | Should -Be @('page 1 note', 'page 2 note', 'page 3 note', 'page 5 note', 'page 7 note', 'page 8 note')
      }
      It 'supports spaces in range expression' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 6

        $result = Get-PowerPointSpeakerNote -Path $path -Range '2, 4-5'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(2, 4, 5)
      }
      It 'works with non sorted range token order' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 6

        $result = Get-PowerPointSpeakerNote -Path $path -Range '4-6, 2'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(4, 5, 6, 2)
      }
      It 'supports reverse order range' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 5

        $result = Get-PowerPointSpeakerNote -Path $path -Range '3-1'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(3, 2, 1)
      }
    }
  }
  Describe 'Get-PowerPointSpeakerNote.Unit' {
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
        Mock -CommandName New-PowerPointObject -MockWith { throw 'new powerpoint object failed' }

        { Get-PowerPointSpeakerNote -Path $path } | Should -Throw
      }
      It 'throws for invalid characters in Range' {
        { Get-PowerPointSpeakerNote -Path $path -Range 'a' } | Should -Throw
        { Get-PowerPointSpeakerNote -Path $path -Range '1-' } | Should -Throw
        { Get-PowerPointSpeakerNote -Path $path -Range '1:2' } | Should -Throw
      }
    }
  }
  Describe 'Export-PowerPointAsFixedFormat' {
    BeforeEach {
      $path = Get-TempFile
      $passwordToOpen = Get-Password
      $passwordToModify = Get-Password
      $pdf = $env:TEMP | Join-Path -ChildPath "Exported.$([guid]::NewGuid().ToString('N')).pdf"
      $xps = $env:TEMP | Join-Path -ChildPath "Exported.$([guid]::NewGuid().ToString('N')).xps"
    }
    AfterEach {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
      }
      if (Test-Path -LiteralPath $pdf) {
        Remove-Item -LiteralPath $pdf -Force
      }
      if (Test-Path -LiteralPath $xps) {
        Remove-Item -LiteralPath $xps -Force
      }
    }
    Context 'ParameterSetName' {
      It 'exports a presentation to PDF by default' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        $result = Export-PowerPointAsFixedFormat -Path $path -Destination $pdf

        $result | Should -Not -BeNullOrEmpty
        $result.FullName | Should -Be ([Path]::GetFullPath($pdf))
        Test-Path -LiteralPath $pdf | Should -BeTrue
      }
      It 'exports by ValueFromPipelineByPropertyName using FullName alias' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        $result = [PSCustomObject]@{ FullName = $path } |
        Export-PowerPointAsFixedFormat -Destination $pdf

        $result | Should -Not -BeNullOrEmpty
        $result.FullName | Should -Be ([Path]::GetFullPath($pdf))
        Test-Path -LiteralPath $pdf | Should -BeTrue
      }
    }
    Context 'SupportShouldProcess' {
      It 'does not create file when WhatIf is specified' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -WhatIf

        Test-Path -LiteralPath $pdf | Should -BeFalse
      }
      It 'overwrites existing file when Force is specified' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        { Export-PowerPointAsFixedFormat -Path $path -Destination $pdf } | Should -Not -Throw
        { Export-PowerPointAsFixedFormat -Path $path -Destination $pdf } | Should -Throw

        $result = Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -Force
        $result | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath $pdf | Should -BeTrue
      }
      It 'throws when file exists and NoClobber is specified' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        Export-PowerPointAsFixedFormat -Path $path -Destination $pdf | Out-Null

        { Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -NoClobber } | Should -Throw
      }
    }
    Context 'Other parameters' {
      It 'exports a presentation to PDF with PasswordToOpen' {
        New-PowerPointFileWithSpeakerNote -Path $path -PasswordToOpen $passwordToOpen -SlideCount 2
        $result = Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -PasswordToOpen $passwordToOpen

        $result | Should -Not -BeNullOrEmpty
        $result.FullName | Should -Be ([Path]::GetFullPath($pdf))
        Test-Path -LiteralPath $pdf | Should -BeTrue
      }
      It 'exports a presentation to PDF with PasswordToModify' {
        New-PowerPointFileWithSpeakerNote -Path $path -PasswordToModify $passwordToModify -SlideCount 2
        $result = Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -PasswordToModify $passwordToModify

        $result | Should -Not -BeNullOrEmpty
        $result.FullName | Should -Be ([Path]::GetFullPath($pdf))
        Test-Path -LiteralPath $pdf | Should -BeTrue
      }
      It 'throws when PasswordToOpen is incorrect' {
        New-PowerPointFileWithSpeakerNote -Path $path -PasswordToOpen $passwordToOpen -SlideCount 2

        { Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -PasswordToOpen (Get-Password) } | Should -Throw
      }
      It 'exports a presentation with FixedFormatType' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        $result = Export-PowerPointAsFixedFormat -Path $path -Destination $xps -FixedFormatType ([PpFixedFormatType]::ppFixedFormatTypeXPS)

        $result | Should -Not -BeNullOrEmpty
        $result.FullName | Should -Be ([Path]::GetFullPath($xps))
        Test-Path -LiteralPath $xps | Should -BeTrue
      }
      It 'exports a presentation with IncludeDocumentProperties' {
        New-PowerPointFileWithSpeakerNote -Path $path -SlideCount 2
        $result = Export-PowerPointAsFixedFormat -Path $path -Destination $pdf -IncludeDocumentProperties

        $result | Should -Not -BeNullOrEmpty
        $result.FullName | Should -Be ([Path]::GetFullPath($pdf))
        Test-Path -LiteralPath $pdf | Should -BeTrue
      }
    }
    Context 'Edge cases' {
      It 'throws when source file does not exist' {
        $invalidPath = $env:TEMP | Join-Path -ChildPath 'NonExistent.pptx'
        { Export-PowerPointAsFixedFormat -Path $invalidPath -Destination $pdf } | Should -Throw
      }
      It 'throws when destination is an invalid path' {
        $invalidDest = 'Z:\Invalid\Path\File.pdf'
        { Export-PowerPointAsFixedFormat -Path $path -Destination $invalidDest } | Should -Throw
      }
    }
  }
}

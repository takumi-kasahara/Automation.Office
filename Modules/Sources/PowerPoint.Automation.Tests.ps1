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
  }
  Describe 'Get-PowerPointSpeakerNote' {
    BeforeAll {
      function New-PowerPointFileWithSpeakerNote {
        [CmdletBinding()]
        [OutputType([void])]
        [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates a new COM object for PowerPoint application and does not change persistent state')]
        param (
          [Parameter(Mandatory)]
          [string]
          $Path
        )
        $app = New-PowerPointObject
        try {
          $presentation = $app.Presentations.Add([MsoTriState]::msoFalse)
          try {
            $slide1 = $presentation.Slides.Add(1, [PpSlideLayout]::ppLayoutText)
            $slide2 = $presentation.Slides.Add(2, [PpSlideLayout]::ppLayoutText)

            $slide1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'page 1 note'
            $slide2.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'page 2 note'

            $presentation.SaveCopyAs2(
              [Path]::GetFullPath($Path)
              , [PpSaveAsFileType]::ppSaveAsOpenXMLPresentation
              , [type]::Missing
              , $false
            )
          }
          finally {
            $presentation.Close()
          }
        }
        finally {
          try {
            if ($app) {
              $app.Quit()
            }
          }
          finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      function New-PowerPointFileWithHiddenSpeakerNote {
        [CmdletBinding()]
        [OutputType([void])]
        param (
          [Parameter(Mandatory)]
          [string]
          $Path
        )
        $app = New-PowerPointObject
        try {
          $presentation = $app.Presentations.Add([MsoTriState]::msoFalse)
          try {
            $slide1 = $presentation.Slides.Add(1, [PpSlideLayout]::ppLayoutText)
            $slide2 = $presentation.Slides.Add(2, [PpSlideLayout]::ppLayoutText)

            $slide1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'visible note'
            $slide2.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'hidden note'
            $slide2.SlideShowTransition.Hidden = [MsoTriState]::msoTrue

            $presentation.SaveCopyAs2(
              [Path]::GetFullPath($Path)
              , [PpSaveAsFileType]::ppSaveAsOpenXMLPresentation
              , [type]::Missing
              , $false
            )
          }
          finally {
            $presentation.Close()
          }
        }
        finally {
          try {
            if ($app) {
              $app.Quit()
            }
          }
          finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
      function New-PowerPointFileWithSequentialSpeakerNotes {
        [CmdletBinding()]
        [OutputType([void])]
        param (
          [Parameter(Mandatory)]
          [string]
          $Path,
          [Parameter(Mandatory)]
          [ValidateRange(1, [int]::MaxValue)]
          [int]
          $SlideCount
        )
        $app = New-PowerPointObject
        try {
          $presentation = $app.Presentations.Add([MsoTriState]::msoFalse)
          try {
            foreach ($index in 1..$SlideCount) {
              $slide = $presentation.Slides.Add($index, [PpSlideLayout]::ppLayoutText)
              $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "page $index note"
            }

            $presentation.SaveCopyAs2(
              [Path]::GetFullPath($Path)
              , [PpSaveAsFileType]::ppSaveAsOpenXMLPresentation
              , [type]::Missing
              , $false
            )
          }
          finally {
            $presentation.Close()
          }
        }
        finally {
          try {
            if ($app) {
              $app.Quit()
            }
          }
          finally {
            Get-Variable |
            Where-Object -Property Value -Is [__ComObject] |
            Clear-Variable -Force -WhatIf:$false -Confirm:$false
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
          }
        }
      }
    }
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
        New-PowerPointFileWithSpeakerNote -Path $path

        $result = Get-PowerPointSpeakerNote -Path $path

        $result | Should -Not -BeNullOrEmpty
        $result.Path | Should -Be ([Path]::GetFullPath($path))
        $result.Items -is [SpeakerNote[]] | Should -BeTrue
        $result.Items[0].Page | Should -Be 1
        $result.Items[0].Text | Should -Be 'page 1 note'
      }
      It 'returns speaker notes by LiteralPath' {
        New-PowerPointFileWithSpeakerNote -Path $path

        $result = Get-PowerPointSpeakerNote -LiteralPath $path

        $result | Should -Not -BeNullOrEmpty
        $result.Items[1].Page | Should -Be 2
        $result.Items[1].Text | Should -Be 'page 2 note'
      }
      It 'returns speaker notes by ValueFromPipelineByPropertyName' {
        New-PowerPointFileWithSpeakerNote -Path $path

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
        New-PowerPointFileWithSequentialSpeakerNotes -Path $path -SlideCount 8

        $result = Get-PowerPointSpeakerNote -Path $path -Range '1-3, 5, 7-8'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(1, 2, 3, 5, 7, 8)
        $result.Items.Text | Should -Be @('page 1 note', 'page 2 note', 'page 3 note', 'page 5 note', 'page 7 note', 'page 8 note')
      }
      It 'supports spaces in range expression' {
        New-PowerPointFileWithSequentialSpeakerNotes -Path $path -SlideCount 6

        $result = Get-PowerPointSpeakerNote -Path $path -Range '2, 4-5'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(2, 4, 5)
      }
      It 'works with non sorted range token order' {
        New-PowerPointFileWithSequentialSpeakerNotes -Path $path -SlideCount 6

        $result = Get-PowerPointSpeakerNote -Path $path -Range '4-6, 2'

        $result | Should -Not -BeNullOrEmpty
        $result.Items.Page | Should -Be @(4, 5, 6, 2)
      }
      It 'supports reverse order range' {
        New-PowerPointFileWithSequentialSpeakerNotes -Path $path -SlideCount 5

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
}

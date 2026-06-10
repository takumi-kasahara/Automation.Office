using namespace Microsoft.Office.Core
using namespace Microsoft.Office.Interop.PowerPoint
using namespace System.Management.Automation
using namespace System.Runtime.InteropServices

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Set-StrictMode -Version Latest

#region Private
class SpeakerNote {
  [int]$Page
  [string]$Text
}
class SpeakerNotes {
  [string]$Path
  [SpeakerNote[]]$Items
}
#endregion
#region Public
function Get-PowerPointSpeakerNote {
  <#
  .SYNOPSIS
    Gets speaker notes from one or more PowerPoint presentation files.

  .DESCRIPTION
    Opens one or more PowerPoint presentations in read-only mode and extracts speaker notes from each slide.

    The cmdlet supports both `-Path` and `-LiteralPath` parameter sets and can open protected presentations by using `-PasswordToOpen` and `-PasswordToModify`.

    Each output object contains the source file path and a collection of speaker notes, where each note includes the slide number and the corresponding text.

  .PARAMETER Path
    Specifies presentation paths. Wildcards are supported.

  .PARAMETER LiteralPath
    Specifies presentation paths literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

    Pass a `SecureString` value. If omitted, the presentation is opened without an open password.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

    Pass a `SecureString` value. If omitted, the presentation is opened without a modify password.

  .PARAMETER Force
    Includes speaker notes from hidden slides.

    If omitted, speaker notes from hidden slides are not included in the output.

  .EXAMPLE
    Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx"

    Gets speaker notes from the specified presentation.

  .EXAMPLE
    Get-PowerPointSpeakerNote -LiteralPath "$env:TEMP\Presentation.pptx"

    Gets speaker notes from the specified presentation by using a literal path.

  .EXAMPLE
    $note = Get-PowerPointSpeakerNote -LiteralPath "$env:TEMP\Presentation.pptx"
    $note.Items |
    ForEach-Object { $_.Text -split "`r" } |
    Where-Object { $_.Trim() -ne [string]::Empty } |
    Out-File -LiteralPath "$([IO.Path]::GetFileName($note.Path)).txt" -Append -Encoding utf8

    Exports each speaker note to a separate text file named after the source presentation and slide number.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx" -PasswordToOpen $password

    Gets speaker notes from a presentation protected with an open password.

  .EXAMPLE
    Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx" -Force

    Gets speaker notes including notes on hidden slides.

  .OUTPUTS
    SpeakerNotes
      Returns an object that contains the source path and a collection of `SpeakerNote` objects.

  .NOTES
    The presentation is always opened in read-only mode and the PowerPoint COM object is released
    after processing.
  #>
  [CmdletBinding(DefaultParameterSetName = 'PathSet')]
  [OutputType([SpeakerNotes])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [switch]
    $Force
  )
  begin {
    $app = New-PowerPointObject
  }
  process {
    try {
      $items = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
        'PathSet' {
          Get-Item -Path $Path -Force
        }
        'LiteralPathSet' {
          Get-Item -LiteralPath $LiteralPath -Force
        }
      }
      $items |
      ForEach-Object {
        $item = $_
        $file = Open-PowerPointFile -Application $app -Path $item.FullName -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly
        $notes = @()
        try {
          foreach ($slide in @($file.Slides)) {
            if (-not $Force -and ($slide.SlideShowTransition.Hidden -ne [MsoTriState]::msoFalse)) {
              continue
            }
            $text = [string]::Empty
            try {
              $text = [string]$slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text.TrimEnd("`r", "`n")
            }
            catch {
              $text = [string]::Empty
            }
            $note = [SpeakerNote]::new()
            $note.Page = [int]$slide.SlideIndex
            $note.Text = $text
            if ($note.Text) {
              $notes += $note
            }
          }
          $speakerNotes = [SpeakerNotes]::new()
          $speakerNotes.Path = $item.FullName
          $speakerNotes.Items = $notes
          return $speakerNotes
        }
        finally {
          $file.Close()
        }
      }
    }
    catch {
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
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
  end {
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
#endregion

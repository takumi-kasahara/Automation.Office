using module .\VBProject.psm1
using namespace System.IO
using namespace System.Text

Set-StrictMode -Version Latest

function Export-PowerPointVBProject {
  <#
  .SYNOPSIS
    Exports VBA project metadata and components from a PowerPoint presentation.

  .DESCRIPTION
    Opens a file and exports references and components into a unified `VBProject.json` format.

    The companion component folder is created next to the destination JSON unless -ComponentRoot is specified.
    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.

  .PARAMETER Path
    Specifies the presentation path literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER Destination
    Specifies the export file path for `VBProject.json`.

  .PARAMETER ComponentRoot
    Specifies the directory path where component files are exported.
    When omitted, the directory is derived from -Destination by removing its extension.

  .PARAMETER Force
    Overrides read-only destination file attributes.

  .PARAMETER NoClobber
    Produces an error if -Destination already exists. If omitted, existing output is overwritten.

  .EXAMPLE
    Export-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Destination "$env:TEMP\VBProject.json"

    Exports VBA project metadata and components from a PowerPoint presentation to VBProject.json and creates a companion component folder.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Export-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -PasswordToOpen $password -Destination "$env:TEMP\VBProject.json"

    Exports VBA project from a presentation protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Export-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -PasswordToModify $password -Destination "$env:TEMP\VBProject.json"

    Exports VBA project from a presentation protected with a modify password.

  .EXAMPLE
    Export-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Destination "$env:TEMP\VBProject.json" -WhatIf

    Shows what would happen if the command were executed without actually exporting the VBA project.

  .OUTPUTS
    System.IO.FileInfo
      Returns the created export file.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Alias('FilePath', 'FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    [string]
    $ComponentRoot,
    [switch]
    $Force,
    [switch]
    $NoClobber
  )
  process {
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($Path, "Export VBProject to $Destination"))) {
      return
    }
    $app = New-PowerPointObject
    $arguments = @{
      Destination   = $Destination
      ComponentRoot = $ComponentRoot
      Force         = $Force
      NoClobber     = $NoClobber
    }
    try {
      Open-PowerPointFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly -Action {
        param (
          [Microsoft.Office.Interop.PowerPoint.Presentation]
          $Presentation
        )
        return Export-VBProject -VBProject $Presentation.VBProject @arguments
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
function Import-PowerPointVBProject {
  <#
  .SYNOPSIS
    Imports VBA project metadata and components into a PowerPoint presentation.

  .DESCRIPTION
    Opens a file and imports references and components from unified `VBProject.json` format.

    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.
    You can import into protected presentations with -PasswordToOpen or -PasswordToModify.

  .PARAMETER Path
    Specifies the presentation path literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the presentation.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the presentation.

  .PARAMETER Source
    Specifies the source file path for `VBProject.json`.

  .PARAMETER Force
    Forces import processing when a file is marked as read-only recommended.

  .PARAMETER Hidden
    Imports while keeping the presentation window hidden.

  .EXAMPLE
    Import-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Source "$env:TEMP\VBProject.json"

    Imports VBA project from JSON file to PowerPoint presentation.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Import-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Source "$env:TEMP\VBProject.json" -PasswordToOpen $password

    Imports into presentation protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Import-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Source "$env:TEMP\VBProject.json" -PasswordToModify $password

    Imports into presentation protected with a modify password.

  .EXAMPLE
    Import-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Source "$env:TEMP\VBProject.json" -WhatIf

    Shows what would happen if the command were executed without actually importing the VBA project.

  .EXAMPLE
    Import-PowerPointVBProject -Path "$env:TEMP\Presentation.pptm" -Source "$env:TEMP\VBProject.json" -Force -Confirm

    Forces overwrite and confirms before execution.

  .OUTPUTS
    None.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Alias('FilePath', 'FullName')]
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Path,
    [SecureString]
    $PasswordToOpen = $null,
    [SecureString]
    $PasswordToModify = $null,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Source,
    [switch]
    $Hidden
  )
  process {
    if (-not $PSCmdlet.ShouldProcess($Path, "Import VBProject from $Source")) {
      return
    }
    if ((Get-Item -LiteralPath $Path -Force).IsReadOnly) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $Path))
    }
    $app = New-PowerPointObject
    try {
      Open-PowerPointFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Action {
        param (
          [Microsoft.Office.Interop.PowerPoint.Presentation]
          $Presentation
        )
        if ($Presentation.ReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $Path))
        }
        if ($Presentation.Windows.Count -gt 0) {
          $Presentation.Windows(1).Visible = -not $Hidden
        }
        Import-VBProject -VBProject $Presentation.VBProject -Source $Source -TargetExe 'POWERPNT.EXE'
        $Presentation.Save()
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

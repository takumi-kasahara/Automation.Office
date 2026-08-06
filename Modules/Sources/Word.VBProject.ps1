using namespace System.IO
using namespace System.Text

Import-Module -Name ($PSScriptRoot | Join-Path -ChildPath 'VBProject.psm1')
Set-StrictMode -Version Latest

function Export-WordVBProject {
  <#
  .SYNOPSIS
    Exports VBA project metadata and components from a Word document.

  .DESCRIPTION
    Opens a file and exports references and components into a unified `VBProject.json` format.

    The companion component folder is created next to the destination JSON unless -ComponentRoot is specified.
    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.


  .PARAMETER Path
    Specifies the document path literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the document.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the document.

  .PARAMETER Destination
    Specifies the export file path for `VBProject.json`.

  .PARAMETER ComponentRoot
    Specifies the directory path where component files are exported.
    When omitted, the directory is derived from -Destination by removing its extension.

  .PARAMETER Force
    Overrides read-only destination file attributes.

  .PARAMETER NoClobber
    Produces an error if -Destination already exists. If omitted, existing output is overwritten.

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
    $app = New-WordObject
    try {
      Open-WordFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly -Action {
        param (
          [Parameter(Mandatory)]
          [Microsoft.Office.Interop.Word.Document]
          $Document
        )
        return Export-VBProject -VBProject $Document.VBProject -Destination $Destination -ComponentRoot $ComponentRoot -Force:$Force -NoClobber:$NoClobber
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
function Import-WordVBProject {
  <#
  .SYNOPSIS
    Imports VBA project metadata and components into a Word document.

  .DESCRIPTION
    Opens a file and imports references and components from unified `VBProject.json` format.

    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.
    You can import into protected documents with -PasswordToOpen or -PasswordToModify.

  .PARAMETER Path
    Specifies the document path literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the document.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the document.

  .PARAMETER Source
    Specifies the source file path for `VBProject.json`.

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
    $Source
  )
  process {
    if (-not $PSCmdlet.ShouldProcess($Path, "Import VBProject from $Source")) {
      return
    }
    if ((Get-Item -LiteralPath $Path -Force).IsReadOnly) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $Path))
    }
    $app = New-WordObject
    try {
      Open-WordFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Action {
        param (
          [Parameter(Mandatory)]
          [Microsoft.Office.Interop.Word.Document]
          $Document
        )
        if ($Document.ReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $Path))
        }
        Import-VBProject -VBProject $Document.VBProject -Source $Source -TargetExe 'WINWORD.EXE'
        $Document.Save()
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

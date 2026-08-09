using module .\VBProject.psm1
using namespace System.IO
using namespace System.Text

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

  .EXAMPLE
    Export-WordVBProject -Path "$env:TEMP\Document.docm" -Destination "$env:TEMP\VBProject.json"

    Exports VBA project metadata and components from a Word document to VBProject.json and creates a companion component folder.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Export-WordVBProject -Path "$env:TEMP\Document.docm" -PasswordToOpen $password -Destination "$env:TEMP\VBProject.json"

    Exports VBA project from a document protected with an open password.

  .EXAMPLE
    Export-WordVBProject -Path "$env:TEMP\Document.docm" -PasswordToModify $password -Destination "$env:TEMP\VBProject.json"

    Exports VBA project from a document protected with a modify password.

  .EXAMPLE
    Export-WordVBProject -Path "$env:TEMP\Document.docm" -Destination "$env:TEMP\VBProject.json" -WhatIf

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
    $app = New-WordObject
    $arguments = @{
      Destination   = $Destination
      ComponentRoot = $ComponentRoot
      Force         = $Force
      NoClobber     = $NoClobber
    }
    try {
      Open-WordFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly -Action {
        param (
          [Microsoft.Office.Interop.Word.Document]
          $Document
        )
        return Export-VBProject -VBProject $Document.VBProject @arguments
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

  .EXAMPLE
    Import-WordVBProject -Path "$env:TEMP\Document.docm" -Source "$env:TEMP\VBProject.json"

    Imports VBA project from JSON file to Word document.

  .EXAMPLE
    $password = Read-Host -AsSecureString -Prompt 'Enter password'
    Import-WordVBProject -Path "$env:TEMP\Document.docm" -Source "$env:TEMP\VBProject.json" -PasswordToOpen $password

    Imports into document protected with an open password.

  .EXAMPLE
    $password = Read-Host -AsSecureString
    Import-WordVBProject -Path "$env:TEMP\Document.docm" -Source "$env:TEMP\VBProject.json" -PasswordToModify $password

    Imports into document protected with a modify password.

  .EXAMPLE
    Import-WordVBProject -Path "$env:TEMP\Document.docm" -Source "$env:TEMP\VBProject.json" -WhatIf

    Shows what would happen if the command were executed without actually importing the VBA project.

  .EXAMPLE
    Import-WordVBProject -Path "$env:TEMP\Document.docm" -Source "$env:TEMP\VBProject.json" -Hidden

    Imports while keeping the Word window hidden.

  .EXAMPLE
    Import-WordVBProject -Path "$env:TEMP\Document.docm" -Source "$env:TEMP\VBProject.json" -Force -Confirm

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

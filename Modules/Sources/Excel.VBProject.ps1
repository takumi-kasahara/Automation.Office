using module .\VBProject.psm1
using namespace System.IO
using namespace System.Text

Set-StrictMode -Version Latest

function Export-ExcelVBProject {
  <#
  .SYNOPSIS
    Exports VBA project metadata and components from an Excel workbook.

  .DESCRIPTION
    Opens a file and exports references and components into a unified `VBProject.json` format.

    The companion component folder is created next to the destination JSON unless -ComponentRoot is specified.
    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.

  .PARAMETER Path
    Specifies the workbook path literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

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
    $app = New-ExcelObject
    $arguments = @{
      Destination   = $Destination
      ComponentRoot = $ComponentRoot
      Force         = $Force
      NoClobber     = $NoClobber
    }
    try {
      Open-ExcelFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -ReadOnly -Action {
        param(
          [Parameter(Mandatory)]
          [Microsoft.Office.Interop.Excel.Workbook]
          $Workbook
        )
        return Export-VBProject -VBProject $Workbook.VBProject @arguments
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
function Import-ExcelVBProject {
  <#
  .SYNOPSIS
    Imports VBA project metadata and components into an Excel workbook.

  .DESCRIPTION
    Opens a file and imports references and components from unified `VBProject.json` format.

    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.
    You can import into protected workbooks with -PasswordToOpen or -PasswordToModify.

  .PARAMETER Path
    Specifies the workbook path literally. Wildcards are not interpreted.

  .PARAMETER PasswordToOpen
    Specifies the password required to open the workbook.

  .PARAMETER PasswordToModify
    Specifies the password required to modify the workbook.

  .PARAMETER Source
    Specifies the source file path for `VBProject.json`.

  .PARAMETER Force
    Forces import processing when a file is marked as read-only recommended.

  .PARAMETER Hidden
    Imports while keeping the workbook window hidden.

  .OUTPUTS
    None.

  .NOTES
    If a file is read-only recommended and -Force is not specified, the cmdlet throws an exception.
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
    $Force,
    [switch]
    $Hidden
  )
  process {
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($Path, "Import VBProject from $Source"))) {
      return
    }
    if ((Get-Item -LiteralPath $Path -Force).IsReadOnly) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $Path))
    }
    $app = New-ExcelObject
    try {
      Open-ExcelFile -Application $app -Path $Path -PasswordToOpen $PasswordToOpen -PasswordToModify $PasswordToModify -Force:$Force -Action {
        param(
          [Parameter(Mandatory)]
          [Microsoft.Office.Interop.Excel.Workbook]
          $file
        )
        if ($file.ReadOnly) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'FileIsReadOnly' -TargetObject $Path))
        }
        if ($file.Windows.Count -gt 0) {
          $file.Windows(1).Visible = -not $Hidden
        }
        Import-VBProject -VBProject $file.VBProject -Source $Source -TargetExe 'EXCEL.EXE'
        $file.Save()
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

using assembly Microsoft.Office.Interop.Access
using module .\VBProject.psm1
using namespace System.IO
using namespace System.Text

Set-StrictMode -Version Latest

function Export-AccessVBProject {
  <#
  .SYNOPSIS
    Exports VBA project metadata and components from an Access database.

  .DESCRIPTION
    Opens a database and exports references and components into a unified `VBProject.json` format.

    The companion component folder is created next to the destination JSON unless -ComponentRoot is specified.
    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.

  .PARAMETER Path
    Specifies the database path literally. Wildcards are not interpreted.

  .PARAMETER Password
    Specifies the password required to open the database.

  .PARAMETER Destination
    Specifies the export file path for `VBProject.json`.

  .PARAMETER ComponentRoot
    Specifies the directory path where component files are exported.
    When omitted, the directory is derived from -Destination by removing its extension.

  .PARAMETER Force
    Overrides read-only destination file attributes.

  .PARAMETER NoClobber
    Produces an error if -Destination already exists. If omitted, existing output is overwritten.

  .PARAMETER IncludeAcObjectType
    Specifies an array of AcObjectType values to include when exporting the VBProject.
    When specified, only the specified object types are included in the export.
    Corresponds to the IncludeAcObjectType parameter in the Access VBA object model.

  .PARAMETER ExcludeAcObjectType
    Specifies an array of AcObjectType values to exclude when exporting the VBProject.
    When specified, the specified object types are omitted from the export.
    Corresponds to the ExcludeAcObjectType parameter in the Access VBA object model.

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
    $Password = $null,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    [string]
    $ComponentRoot,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [Microsoft.Office.Interop.Access.AcObjectType[]]
    $IncludeAcObjectType,
    [Microsoft.Office.Interop.Access.AcObjectType[]]
    $ExcludeAcObjectType
  )
  process {
    if (-not (($Force -and -not $WhatIfPreference) -or $PSCmdlet.ShouldProcess($Path, "Export VBProject to $Destination"))) {
      return
    }
    $app = New-AccessObject
    $arguments = @{
      Destination         = $Destination
      ComponentRoot       = $ComponentRoot
      Force               = $Force
      NoClobber           = $NoClobber
      IncludeAcObjectType = $IncludeAcObjectType
      ExcludeAcObjectType = $ExcludeAcObjectType
    }
    try {
      Open-AccessFile -Application $app -Path $Path -Password $Password
      try {
        return Export-VBProject -Application $app @arguments
      } finally {
        $app.CloseCurrentDatabase()
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
function Import-AccessVBProject {
  <#
  .SYNOPSIS
    Imports VBA project metadata and components into an Access database.

  .DESCRIPTION
    Opens a database and imports references and components from unified `VBProject.json` format.

    This cmdlet supports ShouldProcess, so you can use -WhatIf and -Confirm.

  .PARAMETER Path
    Specifies the database path literally. Wildcards are not interpreted.

  .PARAMETER Password
    Specifies the password required to open the database.

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
    $Password = $null,
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
    $app = New-AccessObject
    try {
      Open-AccessFile -Application $app -Path $Path -Password $Password
      try {
        Import-VBProject -Application $app -Source $Source -TargetExe 'MSACCESS.EXE'
      } finally {
        $app.CloseCurrentDatabase()
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

using assembly Microsoft.Office.Interop.Access
using assembly Microsoft.Vbe.Interop
using module .\DialogSuppressor.psm1
using namespace Microsoft.Office.Interop.Access
using namespace Microsoft.Office.Interop.Access.Dao
using namespace Microsoft.Vbe.Interop
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Runtime.InteropServices
using namespace System.Text

Set-StrictMode -Version Latest

# https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.vbcomponent?view=office-pia
class VBComponentInfo {
  [string]$Name
  [string]$Path
  [string]$Type
}
# https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.reference?view=office-pia
class ReferenceInfo {
  [bool]$BuiltIn
  [string]$Description
  [string]$FullPath
  [string]$Guid
  [bool]$IsBroken
  [int]$Major
  [int]$Minor
  [string]$Name
  [int]$Type
}
function Get-VBProjectReference {
  [CmdletBinding()]
  [OutputType([ReferenceInfo[]])]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $VBProject
  )
  try {
    return [ReferenceInfo[]]@(
      $VBProject.References |
      ForEach-Object {
        [ReferenceInfo]@{
          BuiltIn     = $_.BuiltIn
          Description = $_.Description
          FullPath    = $_.FullPath
          Guid        = $_.Guid
          IsBroken    = $_.IsBroken
          Major       = $_.Major
          Minor       = $_.Minor
          Name        = $_.Name
          Type        = $_.Type
        }
      }
    )
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Import-VBProjectReference {
  [CmdletBinding()]
  [OutputType([void])]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $VBProject,
    [Parameter(Mandatory)]
    [ReferenceInfo[]]
    $References
  )
  $activity = 'Importing References'
  try {
    $VBProject.References |
    ForEach-Object {
      if (-not $_.BuiltIn -and -not $_.IsBroken) {
        $description = $_.Description
        Write-Progress -Activity $activity -Status "Removing: $description"
        try {
          $VBProject.References.Remove($_)
          "Removed:`t$description" | Out-Host
        } catch [COMException] {
          Write-Warning -Message $_.Exception.Message
        }
      }
    }
    $References |
    Where-Object {
      $isBuiltIn = $_.PSObject.Properties.Match('BuiltIn').Count -gt 0 -and $_.BuiltIn
      $isBroken = $_.PSObject.Properties.Match('IsBroken').Count -gt 0 -and $_.IsBroken
      -not ($isBuiltIn -or $isBroken)
    } |
    ForEach-Object {
      $description = $_.Description
      Write-Progress -Activity $activity -Status "Importing: $description"
      # https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/addfromguid-method-vba-add-in-object-model
      try {
        $VBProject.References.AddFromGuid($_.Guid, [int]$_.Major, [int]$_.Minor) | Write-Verbose
        "Imported:`t$description" | Out-Host
      } catch [COMException] {
        Write-Warning -Message $_.Exception.Message
      }
    }
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Progress -Activity $activity -Completed
  }
}
function Get-AccessObject {
  [CmdletBinding()]
  [OutputType([__ComObject[]])]
  param (
    [Parameter(Mandatory)]
    [__ComObject]
    $Application
  )
  # https://learn.microsoft.com/en-us/office/vba/api/access.acobjecttype
  # https://learn.microsoft.com/en-us/office/vba/api/access.accessobject
  $objects = @()
  try {
    # https://learn.microsoft.com/en-us/office/vba/api/access.currentproject
    if ($Application.CurrentProject) {
      $project = $Application.CurrentProject
      if ($project.AllForms -and $project.AllForms.Count -gt 0) {
        $objects += $project.AllForms
      }
      if ($project.AllReports -and $project.AllReports.Count -gt 0) {
        $objects += $project.AllReports
      }
      if ($project.AllMacros -and $project.AllMacros.Count -gt 0) {
        $objects += $project.AllMacros
      }
      if ($project.AllModules -and $project.AllModules.Count -gt 0) {
        $objects += $project.AllModules
      }
    }
    # https://learn.microsoft.com/en-us/office/vba/api/access.currentdata
    if ($Application.CurrentData) {
      $data = $Application.CurrentData
      if ($data.AllTables -and $data.AllTables.Count -gt 0) {
        $objects += $data.AllTables
      }
      if ($data.AllFunctions -and $data.AllFunctions.Count -gt 0) {
        $objects += $data.AllFunctions
      }
      if ($data.AllQueries -and $data.AllQueries.Count -gt 0) {
        $objects += $data.AllQueries
      }
      if ($data.AllViews -and $data.AllViews.Count -gt 0) {
        $objects += $data.AllViews
      }
      if ($data.AllStoredProcedures -and $data.AllStoredProcedures.Count -gt 0) {
        $objects += $data.AllStoredProcedures
      }
      if ($data.AllDatabaseDiagrams -and $data.AllDatabaseDiagrams.Count -gt 0) {
        $objects += $data.AllDatabaseDiagrams
      }
    }
    return $objects |
    Where-Object -Property Type -NE ([AcObjectType]::acModule) |
    Where-Object { -not ($_.Type -eq [AcObjectType]::acTable -and $_.Name -like 'MSys*') <# system tables #> }
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Export-VBProjectComponent {
  [CmdletBinding()]
  [OutputType([VBComponentInfo[]])]
  param (
    [__ComObject]
    $Application,
    [Parameter(Mandatory)]
    [__ComObject]
    $VBProject,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Leaf) })]
    [string]
    $Destination,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [Microsoft.Office.Interop.Access.AcObjectType[]]
    $IncludeAcObjectType,
    [Microsoft.Office.Interop.Access.AcObjectType[]]
    $ExcludeAcObjectType
  )
  $activity = 'Getting Components'
  try {
    $resolved = [Path]::GetFullPath($Destination)
    if ((Test-Path -LiteralPath $resolved) -and $NoClobber) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
    }
    if (-not (Test-Path -LiteralPath $resolved)) {
      New-Item -Path $resolved -ItemType Directory | Out-Null
    }
    $exported = [VBComponentInfo[]]@()
    if ($Application) {
      $included = $true
      if ($includeAcObjectType) {
        $included = $IncludeAcObjectType -contains [AcObjectType]::acTable
        if ($included -and $ExcludeAcObjectType) {
          $included = $included -and -not ($ExcludeAcObjectType -contains [AcObjectType]::acTable)
        }
      } elseif ($ExcludeAcObjectType) {
        $included = -not ($ExcludeAcObjectType -contains [AcObjectType]::acTable)
      } else {
        $included = $true
      }
      if ($included) {
        $tableDefs = $resolved | Join-Path -ChildPath 'TableDefs.accdb'
        New-AccessFile -Path $tableDefs -RemovePersonalInformation -Force | Out-Null
        $exported += [VBComponentInfo]@{
          Name = 'TableDefs'
          Path = [PathCompatibility]::GetRelativePath($resolved, $tableDefs)
          type = [AcObjectType]::acTable
        }
      }
      Get-AccessObject -Application $Application |
      Where-Object {
        $included = $true
        if ($IncludeAcObjectType) {
          $included = $IncludeAcObjectType -contains [AcObjectType]$_.Type
          if ($included -and $ExcludeAcObjectType) {
            $included = $included -and -not ($ExcludeAcObjectType -contains [AcObjectType]$_.Type)
          }
        } elseif ($ExcludeAcObjectType) {
          $included = -not ($ExcludeAcObjectType -contains [AcObjectType]$_.Type)
        } else {
          $included = $true
        }
        return $included
      } |
      ForEach-Object {
        $filename = "$($_.Name).txt"
        $path = $resolved | Join-Path -ChildPath $fileName
        $isReadOnly = $false
        if (Test-Path -LiteralPath $path) {
          if ((Get-Item -LiteralPath $path -Force).PsIsContainer) {
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $path))
          }
          if ($NoClobber) {
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $path))
          }
          $isReadOnly = (Get-Item -LiteralPath $path -Force).IsReadOnly
          if ($isReadOnly -and $Force) {
            (Get-Item -LiteralPath $path -Force).IsReadOnly = $false
          }
        }
        Write-Progress -Activity $activity -Status "Exporting: $($_.Name) to $path as $([AcObjectType]$_.Type)"
        try {
          # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text
          if ([AcObjectType]$_.Type -eq [AcObjectType]::acTable) {
            # https://learn.microsoft.com/en-us/office/vba/api/access.docmd.transferdatabase
            $Application.DoCmd.TransferDatabase(
              [AcDataTransferType]::acExport  # TransferType
              , 'Microsoft Access'            # DatabaseType
              , $tableDefs                    # DatabaseName
              , [AcObjectType]$_.Type         # ObjectType
              , $_.Name                       # Source
              , $_.Name                       # Destination
              , $true                         # StructureOnly
            )
            # https://learn.microsoft.com/en-us/office/vba/api/access.docmd.transfertext
            # NOTE stringified FileName to ensure correct type for COM interop.
            $Application.DoCmd.TransferText(
              [AcTextTransferType]::acExportDelim # TransferType
              , [type]::Missing                   # SpecificationName
              , $_.Name                           # TableName
              , "$path"                           # FileName
              , $true                             # HasFieldNames
              , [type]::Missing                   # HTMLTableName
              , 1200                              # CodePage
            )
          } else {
            # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text
            $Application.SaveAsText(
              [AcObjectType]$_.Type # ObjectType
              , $_.Name             # ObjectName
              , "$path"             # FileName
            )
          }
        } finally {
          if ($isReadOnly -and $Force) {
            (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
          }
        }
        $exported += [VBComponentInfo]@{
          Name = $_.Name
          Path = [PathCompatibility]::GetRelativePath($resolved, $path)
          Type = [AcObjectType]$_.Type
        }
        "Exported:`t$($_.Name) as $([AcObjectType]$_.Type) to $($path)" | Out-Host
      }
    }
    $VBProject.VBComponents |
    ForEach-Object {
      $extension = Convert-ComponentTypeToExtension -ComponentType $_.Type
      $fileName = "$($_.Name)$extension"
      $path = $resolved | Join-Path -ChildPath $fileName
      $isReadOnly = $false
      if (Test-Path -LiteralPath $path) {
        if ((Get-Item -LiteralPath $path -Force).PsIsContainer) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $path))
        }
        if ($NoClobber) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $path))
        }
        $isReadOnly = (Get-Item -LiteralPath $path -Force).IsReadOnly
        if ($isReadOnly -and $Force) {
          (Get-Item -LiteralPath $path -Force).IsReadOnly = $false
        }
      }
      Write-Progress -Activity $activity -Status "Exporting: $($_.Name) to $path as $([vbext_ComponentType]$_.Type)"
      try {
        if ([vbext_ComponentType]$_.Type -eq [vbext_ComponentType]::vbext_ct_Document) {
          $contents = $_.CodeModule.Lines(1, $_.CodeModule.CountOfLines) -join [Environment]::NewLine
          if (-not $contents.EndsWith([Environment]::NewLine)) {
            $contents += [Environment]::NewLine
          }
          [File]::WriteAllText($path, $contents, [UTF8Encoding]::new($false))
        } else {
          $_.Export($path)
        }
      } finally {
        if ($isReadOnly -and $Force) {
          (Get-Item -LiteralPath $path -Force).IsReadOnly = $true
        }
      }
      $exported += [VBComponentInfo]@{
        Name = $_.Name
        Path = [PathCompatibility]::GetRelativePath($resolved, $path)
        Type = [vbext_ComponentType]$_.Type
      }
      "Exported:`t$($_.Name) as $([vbext_ComponentType]$_.Type) to $($path)" | Out-Host
    }
    return $exported
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Progress -Activity $activity -Completed
  }
}
function Import-VBProjectComponent {
  [CmdletBinding()]
  [OutputType([void])]
  param (
    [__ComObject]
    $Application,
    [Parameter(Mandatory)]
    [__ComObject]
    $VBProject,
    [Parameter(Mandatory)]
    [VBComponentInfo[]]
    $Components,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]
    $ComponentRoot
  )
  $activity = 'Importing Components'
  try {
    $root = [Path]::GetFullPath($ComponentRoot)
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemNotFound' -TargetObject $root))
    }
    if ($Application) {
      $tableDefs = $Components |
      Where-Object -Property Type -EQ ([AcObjectType]::acTable) |
      Where-Object { [Path]::GetFileName($_.Path) -eq 'TableDefs.accdb' } |
      Where-Object { Test-Path -LiteralPath ([Path]::GetFullPath(($root | Join-Path -ChildPath $_.Path))) -PathType Leaf } |
      ForEach-Object { [Path]::GetFullPath(($root | Join-Path -ChildPath $_.Path)) } |
      Select-Object -First 1
      $Components |
      Where-Object -Property Type -NE ([AcObjectType]::acModule) |
      Where-Object -Property Type -NotIn ([Enum]::GetValues([vbext_ComponentType])) |
      Where-Object { $_.Type -ne [AcObjectType]::acTable -or [Path]::GetFileName($_.Path) -ne 'TableDefs.accdb' } |
      ForEach-Object {
        $_.Path = [string]$_.Path
        $path = Join-Path -Path $root -ChildPath $_.Path
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
          $fallbackPath = Join-Path -Path $root -ChildPath ([Path]::GetFileName($_.Path))
          if (Test-Path -LiteralPath $fallbackPath -PathType Leaf) {
            $path = $fallbackPath
          }
        }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
          Write-Warning -Message "Skipping import of $($_.Name) as $([AcObjectType]$_.Type) because the file $path does not exist."
          return
        }
        $tempDir = $null
        $importPath = $path
        if ([Path]::GetExtension($path) -in '.bas', '.cls', '.frm', '.vba') {
          $content = [File]::ReadAllText($path)
          if ($content -match '(?<!\r)\n') {
            # Normalize line endings to CRLF and ensure UTF-8 encoding without BOM for proper COM interop
            $normalized = $content -replace "`r`n|`n|`r", "`r`n"
            $tempDir = $env:TEMP | Join-Path -ChildPath ([Path]::GetRandomFileName())
            New-Item -Path $tempDir -ItemType Directory -Force -WhatIf:$false -Confirm:$false | Out-Null
            $tempFile = $tempDir | Join-Path -ChildPath ([Path]::GetFileName($path))
            [File]::WriteAllText($tempFile, $normalized, [UTF8Encoding]::new($false))
            $importPath = $tempFile
          }
        }
        try {
          Write-Progress -Activity $activity -Status "Importing: $($_.Name) from $_.Path as $([AcObjectType]$_.Type)"
          try {
            # https://learn.microsoft.com/en-us/office/vba/api/access.docmd.deleteobject
            $Application.DoCmd.DeleteObject([AcObjectType]$_.Type, $_.Name)
            "Removed:`t$($_.Name) as $([AcObjectType]$_.Type)" | Out-Host
          } catch [COMException] {
            Write-Warning -Message $_.Exception.Message
          }
          if ([AcObjectType]$_.Type -eq [AcObjectType]::acTable) {
            try {
              if ($null -ne $tableDefs -and (Test-Path -LiteralPath $tableDefs -PathType Leaf)) {
                # https://learn.microsoft.com/en-us/office/vba/api/access.docmd.transferdatabase
                $Application.DoCmd.TransferDatabase(
                  [AcDataTransferType]::acImport  # TransferType
                  , 'Microsoft Access'            # DatabaseType
                  , $tableDefs                    # DatabaseName
                  , [AcObjectType]$_.Type         # ObjectType
                  , $_.Name                       # Source
                  , $_.Name                       # Destination
                  , $true                         # StructureOnly
                )
                "Imported:`t$($_.Name) as $([AcObjectType]$_.Type) from $($tableDefs)" | Out-Host
              }
              # https://learn.microsoft.com/en-us/office/vba/api/access.docmd.transfertext
              $Application.DoCmd.TransferText(
                [AcTextTransferType]::acImportDelim # TransferType
                , [type]::Missing                   # SpecificationName
                , $_.Name                           # TableName
                , "$importPath"                     # FileName
                , $true                             # HasFieldNames
                , [type]::Missing                   # HTMLTableName
                , 1200                              # CodePage
              )
            } catch [COMException] {
              Write-Warning -Message $_.Exception.Message
            }
          } else {
            # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-load-from-text
            # NOTE stringified FileName to ensure correct type for COM interop.
            $Application.LoadFromText(
              [AcObjectType]$_.Type # ObjectType
              , $_.Name             # ObjectName
              , "$importPath"       # FileName
            )
          }
          "Imported:`t$($_.Name) as $([AcObjectType]$_.Type) from $($importPath)" | Out-Host
        } finally {
          if ($null -ne $tempDir -and (Test-Path -LiteralPath $tempDir -PathType Container)) {
            Remove-Item -LiteralPath $tempDir -Recurse -Force -WhatIf:$false -Confirm:$false
          }
        }
      }
    }
    $VBProject.VBComponents |
    Where-Object { [vbext_ComponentType]$_.Type -ne [vbext_ComponentType]::vbext_ct_Document } |
    ForEach-Object {
      Write-Progress -Activity $activity -Status "Removing: $($_.Name)"
      $VBProject.VBComponents.Remove($_)
      "Removed:`t$($_.Name) as $([vbext_ComponentType]$_.Type)" | Out-Host
    }
    $component = $null
    $components |
    Where-Object -Property Type -In ([Enum]::GetValues([vbext_ComponentType])) |
    ForEach-Object {
      $path = Join-Path -Path $root -ChildPath $_.Path
      if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $fallbackPath = Join-Path -Path $root -ChildPath ([Path]::GetFileName($_.Path))
        if (Test-Path -LiteralPath $fallbackPath -PathType Leaf) {
          $path = $fallbackPath
        }
      }
      if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemNotFound' -TargetObject $path))
      }
      $tempDir = $null
      $importPath = $path
      if ([Path]::GetExtension($path) -in '.bas', '.cls', '.frm', '.vba') {
        $content = [File]::ReadAllText($path)
        if ($content -match '(?<!\r)\n') {
          $normalized = $content -replace "`r`n|`n|`r", "`r`n"
          $tempDir = $env:TEMP | Join-Path -ChildPath ([Path]::GetRandomFileName())
          New-Item -Path $tempDir -ItemType Directory -Force -WhatIf:$false -Confirm:$false | Out-Null
          $tempFile = $tempDir | Join-Path -ChildPath ([Path]::GetFileName($path))
          [File]::WriteAllText($tempFile, $normalized, [UTF8Encoding]::new($false))
          if ([Path]::GetExtension($path) -eq '.frm') {
            $frxPath = [Path]::ChangeExtension($path, '.frx')
            if (Test-Path -LiteralPath $frxPath -PathType Leaf) {
              Copy-Item -LiteralPath $frxPath -Destination ($tempDir | Join-Path -ChildPath ([Path]::GetFileName($frxPath))) -Force -WhatIf:$false -Confirm:$false
            }
          }
          $importPath = $tempFile
        }
      }
      try {
        Write-Progress -Activity $activity -Status "Importing: $($_.Name) from $($_.Path) as $([vbext_ComponentType]$_.Type)"
        if ([vbext_ComponentType]$_.Type -ne [vbext_ComponentType]::vbext_ct_Document) {
          $component = $VBProject.VBComponents.Import($importPath)
          $component.Name = $_.Name
          if ($Application) {
            # https://learn.microsoft.com/en-us/office/vba/api/access.docmd.save
            $Application.DoCmd.Save([AcObjectType]::acModule, $component.Name)
          }
        } else {
          $component = $VBProject.VBComponents.Item($_.Name)
          $codeModule = $component.CodeModule
          $codeModule.DeleteLines(1, $codeModule.CountOfLines)
          $codeModule.AddFromFile($importPath)
        }
      } finally {
        if ($null -ne $tempDir -and (Test-Path -LiteralPath $tempDir -PathType Container)) {
          Remove-Item -LiteralPath $tempDir -Recurse -Force -WhatIf:$false -Confirm:$false
        }
      }
    }
    if ($component) {
      "Imported:`t$($component.Name) as $([vbext_ComponentType]$component.Type) from $($importPath)" | Out-Host
    }
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Progress -Activity $activity -Completed
  }
}
function Export-VBProject {
  [CmdletBinding(DefaultParameterSetName = 'DefaultSet')]
  [OutputType([System.IO.FileInfo])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'AccessSet')]
    [__ComObject]
    $Application,
    [Parameter(Mandatory, ParameterSetName = 'DefaultSet')]
    [__ComObject]
    $VBProject,
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
    [Parameter(ParameterSetName = 'AccessSet')]
    [Microsoft.Office.Interop.Access.AcObjectType[]]
    $IncludeAcObjectType,
    [Parameter(ParameterSetName = 'AccessSet')]
    [Microsoft.Office.Interop.Access.AcObjectType[]]
    $ExcludeAcObjectType
  )
  $activity = 'Exporting VBProject'
  $resolvedDestination = $null
  try {
    $resolvedDestination = [Path]::GetFullPath($Destination)
    if (Test-Path -LiteralPath $resolvedDestination -PathType Container) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolvedDestination))
    }
    if ($NoClobber -and (Test-Path -LiteralPath $resolvedDestination -PathType Leaf)) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolvedDestination))
    }
    $resolvedComponentRoot = if ([string]::IsNullOrEmpty($ComponentRoot)) {
      $destinationDirectory = [Path]::GetDirectoryName($resolvedDestination)
      if ([string]::IsNullOrEmpty($destinationDirectory)) {
        $destinationDirectory = [Environment]::CurrentDirectory
      }
      $destinationName = [Path]::GetFileNameWithoutExtension($resolvedDestination)
      $destinationDirectory | Join-Path -ChildPath $destinationName
    } else {
      [Path]::GetFullPath($ComponentRoot)
    }

    Write-Progress -Activity $activity -Status 'Exporting references'
    $references = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'AccessSet' {
        @(Get-VBProjectReference -VBProject $Application.VBE.ActiveVBProject)
      }
      'DefaultSet' {
        @(Get-VBProjectReference -VBProject $VBProject)
      }
    }
    Write-Progress -Activity $activity -Status 'Exporting components'
    $components = switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'AccessSet' {
        @(Export-VBProjectComponent -Application $Application -VBProject $Application.VBE.ActiveVBProject -Destination $resolvedComponentRoot -Force:$Force -NoClobber:$NoClobber -IncludeAcObjectType $IncludeAcObjectType -ExcludeAcObjectType $ExcludeAcObjectType)
      }
      'DefaultSet' {
        @(Export-VBProjectComponent -VBProject $VBProject -Destination $resolvedComponentRoot -Force:$Force -NoClobber:$NoClobber)
      }
    }

    $destinationDirectory = [Path]::GetDirectoryName($resolvedDestination)
    if ([string]::IsNullOrEmpty($destinationDirectory)) {
      $destinationDirectory = [Environment]::CurrentDirectory
    }
    $components = @($components) |
    ForEach-Object {
      $absoluteComponentPath = [Path]::GetFullPath(($resolvedComponentRoot | Join-Path -ChildPath $_.Path))
      [VBComponentInfo]@{
        Name = $_.Name
        Path = [PathCompatibility]::GetRelativePath($destinationDirectory, $absoluteComponentPath)
        Type = $_.Type
      }
    }
    $metadata = [PSCustomObject]@{
      VBComponents = @($components)
      References   = @($references)
    }
    $metadata |
    ConvertTo-Json |
    Out-File -LiteralPath $resolvedDestination -Encoding utf8 -Force:$Force
    return Get-Item -LiteralPath $resolvedDestination -Force
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Progress -Activity $activity -Completed
  }
}
function Import-VBProject {
  [CmdletBinding(DefaultParameterSetName = 'DefaultSet')]
  [OutputType([void])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'AccessSet')]
    [__ComObject]
    $Application,
    [Parameter(Mandatory, ParameterSetName = 'DefaultSet')]
    [__ComObject]
    $VBProject,
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]
    $Source,
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $TargetExe
  )
  $dialogSuppressor = Start-DialogSuppressor -TargetExe $TargetExe
  $activity = 'Importing VBProject'
  try {
    $resolvedSource = [Path]::GetFullPath($Source)
    $sourceDirectory = $resolvedSource | Split-Path -Parent

    $metadata = Get-Content -LiteralPath $resolvedSource -Encoding UTF8 -Raw | ConvertFrom-Json
    $components = @($metadata.VBComponents) |
    ForEach-Object {
      [VBComponentInfo]@{
        Name = $_.Name
        Path = $_.Path
        Type = $_.Type
      }
    }
    Write-Progress -Activity $activity -Status 'Importing references'
    switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'AccessSet' {
        Import-VBProjectReference -VBProject $Application.VBE.ActiveVBProject -References @($metadata.References)
      }
      'DefaultSet' {
        Import-VBProjectReference -VBProject $VBProject -References @($metadata.References)
      }
    }
    Write-Progress -Activity $activity -Status 'Importing components'
    switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'AccessSet' {
        Import-VBProjectComponent -Application $Application -VBProject $Application.VBE.ActiveVBProject -Components $components -ComponentRoot $sourceDirectory
      }
      'DefaultSet' {
        Import-VBProjectComponent -VBProject $VBProject -Components $components -ComponentRoot $sourceDirectory
      }
    }
  } finally {
    if ($dialogSuppressor) {
      Stop-DialogSuppressor -Job $dialogSuppressor
    }
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Progress -Activity $activity -Completed
  }
}
function Convert-ComponentTypeToExtension {
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory)]
    [vbext_ComponentType]
    $ComponentType
  )
  # https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.vbext_componenttype?view=office-pia
  switch ([vbext_ComponentType]::$ComponentType) {
    { $_ -eq [vbext_ComponentType]::vbext_ct_StdModule } {
      '.bas'
    }
    { $_ -eq [vbext_ComponentType]::vbext_ct_ClassModule } {
      '.cls'
    }
    { $_ -eq [vbext_ComponentType]::vbext_ct_MSForm } {
      '.frm'
    }
    { $_ -eq [vbext_ComponentType]::vbext_ct_Document } {
      '.vba'
    }
    default {
      throw [ArgumentException]::new("Undefined:$ComponentType")
    }
  }
}
function Convert-VBProjectComponentToCrLf {
  [CmdletBinding()]
  [OutputType([void])]
  param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]
    $Source
  )
  Get-ChildItem -LiteralPath $Source -File |
  ForEach-Object {
    $item = $_
    if ($item.Extension -notin '.bas', '.cls', '.frm', '.vba') {
      return
    }
    $content = [File]::ReadAllText($item.FullName)
    $normalized = $content -replace "`r`n|`n|`r", "`r`n"
    [File]::WriteAllText($item.FullName, $normalized, [UTF8Encoding]::new($false))
  }
}

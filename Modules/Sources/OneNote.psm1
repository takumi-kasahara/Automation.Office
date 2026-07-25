using namespace Microsoft.Office.Interop.OneNote
using namespace System.IO
using namespace System.Runtime.InteropServices
using namespace System.Xml

Add-Type -AssemblyName Microsoft.Office.Interop.OneNote
Set-StrictMode -Version Latest

#region Public
function New-OneNoteNotebook {
  <#
  .SYNOPSIS
    Creates a new OneNote notebook.

  .DESCRIPTION
    Creates a new OneNote notebook at the specified file system path using the OneNote COM API.
    If the directory does not exist, it will be created automatically.
    Returns the object ID of the created notebook.
    Supports -WhatIf and -Confirm to preview or confirm the operation.

  .PARAMETER Path
    Required. The file system path where the new notebook directory will be created.
    The path must be a valid directory path.

  .EXAMPLE
    $notebookPath = $env:TEMP | Join-Path -ChildPath 'MyNotebook'
    $notebookId = New-OneNoteNotebook -Path $notebookPath
    $notebookId

    Creates a new OneNote notebook and returns its object ID.

  .EXAMPLE
    $notebookId = New-OneNoteNotebook -Path 'C:\Notes\Work'
    $sectionId = New-OneNoteSection -Path 'C:\Notes\Work\Tasks.one' -ParentId $notebookId

    Creates a notebook and then adds a section to it using ParentId.

  .OUTPUTS
    System.String
      The OneNote object ID of the created notebook.

  .NOTES
    This function uses the OneNote COM API (OpenHierarchy) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([string])]
  param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -IsValid })]
    [string]
    $Path
  )
  $app = Get-OneNoteApplication
  try {
    if (-not $PSCmdlet.ShouldProcess($Path, 'New Notebook')) {
      return
    }
    if (-not (Test-Path -LiteralPath $Path)) {
      New-Item -Path $Path -ItemType Directory -WhatIf:$WhatIfPreference -Confirm:$false | Out-Null
    }
    $objectId = $null
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#openhierarchy-method
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#createfiletype
    $app.OpenHierarchy(
      $Path                           # BSTR            bstrPath
      , [string]::Empty               # BSTR            bstrRelativeToObjectID
      , [ref]$objectId                # BSTR*           pbstrObjectID
      , [CreateFileType]::cftNotebook # CreateFileType  cftIfNotExist
    )
    return $objectId
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function New-OneNoteSection {
  <#
  .SYNOPSIS
    Creates a new OneNote section file.

  .DESCRIPTION
    Creates a new OneNote section (.one) file at the specified path using the OneNote COM API.
    Returns the object ID of the created section.
    Supports -WhatIf and -Confirm to preview or confirm the operation.

  .PARAMETER Path
    Required. The file system path where the new section file will be created.
    The path must be a valid leaf path. The file does not need to exist beforehand.

  .PARAMETER ParentId
    Required only in ParentIdSet. The OneNote object ID of the parent notebook or section group.
    When specified, only the file name part of Path is used for OpenHierarchy.

  .EXAMPLE
    $sectionId = New-OneNoteSection -Path 'C:\Notes\MySection.one'
    $sectionId

    Creates a new OneNote section at the given path and returns its object ID.

  .EXAMPLE
    $path = $env:TEMP | Join-Path -ChildPath 'Section.one'
    $sectionId = New-OneNoteSection -Path $path
    $pageId = New-OneNotePage -SectionId $sectionId

    Creates a section (PathSet default), then creates a page within that section.

  .EXAMPLE
    $notebookPath = $env:TEMP | Join-Path -ChildPath 'MyNotebook'
    $notebookId = New-OneNoteNotebook -Path $notebookPath
    $sectionPath = $notebookPath | Join-Path -ChildPath 'Section.one'
    $sectionId = New-OneNoteSection -Path $sectionPath -ParentId $notebookId

    Creates a section by ParentIdSet under the specified notebook.

  .OUTPUTS
    System.String
      The OneNote object ID of the created section.

  .NOTES
    This function uses the OneNote COM API (OpenHierarchy) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'PathSet')]
  [OutputType([string])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'ParentIdSet', Position = 0)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Path,
    [Parameter(Mandatory, ParameterSetName = 'ParentIdSet')]
    [string]
    $ParentId
  )
  $app = Get-OneNoteApplication
  try {
    switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'PathSet' {
        $resolved = [Path]::GetFullPath($Path)
        $relativeToObjectId = [string]::Empty
      }
      'ParentIdSet' {
        $resolved = [Path]::GetFileName($Path)
        $relativeToObjectId = $ParentId
      }
    }
    if (-not $PSCmdlet.ShouldProcess($Path, 'New Section')) {
      return
    }
    $objectId = $null
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#openhierarchy-method
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#createfiletype
    $app.OpenHierarchy(
      $resolved                       # BSTR            bstrPath
      , $relativeToObjectId           # BSTR            bstrRelativeToObjectID
      , [ref]$objectId                # BSTR*           pbstrObjectID
      , [CreateFileType]::cftSection  # CreateFileType  cftIfNotExist
    )
    return $objectId
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function New-OneNoteSectionGroup {
  <#
  .SYNOPSIS
    Creates a new OneNote section group.

  .DESCRIPTION
    Creates a new OneNote section group (folder) at the specified path using the OneNote COM API.
    Returns the object ID of the created section group.
    Supports -WhatIf and -Confirm to preview or confirm the operation.

    In PathSet (default), the full path is resolved and passed directly to OpenHierarchy.
    In ParentIdSet, only the leaf folder name is used, relative to the parent object ID.

  .PARAMETER Path
    Required. The file system path of the new section group directory.
    The section group name must not contain a dot (.), as OneNote COM rejects such names.
    The path must be a valid leaf path.

  .PARAMETER ParentId
    Required only in ParentIdSet. The OneNote object ID of the parent notebook or section group.
    When specified, only the folder name part of Path is used for OpenHierarchy.

  .EXAMPLE
    $notebookPath = $env:TEMP | Join-Path -ChildPath 'MyNotebook'
    $notebookId = New-OneNoteNotebook -Path $notebookPath
    $groupPath = $notebookPath | Join-Path -ChildPath 'ProjectGroup'
    $groupId = New-OneNoteSectionGroup -Path $groupPath
    $groupId

    Creates a section group by path and returns its object ID.

  .EXAMPLE
    $groupPath = $notebookPath | Join-Path -ChildPath 'ProjectGroup'
    $groupId = New-OneNoteSectionGroup -Path $groupPath -ParentId $notebookId

    Creates a section group under the specified notebook by parent ID.

  .OUTPUTS
    System.String
      The OneNote object ID of the created section group.

  .NOTES
    This function uses the OneNote COM API (OpenHierarchy) internally.
    Section group names must not contain dots to avoid COM error 0x80042018.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'PathSet')]
  [OutputType([string])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'ParentIdSet', Position = 0)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and -not (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Path,
    [Parameter(Mandatory, ParameterSetName = 'ParentIdSet')]
    [string]
    $ParentId
  )
  $app = Get-OneNoteApplication
  try {
    switch -Exact -CaseSensitive ($PSCmdlet.ParameterSetName) {
      'PathSet' {
        $resolved = [Path]::GetFullPath($Path)
        $relativeToObjectId = [string]::Empty
      }
      'ParentIdSet' {
        $resolved = [Path]::GetFileName($Path)
        $relativeToObjectId = $ParentId
      }
    }
    if (-not $PSCmdlet.ShouldProcess($Path, 'New SectionGroup')) {
      return
    }
    $objectId = $null
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#openhierarchy-method
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#createfiletype
    $app.OpenHierarchy(
      $resolved         # BSTR            bstrPath
      , $relativeToObjectId         # BSTR            bstrRelativeToObjectID
      , [ref]$objectId              # BSTR*           pbstrObjectID
      , [CreateFileType]::cftFolder # CreateFileType  cftIfNotExist
    )
    return $objectId
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function New-OneNotePage {
  <#
  .SYNOPSIS
    Creates a new page in a OneNote section.

  .DESCRIPTION
    Creates a new page in the specified OneNote section using the OneNote COM API.
    Returns the object ID of the created page.

  .PARAMETER SectionId
    Required. The object ID of the section in which the new page will be created.
    Obtain this value from New-OneNoteSection or Get-OneNoteHierarchy.

  .EXAMPLE
    $path = $env:TEMP | Join-Path -ChildPath 'NoteBook.one'
    $sectionId = New-OneNoteSection -Path $path
    $pageId = New-OneNotePage -SectionId $sectionId
    $pageId

    Creates a new section, then creates a page within it and returns its object ID.

  .OUTPUTS
    System.String
    The OneNote object ID of the created page.

  .NOTES
    This function uses the OneNote COM API (CreateNewPage) internally.
    The page is created with the default page style (npsDefault).
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([string])]
  param (
    [Parameter(Mandatory)]
    [string]
    $SectionId
  )
  $app = Get-OneNoteApplication
  try {
    if (-not $PSCmdlet.ShouldProcess($SectionId, 'New Page')) {
      return
    }
    $pageId = $null
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#createnewpage-method
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#createfiletype
    $app.CreateNewPage(
      $SectionId                    # BSTR          bstrSectionID
      , [ref]$pageId                # BSTR*         pbstrPageID
      , [NewPageStyle]::npsDefault  # NewPageStyle  npsNewPageStyle
    )
    return $pageId
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Get-OneNoteHierarchy {
  <#
  .SYNOPSIS
    Retrieves the OneNote hierarchy as XML nodes.

  .DESCRIPTION
    Retrieves the hierarchy of notebooks, section groups, sections, or pages
    starting from the specified object using the OneNote COM API.
    Returns the parsed XML document element.

  .PARAMETER Id
    Optional. The OneNote object ID of the starting node.
    If omitted or empty, the hierarchy starts from the root.

  .PARAMETER HierarchyScope
    Optional. The depth of the hierarchy to retrieve.
    Defaults to hsSelf (only the specified object).
    See HierarchyScope enumeration for available values (e.g., hsNotebooks, hsSections, hsPages).

  .PARAMETER Force
    Optional. Overwrites read-only output files when present.

  .PARAMETER NoClobber
    Optional. Throws an error if the output file already exists.

  .EXAMPLE
    $notebookId = New-OneNoteNotebook -Path 'C:\Notes\Work'
    $hierarchy = Get-OneNoteHierarchy -Id $notebookId
    $hierarchy

    Retrieves the hierarchy of the specified notebook.

  .EXAMPLE
    $hierarchy = Get-OneNoteHierarchy -Id $notebookId -HierarchyScope hsPages
    $hierarchy.Section.Page | Select-Object -ExpandProperty name

    Retrieves the full page hierarchy under the specified notebook.

  .OUTPUTS
    System.Xml.XmlNode
      The XML node representing the requested hierarchy element.

  .NOTES
    This function uses the OneNote COM API (GetHierarchy) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding()]
  [OutputType([System.Xml.XmlNode])]
  param (
    [string]
    $Id,
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#hierarchyscope
    [Microsoft.Office.Interop.OneNote.HierarchyScope]
    $HierarchyScope = [HierarchyScope]::hsSelf
  )
  $app = Get-OneNoteApplication
  try {
    [xml]$xml = $null
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#gethierarchy-method
    $app.GetHierarchy(
      $Id                       # BSTR            bstrStartNodeID
      , $HierarchyScope         # HierarchyScope  hsScope
      , [ref]$xml               # BSTR*           pbstrHierarchyXmlOut
      , [XMLSchema]::xsCurrent  # XMLSchema       xsSchema
    )
    return $xml
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Get-OneNotePageContent {
  <#
  .SYNOPSIS
    Retrieves the content of a OneNote page as XML nodes.

  .DESCRIPTION
    Retrieves the XML content of the specified OneNote page using the OneNote COM API.
    Returns the parsed XML document element, which includes the page structure,
    text content, and optionally binary object references depending on PageInfo.

  .PARAMETER Id
    Optional. The OneNote object ID of the page to retrieve.

  .PARAMETER PageInfo
    Optional. Specifies what additional information to include in the page content.
    Defaults to piBasic (basic content without binary data).
    Use piAll to include binary object callback IDs required for Export-OneNoteBinaryObject.

  .EXAMPLE
    $pageId = New-OneNotePage -SectionId $sectionId
    $content = Get-OneNotePageContent -Id $pageId
    $content.DocumentElement.name

    Retrieves the XML content of the specified page.

  .EXAMPLE
    $content = Get-OneNotePageContent -Id $pageId -PageInfo piAll
    $content.SelectNodes('//one:CallbackID', $ns)

    Retrieves the page content including binary object callback IDs.

  .OUTPUTS
    System.Xml.XmlNode
      The XML node representing the page content.

  .NOTES
    This function uses the OneNote COM API (GetPageContent) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding()]
  [OutputType([System.Xml.XmlNode])]
  param (
    [string]
    $Id,
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#pageinfo-updated-for-onenote-2013
    [Microsoft.Office.Interop.OneNote.PageInfo]
    $PageInfo = [PageInfo]::piBasic
  )
  $app = Get-OneNoteApplication
  try {
    [xml]$xml = $null
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#page-content-methods
    $app.GetPageContent(
      $Id                       # BSTR      bstrPageID
      , [ref]$xml               # BSTR*     pbstrPageXmlOut
      , $PageInfo               # PageInfo  pageInfoToExport
      , [XMLSchema]::xsCurrent  # XMLSchema xsSchema
    )
    return $xml
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Export-OneNoteHierarchy {
  <#
  .SYNOPSIS
    Exports a OneNote hierarchy as an XML file.

  .DESCRIPTION
    Retrieves the hierarchy of the specified OneNote object and saves it as an XML file
    in the specified destination directory.
    The output file is named `<Id>.hierarchy.xml` by default,
    or `<name>.hierarchy.xml` when -UseName is specified.
    File timestamps (CreationTime, LastWriteTime) are set from the hierarchy XML attributes when available.

  .PARAMETER Id
    Required. The OneNote object ID of the element whose hierarchy will be exported.

  .PARAMETER Destination
    Required. The directory path where the hierarchy XML file will be saved.
    The directory will be created if it does not exist.

  .PARAMETER HierarchyScope
    Optional. The depth of the hierarchy to export.
    Defaults to hsSelf (only the specified object).

  .PARAMETER UseName
    Optional. When specified, the output file is named using the OneNote object's display name
    instead of its object ID.

  .EXAMPLE
    $notebookId = New-OneNoteNotebook -Path 'C:\Notes\Work'
    Export-OneNoteHierarchy -Id $notebookId -Destination 'C:\Export'

    Exports the hierarchy of the specified notebook to C:\Export\<notebookId>.hierarchy.xml.

  .EXAMPLE
    Export-OneNoteHierarchy -Id $notebookId -Destination 'C:\Export' -HierarchyScope hsPages -UseName

    Exports the full page hierarchy with the notebook's display name as the file name.

  .OUTPUTS
    System.IO.FileInfo
      The exported hierarchy XML file.

  .NOTES
    This function uses Get-OneNoteHierarchy internally.
    File timestamps are preserved from the XML attributes when available.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Parameter(Mandatory)]
    [string]
    $Id,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#hierarchyscope
    [Microsoft.Office.Interop.OneNote.HierarchyScope]
    $HierarchyScope = [HierarchyScope]::hsSelf,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [switch]
    $UseName
  )
  $xml = Get-OneNoteHierarchy -Id $Id -HierarchyScope $HierarchyScope
  if ($null -eq $xml) {
    return
  }
  $documentElement = $xml.DocumentElement
  $name = if ($UseName) {
    ConvertTo-SafeFileName -InputString $documentElement.Attributes['name'].Value
  } else {
    $Id
  }
  $resolved = [Path]::GetFullPath($Destination) | Join-Path -ChildPath "$name.hierarchy.xml"
  if ((Test-Path -LiteralPath $Destination -PathType Leaf) -or (Test-Path -LiteralPath $resolved -PathType Container)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $Destination))
  }
  if ((Test-Path -LiteralPath $resolved) -and $NoClobber) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
  }
  if (-not $PSCmdlet.ShouldProcess($resolved, 'Export OneNoteHierarchy')) {
    return
  }
  if (-not (Test-Path -LiteralPath $Destination)) {
    New-Item -Path $Destination -ItemType Directory -Force:$Force -WhatIf:$WhatIfPreference -Confirm:$false | Out-Null
  }
  $isReadOnly = (Test-Path -LiteralPath $resolved -PathType Leaf) -and (Get-Item -LiteralPath $resolved -Force).IsReadOnly
  if ($isReadOnly -and $Force) {
    (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $false
  }
  $xml.Save($resolved)
  if ($documentElement.HasAttributes) {
    $creationTime = $documentElement.Attributes['dateTime']
    if ($creationTime) {
      Set-ItemProperty -LiteralPath $resolved -Name 'CreationTime' -Value ([datetime]::Parse($creationTime.Value)) -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $lastModifiedTime = $documentElement.Attributes['lastModifiedTime']
    if ($lastModifiedTime) {
      Set-ItemProperty -LiteralPath $resolved -Name 'LastWriteTime' -Value ([datetime]::Parse($lastModifiedTime.Value)) -WhatIf:$WhatIfPreference -Confirm:$false
    }
  }
  if ($isReadOnly -and $Force) {
    (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $true
  }
  return Get-Item -LiteralPath $resolved -Force
}
function Export-OneNotePageContent {
  <#
  .SYNOPSIS
    Exports a OneNote page's content as an XML file.

  .DESCRIPTION
    Retrieves the XML content of the specified OneNote page and saves it as an XML file
    in the specified destination directory.
    The output file is named `<Id>.content.xml` by default,
    or `<name>.content.xml` when -UseName is specified.
    File timestamps (CreationTime, LastWriteTime) are set from the XML attributes when available.

  .PARAMETER Id
    Required. The OneNote object ID of the page whose content will be exported.

  .PARAMETER Destination
    Required. The directory path where the content XML file will be saved.
    The directory will be created if it does not exist.

  .PARAMETER PageInfo
    Optional. Specifies what additional information to include in the exported content.
    Defaults to piBasic.

  .PARAMETER Force
    Optional. Overwrites read-only output files when present.

  .PARAMETER NoClobber
    Optional. Throws an error if the output file already exists.

  .PARAMETER UseName
    Optional. When specified, the output file is named using the page's display name
    instead of its object ID.

  .EXAMPLE
    $pageId = New-OneNotePage -SectionId $sectionId
    Export-OneNotePageContent -Id $pageId -Destination 'C:\Export'

    Exports the page content to C:\Export\<pageId>.content.xml.

  .EXAMPLE
    Export-OneNotePageContent -Id $pageId -Destination 'C:\Export' -PageInfo piAll -UseName

    Exports full page content (including binary references) using the page's display name.

  .OUTPUTS
    System.IO.FileInfo
      The exported content XML file.

  .NOTES
    This function uses Get-OneNotePageContent internally.
    File timestamps are preserved from the XML attributes when available.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Parameter(Mandatory)]
    [string]
    $Id,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#pageinfo-updated-for-onenote-2013
    [Microsoft.Office.Interop.OneNote.PageInfo]
    $PageInfo = [PageInfo]::piBasic,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [switch]
    $UseName
  )
  $xml = Get-OneNotePageContent -Id $Id -PageInfo $PageInfo
  if ($null -eq $xml) {
    return
  }
  $documentElement = $xml.DocumentElement
  $name = if ($UseName) {
    ConvertTo-SafeFileName -InputString $documentElement.Attributes['name'].Value
  } else {
    $Id
  }
  $resolved = [Path]::GetFullPath($Destination) | Join-Path -ChildPath "$name.content.xml"
  if ((Test-Path -LiteralPath $Destination -PathType Leaf) -or (Test-Path -LiteralPath $resolved -PathType Container)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $Destination))
  }
  if ((Test-Path -LiteralPath $resolved) -and $NoClobber) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
  }
  if (-not $PSCmdlet.ShouldProcess($resolved, 'Export OneNotePageContent')) {
    return
  }
  if (-not (Test-Path -LiteralPath $Destination)) {
    New-Item -Path $Destination -ItemType Directory -Force:$Force -WhatIf:$WhatIfPreference -Confirm:$false | Out-Null
  }
  $isReadOnly = (Test-Path -LiteralPath $resolved -PathType Leaf) -and (Get-Item -LiteralPath $resolved -Force).IsReadOnly
  if ($isReadOnly -and $Force) {
    (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $false
  }
  $xml.Save($resolved)
  $documentElement = $xml.DocumentElement
  if ($documentElement.HasAttributes) {
    $creationTime = $documentElement.Attributes['dateTime']
    if ($creationTime) {
      Set-ItemProperty -LiteralPath $resolved -Name 'CreationTime' -Value ([datetime]::Parse($creationTime.Value)) -WhatIf:$WhatIfPreference -Confirm:$false
    }
    $lastModifiedTime = $documentElement.Attributes['lastModifiedTime']
    if ($lastModifiedTime) {
      Set-ItemProperty -LiteralPath $resolved -Name 'LastWriteTime' -Value ([datetime]::Parse($lastModifiedTime.Value)) -WhatIf:$WhatIfPreference -Confirm:$false
    }
  }
  if ($isReadOnly -and $Force) {
    (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $true
  }
  return Get-Item -LiteralPath $resolved -Force
}
function Export-OneNotePageAsDocument {
  <#
  .SYNOPSIS
    Exports a OneNote page as a file.

  .DESCRIPTION
    Publishes the specified OneNote page to a file in the specified destination directory
    using the OneNote COM API Publish method.
    The output file is named `<Id>.<ext>` by default,
    or `<name>.<ext>` when -UseName is specified.
    The file extension is determined by the -PublishFormat parameter.
    Supports -WhatIf and -Confirm.

  .PARAMETER Id
    Required. The OneNote object ID of the page to export.

  .PARAMETER Destination
    Required. The directory path where the document will be saved.
    The directory will be created if it does not exist.

  .PARAMETER PublishFormat
    Optional. The format to use when publishing the page.
    Defaults to pfOneNote (.one).
    Other supported formats: pfWord (.docx), pfPDF (.pdf), pfXPS (.xps), pfMHTML (.mht),
    pfHTML (.html), pfEMF (.emf), pfOneNotePackage (.onepkg), pfOneNote2007 (.one).

  .PARAMETER UseName
    Optional. When specified, the output file is named using the page's display name
    instead of its object ID.

  .PARAMETER Force
    Optional. Overwrites read-only output files when present.

  .PARAMETER NoClobber
    Optional. Throws an error if the output file already exists.

  .EXAMPLE
    $pageId = New-OneNotePage -SectionId $sectionId
    Export-OneNotePageAsDocument -Id $pageId -Destination 'C:\Export'

    Exports the page as a .one file to C:\Export\<pageId>.one.

  .EXAMPLE
    Export-OneNotePageAsDocument -Id $pageId -Destination 'C:\Export' -PublishFormat pfWord -UseName

    Exports the page as a Word document (.docx) using the page's display name.

  .OUTPUTS
    System.IO.FileInfo
      The exported document file.

  .NOTES
    This function uses the OneNote COM API (Publish) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Parameter(Mandatory)]
    [string]
    $Id,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#publishformat
    [Microsoft.Office.Interop.OneNote.PublishFormat]
    $PublishFormat = [PublishFormat]::pfOneNote,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [switch]
    $UseName
  )
  $xml = Get-OneNoteHierarchy -Id $Id
  if ($null -eq $xml) {
    return
  }
  $documentElement = $xml.DocumentElement
  $name = if ($UseName) {
    ConvertTo-SafeFileName -InputString $documentElement.Attributes['name'].Value
  } else {
    $Id
  }
  $ext = Get-ExtensionByPublishFormat -PublishFormat $PublishFormat
  $resolved = [Path]::GetFullPath($Destination) | Join-Path -ChildPath "$name$ext"
  if ((Test-Path -LiteralPath $Destination -PathType Leaf) -or (Test-Path -LiteralPath $resolved -PathType Container)) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $Destination))
  }
  if ((Test-Path -LiteralPath $resolved) -and $NoClobber) {
    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $resolved))
  }
  if (-not $PSCmdlet.ShouldProcess($resolved, 'Export OneNotePageAsDocument')) {
    return
  }
  if (-not (Test-Path -LiteralPath $Destination)) {
    New-Item -Path $Destination -ItemType Directory -Force:$Force -WhatIf:$WhatIfPreference -Confirm:$false | Out-Null
  }
  $isReadOnly = (Test-Path -LiteralPath $resolved -PathType Leaf) -and (Get-Item -LiteralPath $resolved -Force).IsReadOnly
  if ($isReadOnly -and $Force) {
    (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $false
  }
  $app = Get-OneNoteApplication
  try {
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#publish-method
    $app.Publish(
      $Id               # BSTR          bstrHierarchyID
      , $resolved       # BSTR          bstrTargetFilePath
      , $PublishFormat  # PublishFormat pfPublishFormat
      , $null           # BSTR          bstrCLSIDofExporter
    )
    if ($isReadOnly -and $Force) {
      (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $true
    }
    return Get-Item -LiteralPath $resolved -Force
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Export-OneNoteBinaryObject {
  <#
  .SYNOPSIS
    Exports binary objects embedded in a OneNote page.

  .DESCRIPTION
    Retrieves all binary objects (images, attachments) from the specified OneNote page
    and saves each one as a file in a subdirectory under the destination directory.
    The subdirectory is named `<Id>` by default, or `<name>` when -UseName is specified.
    Each file is named by its callback ID, with the extension determined by the file signature.
    File timestamps (CreationTime, LastWriteTime) are set from the parent OE element when available.
    If the page contains no binary objects, no files are created and nothing is returned.

  .PARAMETER Id
    Required. The OneNote object ID of the page from which binary objects will be exported.

  .PARAMETER Destination
    Required. The parent directory path where the binary object subdirectory will be created.
    The directory will be created if it does not exist.

  .PARAMETER Force
    Optional. Overwrites read-only output files when present.

  .PARAMETER NoClobber
    Optional. Throws an error if an output file already exists.

  .PARAMETER UseName
    Optional. When specified, the subdirectory is named using the page's display name
    instead of its object ID.

  .EXAMPLE
    $pageId = New-OneNotePage -SectionId $sectionId
    Export-OneNoteBinaryObject -Id $pageId -Destination 'C:\Export'

    Exports all binary objects from the page to C:\Export\<pageId>\.

  .EXAMPLE
    Export-OneNoteBinaryObject -Id $pageId -Destination 'C:\Export' -UseName

    Exports binary objects to a subdirectory named after the page's display name.

  .OUTPUTS
    System.IO.FileInfo
      One FileInfo object per exported binary object file.
      Nothing is returned if the page contains no binary objects.

  .NOTES
    This function uses the OneNote COM API (GetBinaryPageContent) internally.
    Supported file signatures: PNG, JPEG, GIF, BMP, TIFF.
    Files without a recognized signature are saved with a .bin extension.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([System.IO.FileInfo])]
  param (
    [Parameter(Mandatory)]
    [string]
    $Id,
    [Parameter(Mandatory)]
    [ValidateScript({ (Test-Path -LiteralPath $_ -IsValid) -and (Test-Path -LiteralPath $_ -PathType Container) })]
    [string]
    $Destination,
    [switch]
    $Force,
    [switch]
    $NoClobber,
    [switch]
    $UseName
  )
  $app = Get-OneNoteApplication
  try {
    $xml = Get-OneNotePageContent -Id $Id
    if ($null -eq $xml) {
      return
    }
    $documentElement = $xml.DocumentElement
    $name = if ($UseName) {
      ConvertTo-SafeFileName -InputString $documentElement.Attributes['name'].Value
    } else {
      $Id
    }
    $ns = [XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('one', $documentElement.NamespaceURI)
    $nodes = $xml.SelectNodes('//one:CallbackID[@callbackID]', $ns)
    if ((Test-Path -LiteralPath $Destination -PathType Leaf)) {
      $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $Destination))
    }
    return $nodes |
    ForEach-Object {
      $binary = $null
      $callbackId = $_.Attributes['callbackID'].Value
      # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#getbinarypagecontent-method
      $app.GetBinaryPageContent(
        $Id,          # BSTR  bstrPageID
        $callbackId,  # BSTR  bstrCallbackID
        [ref]$binary  # BSTR* pbstrBinaryObjectB64Out
      )
      $directory = $Destination | Join-Path -ChildPath $name
      if (-not [string]::IsNullOrEmpty($binary)) {
        $bytes = [Convert]::FromBase64String($binary)
        $ext = Get-ExtensionBySignature -Bytes $bytes
        $file = $directory | Join-Path -ChildPath "$callbackId$ext"
        if ((Test-Path -LiteralPath $directory -PathType Leaf)) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $directory))
        }
        if ((Test-Path -LiteralPath $file -PathType Container) -or ((Test-Path -LiteralPath $file -PathType Leaf) -and $NoClobber)) {
          $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -ErrorId 'ItemAlreadyExists' -TargetObject $file))
        }
        if (-not $PSCmdlet.ShouldProcess($file, 'Export OneNoteBinaryObject')) {
          return
        }
        if (-not (Test-Path -LiteralPath $Destination)) {
          New-Item -Path $Destination -ItemType Directory -Force:$Force -WhatIf:$WhatIfPreference -Confirm:$false | Out-Null
        }
        if (-not (Test-Path -LiteralPath $directory)) {
          New-Item -Path $directory -ItemType Directory -Force:$Force -WhatIf:$WhatIfPreference -Confirm:$false | Out-Null
        }
        $isReadOnly = (Test-Path -LiteralPath $resolved -PathType Leaf) -and (Get-Item -LiteralPath $resolved -Force).IsReadOnly
        if ($isReadOnly -and $Force) {
          (Get-Item -LiteralPath $resolved -Force).IsReadOnly = $false
        }
        [File]::WriteAllBytes($file, $bytes)
        $oe = $_.SelectSingleNode('ancestor::one:OE[1]', $ns)
        if ($oe -and $oe.HasAttributes) {
          $creationTime = $oe.Attributes['creationTime']
          if ($creationTime) {
            Set-ItemProperty -LiteralPath $file -Name 'CreationTime' -Value ([datetime]::Parse($creationTime.Value)) -WhatIf:$WhatIfPreference -Confirm:$false
          }
          $lastModifiedTime = $oe.Attributes['lastModifiedTime']
          if ($lastModifiedTime) {
            Set-ItemProperty -LiteralPath $file -Name 'LastWriteTime' -Value ([datetime]::Parse($lastModifiedTime.Value)) -WhatIf:$WhatIfPreference -Confirm:$false
          }
        }
        if ($isReadOnly -and $Force) {
          (Get-Item -LiteralPath $file -Force).IsReadOnly = $true
        }
        return Get-Item -LiteralPath $file -Force
      }
    }
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Import-OneNoteHierarchy {
  <#
  .SYNOPSIS
    Imports a OneNote hierarchy XML file into OneNote.

  .DESCRIPTION
    Reads a hierarchy XML file previously exported by Export-OneNoteHierarchy
    and applies it to the OneNote application using the UpdateHierarchy COM API method.
    Accepts files by path (with wildcard support) or by literal path.
    Supports -WhatIf and -Confirm to preview or confirm the operation.

  .PARAMETER Path
    Required in PathSet. One or more file paths to the hierarchy XML files to import.
    Accepts wildcards. Accepts pipeline input by value and property name.

  .PARAMETER LiteralPath
    Required in LiteralPathSet. One or more literal file paths to the hierarchy XML files to import.
    No wildcard expansion is performed. Accepts pipeline input by property name.
    Alias: PSPath, LP.

  .EXAMPLE
    $notebookId = New-OneNoteNotebook -Path 'C:\Notes\Work'
    $xmlFile = Export-OneNoteHierarchy -Id $notebookId -Destination 'C:\Export'
    Import-OneNoteHierarchy -LiteralPath $xmlFile.FullName -Confirm:$false

    Exports the hierarchy of a notebook and re-imports it from the XML file.

  .EXAMPLE
    Get-ChildItem -Path 'C:\Export\*.hierarchy.xml' | Import-OneNoteHierarchy -Confirm:$false

    Imports all hierarchy XML files in the specified directory via pipeline.

  .OUTPUTS
    None.

  .NOTES
    This function uses the OneNote COM API (UpdateHierarchy) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath
  )
  begin {
    $app = Get-OneNoteApplication
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
        if (-not $PSCmdlet.ShouldProcess($_, 'Import Hierarchy')) {
          return
        }
        $content = $_ | Get-Content -Force -Raw
        # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#updatehierarchy-method
        $app.UpdateHierarchy(
          $content                  # BSTR      bstrChangesXmlIn
          , [XMLSchema]::xsCurrent  # XMLSchema xsSchema
        )
      }
    } catch {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
  end {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function Import-OneNotePageContent {
  <#
  .SYNOPSIS
    Imports a OneNote page content XML file into OneNote.

  .DESCRIPTION
    Reads a page content XML file previously exported by Export-OneNotePageContent
    and applies it to the OneNote application using the UpdatePageContent COM API method.
    Accepts files by path (with wildcard support) or by literal path.
    Supports -WhatIf and -Confirm to preview or confirm the operation.

  .PARAMETER Path
    Required in PathSet. One or more file paths to the page content XML files to import.
    Accepts wildcards. Accepts pipeline input by value and property name.

  .PARAMETER LiteralPath
    Required in LiteralPathSet. One or more literal file paths to the page content XML files to import.
    No wildcard expansion is performed. Accepts pipeline input by property name.
    Alias: PSPath, LP.

  .EXAMPLE
    $pageId = New-OneNotePage -SectionId $sectionId
    $xmlFile = Export-OneNotePageContent -Id $pageId -Destination 'C:\Export'
    Import-OneNotePageContent -LiteralPath $xmlFile.FullName -Confirm:$false

    Exports the content of a page and re-imports it from the XML file.

  .EXAMPLE
    Get-ChildItem -Path 'C:\Export\*.content.xml' | Import-OneNotePageContent -Confirm:$false

    Imports all page content XML files in the specified directory via pipeline.

  .OUTPUTS
    None.

  .NOTES
    This function uses the OneNote COM API (UpdatePageContent) internally.
    COM objects are released after execution via garbage collection.
  #>
  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([void])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]]
    $LiteralPath
  )
  begin {
    $app = Get-OneNoteApplication
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
        if (-not $PSCmdlet.ShouldProcess($_, 'Import PageContent')) {
          return
        }
        $content = $_ | Get-Content -Force -Raw
        # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#updatepagecontent-method
        $app.UpdatePageContent(
          $content                  # BSTR          bstrPageChangesXmlIn
          , 0                       # DATE          dateExpectedLastModified
          , [XMLSchema]::xsCurrent  # XMLSchema     xsSchema
          , $false                  # VARIANT_BOOL  force
        )
      }
    } catch {
      Get-Variable |
      Where-Object -Property Value -Is [__ComObject] |
      Clear-Variable -Force -WhatIf:$false -Confirm:$false
      [GC]::Collect()
      [GC]::WaitForPendingFinalizers()
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
  end {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
#endregion
#region Private
function Get-OneNoteApplication {
  [CmdletBinding()]
  [OutputType([__ComObject])]
  param ()
  return New-Object -ComObject OneNote.Application
}
function Close-OneNoteNotebook {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory)]
    [string]
    $NotebookId
  )
  $app = Get-OneNoteApplication
  try {
    # https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote#closenotebook-method
    $app.CloseNotebook(
      $NotebookId # BSTR          bstrObjectID
      , $false    # VARIANT_BOOL  force
    )
  } finally {
    Get-Variable |
    Where-Object -Property Value -Is [__ComObject] |
    Clear-Variable -Force -WhatIf:$false -Confirm:$false
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
function ConvertTo-SafeFileName {
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $InputString
  )
  return -join (
    $InputString.ToCharArray() |
    Where-Object { $_ -notin [Path]::GetInvalidFileNameChars() }
  )
}
function Get-ExtensionByPublishFormat {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory)]
    [Microsoft.Office.Interop.OneNote.PublishFormat]
    $PublishFormat
  )
  # https://learn.microsoft.com/en-us/office/client-developer/onenote/enumerations-onenote-developer-reference#odc_PublishFormat
  switch ($PublishFormat) {
    { $_ -in [PublishFormat]::pfOneNote } {
      return '.one'
    }
    { $_ -in [PublishFormat]::pfOneNotePackage } {
      return '.onepkg'
    }
    { $_ -in [PublishFormat]::pfMHTML } {
      return '.mht'
    }
    { $_ -in [PublishFormat]::pfPDF } {
      return '.pdf'
    }
    { $_ -in [PublishFormat]::pfXPS } {
      return '.xps'
    }
    { $_ -in [PublishFormat]::pfWord } {
      return '.docx'
    }
    { $_ -in [PublishFormat]::pfEMF } {
      return '.emf'
    }
    { $_ -in [PublishFormat]::pfHTML } {
      return '.html'
    }
    { $_ -in [PublishFormat]::pfOneNote2007 } {
      return '.one'
    }
  }
}
function Get-ExtensionBySignature {
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory)]
    [byte[]]
    $Bytes
  )
  if (
    $Bytes.Count -ge 8 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..7] -DifferenceObject @(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))
  ) {
    return '.png'
  }
  if (
    $Bytes.Count -ge 3 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..2] -DifferenceObject @(0xFF, 0xD8, 0xFF))
  ) {
    return '.jpg'
  }
  if (
    $Bytes.Count -ge 4 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..3] -DifferenceObject @(0x47, 0x49, 0x46, 0x38))
  ) {
    return '.gif'
  }
  if (
    $Bytes.Count -ge 2 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..1] -DifferenceObject @(0x42, 0x4D))
  ) {
    return '.bmp'
  }
  if (
    $Bytes.Count -ge 4 -and (
      -not (Compare-Object -ReferenceObject $Bytes[0..3] -DifferenceObject @(0x49, 0x49, 0x2A, 0x00)) -or
      -not (Compare-Object -ReferenceObject $Bytes[0..3] -DifferenceObject @(0x4D, 0x4D, 0x00, 0x2A))
    )
  ) {
    return '.tiff'
  }
  if (
    $Bytes.Count -ge 8 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..7] -DifferenceObject @(0x01, 0x00, 0x00, 0x00, 0x20, 0x45, 0x4D, 0x46))
  ) {
    return '.emf'
  }
  if (
    $Bytes.Count -ge 5 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..4] -DifferenceObject @(0x25, 0x50, 0x44, 0x46, 0x2D))
  ) {
    return '.pdf'
  }
  # ZIP archive (e.g., docx, xlsx, pptx)
  if (
    $Bytes.Count -ge 4 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..3] -DifferenceObject @(0x50, 0x4B, 0x03, 0x04))
  ) {
    return '.docx'
  }
  # Microsoft Office (e.g., doc, xls, ppt)
  if (
    $Bytes.Count -ge 4 -and
    -not (Compare-Object -ReferenceObject $Bytes[0..3] -DifferenceObject @(0xD0, 0xCF, 0x11, 0xE0))
  ) {
    return '.doc'
  }
  return [string]::Empty
}
#endregion

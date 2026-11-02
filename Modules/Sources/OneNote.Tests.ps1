using module .\..\Automation.Office.psd1
using module .\Path.psm1
using namespace Microsoft.Office.Interop.OneNote
using namespace System.Collections.Generic
using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO

[CmdletBinding()]
[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()
Set-StrictMode -Version Latest

InModuleScope 'Automation.Office' {
  BeforeAll {
    function Get-NotebookPath {
      [CmdletBinding()]
      param ()
      return $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
    }
    function Get-SectionPath {
      [CmdletBinding()]
      param (
        [string]
        $NotebookPath
      )
      return $NotebookPath | Join-Path -ChildPath "Section.$([Guid]::NewGuid().ToString('N')).one"
    }
    function Get-SectionGroupPath {
      [CmdletBinding()]
      param (
        [string]
        $NotebookPath
      )
      return $NotebookPath | Join-Path -ChildPath "$([Guid]::NewGuid().ToString('N'))"

    }
  }
  Describe 'New-OneNoteNotebook' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'creates a new notebook' {
        $notebookId = New-OneNoteNotebook -Path $notebookPath

        $notebookId | Should-NotBeNull
        Test-Path -LiteralPath $notebookPath | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-OneNoteNotebook' -Tag 'CI', 'Mock' {
    BeforeAll {
      Mock -CommandName Get-OneNoteApplication -MockWith { throw }
    }
    BeforeEach {
      $notebookPath = Get-NotebookPath
    }
    Context 'SupportsShouldProcess' {
      It 'does not call the program when WhatIf is specified' {
        New-OneNoteNotebook -Path $notebookPath -WhatIf

        Should-NotInvoke -CommandName Get-OneNoteApplication
      }
    }
  }
  Describe 'New-OneNoteSection' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'creates a new section by Path' {
        $sectionId = New-OneNoteSection -Path $sectionPath

        $sectionId | Should-NotBeNull
        Wait-Path -LiteralPath $sectionPath | Should-BeTrue
      }
      It 'creates a new section by ParentId' {
        $sectionId = New-OneNoteSection -Path $sectionPath -ParentId $notebookId

        $sectionId | Should-NotBeNull
        Wait-Path -LiteralPath $sectionPath | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-OneNoteSection' -Tag 'CI', 'Mock' {
    BeforeAll {
      Mock -CommandName Get-OneNoteApplication -MockWith { throw }
    }
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
    }
    Context 'SupportsShouldProcess' {
      It 'does not call the program when WhatIf is specified' {
        New-OneNoteSection -Path $sectionPath -WhatIf

        Should-NotInvoke -CommandName Get-OneNoteApplication
      }
    }
  }
  Describe 'New-OneNoteSectionGroup' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionGroupPath = Get-SectionGroupPath -NotebookPath $notebookPath
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'creates a new section group by Path' {
        $sectionGroupId = New-OneNoteSectionGroup -Path $sectionGroupPath

        $sectionGroupId | Should-NotBeNull
        Wait-Path -LiteralPath $sectionGroupPath | Should-BeTrue
      }
      It 'creates a new section group by ParentId' {
        $sectionGroupId = New-OneNoteSectionGroup -Path $sectionGroupPath -ParentId $notebookId

        $sectionGroupId | Should-NotBeNull
        Wait-Path -LiteralPath $sectionGroupPath | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-OneNoteSectionGroup' {
    BeforeAll {
      Mock -CommandName Get-OneNoteApplication -MockWith { throw }
    }
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $sectionGroupPath = Get-SectionGroupPath -NotebookPath $notebookPath
    }
    Context 'SupportsShouldProcess' {
      It 'does not call the program when WhatIf is specified' {
        New-OneNoteSectionGroup -Path $sectionGroupPath -WhatIf

        Should-NotInvoke -CommandName Get-OneNoteApplication
      }
    }
  }
  Describe 'New-OneNotePage' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'creates a new page' {
        $pageId = New-OneNotePage -SectionId $sectionId

        $pageId | Should-NotBeNull
      }
    }
    Context 'SupportsShouldProcess' {
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'New-OneNotePage' -Tag 'CI', 'Mock' {
    BeforeAll {
      Mock -CommandName Get-OneNoteApplication -MockWith { throw }
    }
    Context 'SupportsShouldProcess' {
      It 'does not call the program when WhatIf is specified' {
        New-OneNotePage -SectionId 'sectionId' -WhatIf

        Should-NotInvoke -CommandName Get-OneNoteApplication
      }
    }
  }
  Describe 'Get-OneNoteHierarchy' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'retrieves the hierarchy of a notebook' {
        $hierarchy = Get-OneNoteHierarchy -Id $notebookId

        $hierarchy | ForEach-Object { $_ -is [System.Xml.XmlNode] | Should-BeTrue }
      }
    }
    Context 'Other parameters' {
      It 'retrieves the hierarchy of a notebook with HierarchyScope' {
        $hierarchy = Get-OneNoteHierarchy -Id $notebookId -HierarchyScope hsPages

        $hierarchy | ForEach-Object { $_ -is [System.Xml.XmlNode] | Should-BeTrue }
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Get-OneNoteHierarchy' -Tag 'CI', 'Mock' {
    BeforeAll {
      if (-not ('XMLSchema' -as [type])) {
        Add-Type -TypeDefinition @'
public enum XMLSchema
{
  xsCurrent = 0
}
'@
      }
      function New-OneNoteHierarchyReaderMockApplication {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function only constructs an in-memory mock application and does not change system state')]
        param ()
        $app = [PSCustomObject]@{ Calls = [List[object]]::new() }
        $app | Add-Member -MemberType ScriptMethod -Name GetHierarchy -Value {
          param (
            [string]
            $Id,
            [object]
            $HierarchyScope,
            [ref]
            $Xml,
            [object]
            $Schema
          )
          [void]$this.Calls.Add([PSCustomObject]@{
              Id             = $Id
              HierarchyScope = $HierarchyScope
              Schema         = $Schema
            })
          $Xml.Value = '<one:Notebook xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Mock Notebook" />'
        }
        return $app
      }
    }
    Context 'ParameterSetName' {
      It 'returns hierarchy XML with default HierarchyScope' {
        $app = New-OneNoteHierarchyReaderMockApplication
        Mock -CommandName Get-OneNoteApplication -MockWith { $app }

        $result = Get-OneNoteHierarchy -Id 'notebook-id'
        $result | Should-NotBeNull
        $result.DocumentElement.LocalName | Should-Be 'Notebook'
        $result.DocumentElement.Attributes['name'].Value | Should-Be 'Mock Notebook'
        $app.Calls.Count | Should-Be 1
        $app.Calls[0].Id | Should-Be 'notebook-id'
        $app.Calls[0].HierarchyScope.ToString() | Should-Be 'hsSelf'
      }
    }
    Context 'Other parameters' {
      It 'passes HierarchyScope to underlying application call' {
        $app = New-OneNoteHierarchyReaderMockApplication
        Mock -CommandName Get-OneNoteApplication -MockWith { $app }

        $result = Get-OneNoteHierarchy -Id 'notebook-id' -HierarchyScope hsPages
        $result | Should-NotBeNull
        $app.Calls.Count | Should-Be 1
        $app.Calls[0].HierarchyScope.ToString() | Should-Be 'hsPages'
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Get-OneNotePageContent' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'retrieves the content of a page as XML' {
        $pageContent = Get-OneNotePageContent -Id $pageId

        $pageContent | ForEach-Object { $_ -is [System.Xml.XmlNode] | Should-BeTrue }
      }
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Get-OneNotePageContent' -Tag 'CI', 'Mock' {
    BeforeAll {
      if (-not ('XMLSchema' -as [type])) {
        Add-Type -TypeDefinition @'
public enum XMLSchema
{
  xsCurrent = 0
}
'@
      }
      function New-OneNotePageContentReaderMockApplication {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function only constructs an in-memory mock application and does not change system state')]
        param ()
        $app = [PSCustomObject]@{ Calls = [List[object]]::new() }
        $app | Add-Member -MemberType ScriptMethod -Name GetPageContent -Value {
          param (
            [string]
            $Id,
            [ref]
            $Xml,
            [object]
            $PageInfo,
            [object]
            $Schema
          )
          [void]$this.Calls.Add([PSCustomObject]@{
              Id       = $Id
              PageInfo = $PageInfo
              Schema   = $Schema
            })
          $Xml.Value = '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Mock Page" />'
        }
        return $app
      }
    }
    BeforeEach {
      $app = New-OneNotePageContentReaderMockApplication
      Mock -CommandName Get-OneNoteApplication -MockWith { $app }
    }
    Context 'ParameterSetName' {
      It 'returns page content XML with default PageInfo' {
        $result = Get-OneNotePageContent -Id 'page-id'

        $result | Should-NotBeNull
        $result.DocumentElement.LocalName | Should-Be 'Page'
        $result.DocumentElement.Attributes['name'].Value | Should-Be 'Mock Page'
        $app.Calls.Count | Should-Be 1
        $app.Calls[0].Id | Should-Be 'page-id'
        $app.Calls[0].PageInfo | Should-Be ([Microsoft.Office.Interop.OneNote.PageInfo]::piBasic)
      }
    }
    Context 'Other parameters' {
      It 'passes PageInfo to underlying application call' {
        $result = Get-OneNotePageContent -Id 'page-id' -PageInfo piAll

        $result | Should-NotBeNull
        $app.Calls.Count | Should-Be 1
        $app.Calls[0].PageInfo | Should-Be ([Microsoft.Office.Interop.OneNote.PageInfo]::piAll)
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Export-OneNoteHierarchy' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
      $destination = $notebookPath | Join-Path -ChildPath "$notebookId.hierarchy.xml"
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'exports the hierarchy as XML' {
        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath -Force -WhatIf

        Test-Path -LiteralPath $destination | Should-BeFalse
      }
      It 'overwrites read-only output when Force is specified' {
        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath
        (Get-Item -LiteralPath $destination -Force).IsReadOnly = $true

        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath -Force -Confirm
        (Get-Item -LiteralPath $destination -Force).IsReadOnly | Should-BeTrue
      }
      It 'throws with -NoClobber when output exists' {
        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath

        { Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath -NoClobber } | Should-Throw
        { Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath -Force -NoClobber } | Should-Throw
      }
    }
    Context 'Other parameters' {
      It 'exports the hierarchy as XML with HierarchyScope' {
        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath -HierarchyScope hsPages

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports the hierarchy as XML with UseName' {
        $notebookName = Split-Path -Path $notebookPath -Leaf
        $destination = $notebookPath | Join-Path -ChildPath "$notebookName.hierarchy.xml"
        Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath -UseName

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Export-OneNoteHierarchy' -Tag 'CI', 'Mock' {
    BeforeAll {
      $xml = [xml]@'
<one:Notebook xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Sample:Hierarchy?Name" dateTime="2026-05-23T00:00:00.0000000" lastModifiedTime="2026-05-23T00:00:00.0000000" />
'@
    }
    BeforeEach {
      $destination = $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
      New-Item -Path $destination -ItemType Directory -Force | Out-Null
    }
    AfterEach {
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'returns no output when hierarchy is null' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $null }

        $result = Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination
        $result | Should-BeNull
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'notebook-id.hierarchy.xml'

        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'uses sanitized hierarchy name when UseName is specified' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination -UseName -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'notebook-id.hierarchy.xml'
        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
      It 'accepts HierarchyScope' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination -HierarchyScope hsPages -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'notebook-id.hierarchy.xml'

        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws with NoClobber when the output file already exists' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        $item = $destination | Join-Path -ChildPath 'notebook-id.hierarchy.xml'
        New-Item -Path $item -ItemType File -Force | Out-Null

        { Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination -NoClobber } | Should-Throw
        { Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination -Force -NoClobber } | Should-Throw
      }
      It 'throws when resolved output path already exists as a directory' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        $item = $destination | Join-Path -ChildPath 'notebook-id.hierarchy.xml'
        New-Item -Path $item -ItemType Directory -Force | Out-Null

        { Export-OneNoteHierarchy -Id 'notebook-id' -Destination $destination } | Should-Throw
      }
    }
  }
  Describe 'Export-OneNotePageContent' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
      $destination = $notebookPath | Join-Path -ChildPath "$pageId.content.xml"
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'exports the content of a page as XML' {
        Export-OneNotePageContent -Id $pageId -Destination $notebookPath
        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'overwrites read-only output when Force is specified' {
        Export-OneNotePageContent -Id $pageId -Destination $notebookPath
        (Get-Item -LiteralPath $destination -Force).IsReadOnly = $true

        Export-OneNotePageContent -Id $pageId -Destination $notebookPath -Force
        (Get-Item -LiteralPath $destination -Force).IsReadOnly | Should-BeTrue
      }
      It 'throws with -NoClobber when output exists' {
        Export-OneNotePageContent -Id $pageId -Destination $notebookPath

        { Export-OneNotePageContent -Id $pageId -Destination $notebookPath -NoClobber } | Should-Throw
        { Export-OneNotePageContent -Id $pageId -Destination $notebookPath -Force -NoClobber } | Should-Throw
      }
    }
    Context 'Other parameters' {
      It 'exports the content of a page as XML with PageInfo' {
        Export-OneNotePageContent -Id $pageId -Destination $notebookPath -PageInfo piAll

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
      It 'exports the content of a page as XML with UseName' {
        $destination = $notebookPath | Join-Path -ChildPath 'Untitled page.content.xml'
        Export-OneNotePageContent -Id $pageId -Destination $notebookPath -UseName

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Export-OneNotePageContent' -Tag 'CI', 'Mock' {
    BeforeAll {
      $xml = [xml]@'
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Sample:Content?Name" dateTime="2026-05-23T00:00:00.0000000" lastModifiedTime="2026-05-23T00:00:00.0000000" />
'@
    }
    BeforeEach {
      $destination = $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
      New-Item -Path $destination -ItemType Directory -Force | Out-Null
    }
    AfterEach {
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'returns no output when page content is null' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $null }

        $result = Export-OneNotePageContent -Id 'page-id' -Destination $destination
        $result | Should-BeNull
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        Export-OneNotePageContent -Id 'page-id' -Destination $destination -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'page-id.content.xml'
        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'uses sanitized page name when UseName is specified' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        Export-OneNotePageContent -Id 'page-id' -Destination $destination -UseName -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'page-id.content.xml'
        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
    }
    Context 'Edge cases' {
      It 'throws with NoClobber when the output file already exists' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        $item = $destination | Join-Path -ChildPath 'page-id.content.xml'
        New-Item -Path $item -ItemType File -Force | Out-Null

        { Export-OneNotePageContent -Id 'page-id' -Destination $destination -NoClobber } | Should-Throw
        { Export-OneNotePageContent -Id 'page-id' -Destination $destination -Force -NoClobber } | Should-Throw
      }
      It 'throws when resolved output path already exists as a directory' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        $item = $destination | Join-Path -ChildPath 'page-id.content.xml'
        New-Item -Path $item -ItemType Directory -Force | Out-Null

        { Export-OneNotePageContent -Id 'page-id' -Destination $destination } | Should-Throw
        { Export-OneNotePageContent -Id 'page-id' -Destination $destination -Force } | Should-Throw
      }
    }
  }
  Describe 'Export-OneNotePageAsDocument' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
      $destination = $notebookPath | Join-Path -ChildPath "$pageId.one"
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'exports a page as a OneNote document' {
        Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath -WhatIf

        Test-Path -LiteralPath $destination | Should-BeFalse
      }
      It 'throws when the target is read-only' {
        Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath
        (Get-Item -LiteralPath $destination -Force).IsReadOnly = $true

        { Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath } | Should-Throw
      }
      It 'throws with -NoClobber when output exists' {
        Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath

        { Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath -NoClobber } | Should-Throw
        { Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath -Force -NoClobber } | Should-Throw
      }
    }
    Context 'Other parameters' {
      It 'exports a page as a OneNote document with PublishFormat' {
        Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath -PublishFormat pfWord

        Test-Path -LiteralPath ([Path]::ChangeExtension($destination, '.docx')) | Should-BeTrue
      }
      It 'exports a page as a OneNote document with UseName' {
        $destination = $notebookPath | Join-Path -ChildPath 'Untitled page.one'
        Export-OneNotePageAsDocument -Id $pageId -Destination $notebookPath -UseName

        Test-Path -LiteralPath $destination | Should-BeTrue
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Export-OneNotePageAsDocument' -Tag 'CI', 'Mock' {
    BeforeAll {
      $xml = [xml]@'
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Sample:Page?Name" />
'@
    }
    BeforeEach {
      $destination = $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
      New-Item -Path $destination -ItemType Directory -Force | Out-Null
    }
    AfterEach {
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'page-id.one'
        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'uses sanitized page name with UseName when WhatIf is specified' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination -UseName -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'SamplePageName.one'
        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
      It 'uses docx extension when PublishFormat is pfWord and WhatIf is specified' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination -PublishFormat pfWord -WhatIf
        $outputPath = $destination | Join-Path -ChildPath 'page-id.docx'
        Test-Path -LiteralPath $outputPath | Should-BeFalse
      }
      It 'returns no output when hierarchy is null' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $null }

        $result = Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination
        $result | Should-BeNull
      }
    }
    Context 'Edge cases' {
      It 'throws when resolved output path already exists as a directory' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        $conflictDirectory = $destination | Join-Path -ChildPath 'page-id.one'
        New-Item -Path $conflictDirectory -ItemType Directory -Force | Out-Null

        { Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination } | Should-Throw
        { Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination -Force } | Should-Throw
      }
      It 'throws with NoClobber when the output file already exists' {
        Mock -CommandName Get-OneNoteHierarchy -MockWith { $xml }

        $item = $destination | Join-Path -ChildPath 'page-id.one'
        New-Item -Path $item -ItemType File -Force | Out-Null

        { Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination -NoClobber } | Should-Throw
        { Export-OneNotePageAsDocument -Id 'page-id' -Destination $destination -Force -NoClobber } | Should-Throw
      }
    }
  }
  Describe 'Export-OneNoteBinaryObject' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
      $destination = $notebookPath | Join-Path -ChildPath "$pageId.bin"
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'exports a page as a binary object' {
        Export-OneNoteBinaryObject -Id $pageId -Destination $notebookPath
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Export-OneNoteBinaryObject -Id $pageId -Destination $notebookPath -Force -WhatIf
      }
    }
    Context 'Other parameters' {
      It 'exports a page as a binary object with UseName' {
        Export-OneNoteBinaryObject -Id $pageId -Destination $notebookPath -UseName
      }
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Export-OneNoteBinaryObject' -Tag 'CI', 'Mock' {
    BeforeAll {
      $xml = [xml]@'
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Sample:Page?Name">
  <one:Outline>
    <one:OE creationTime="2026-05-23T00:00:00.0000000" lastModifiedTime="2026-05-23T00:00:00.0000000">
      <one:CallbackID callbackID="callback-01" />
    </one:OE>
  </one:Outline>
</one:Page>
'@
      function New-OneNoteMockApplication {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function only constructs an in-memory mock application and does not change system state')]
        param (
          [string]
          $Base64 = [Convert]::ToBase64String([byte[]]@(0x50, 0x4B, 0x03, 0x04))
        )
        $app = [PSCustomObject]@{ Base64 = $Base64; Calls = [List[object]]::new() }
        $app | Add-Member -MemberType ScriptMethod -Name GetBinaryPageContent -Value {
          param (
            [string]
            $Id,
            [string]
            $CallbackId,
            [ref]
            $Binary
          )
          [void]$this.Calls.Add([PSCustomObject]@{
              Id         = $Id
              CallbackId = $CallbackId
            })
          $Binary.Value = $this.Base64
        }
        return $app
      }
    }
    BeforeEach {
      Mock -CommandName Get-OneNoteApplication -MockWith { New-OneNoteMockApplication }

      $destination = $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
      New-Item -Path $destination -ItemType Directory -Force | Out-Null
    }
    AfterEach {
      if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not create output when WhatIf is specified' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        Export-OneNoteBinaryObject -Id 'page-id' -Destination $destination -WhatIf
        $directory = $destination | Join-Path -ChildPath 'page-id'
        Test-Path -LiteralPath $directory | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'uses sanitized page name as output directory when UseName is specified' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        Export-OneNoteBinaryObject -Id 'page-id' -Destination $destination -UseName -WhatIf
        $directory = $destination | Join-Path -ChildPath 'SamplePageName'

        Test-Path -LiteralPath $directory | Should-BeFalse
      }
      It 'returns no output when page content is null' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $null }

        $result = Export-OneNoteBinaryObject -Id 'page-id' -Destination $destination

        $result | Should-BeNull
      }
    }
    Context 'Edge cases' {
      It 'throws when Destination is an existing file' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        $item = $destination | Join-Path -ChildPath 'destination.txt'
        New-Item -Path $item -ItemType File -Force | Out-Null

        { Export-OneNoteBinaryObject -Id 'page-id' -Destination $item } | Should-Throw
        { Export-OneNoteBinaryObject -Id 'page-id' -Destination $item -Force } | Should-Throw
      }
      It 'throws with NoClobber when the output file already exists' {
        Mock -CommandName Get-OneNotePageContent -MockWith { $xml }

        $directory = $destination | Join-Path -ChildPath 'page-id'
        New-Item -Path $directory -ItemType Directory -Force | Out-Null

        $item = $directory | Join-Path -ChildPath 'callback-01.docx'
        New-Item -Path $item -ItemType File -Force | Out-Null

        { Export-OneNoteBinaryObject -Id 'page-id' -Destination $destination -NoClobber } | Should-Throw
        { Export-OneNoteBinaryObject -Id 'page-id' -Destination $destination -Force -NoClobber } | Should-Throw
      }
      It 'returns no output when page content has no binary objects' {
        Mock -CommandName Get-OneNotePageContent -MockWith {
          [xml]@'
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="Sample:Page?Name">
  <one:Outline>
    <one:OE />
  </one:Outline>
</one:Page>
'@
        }

        $result = Export-OneNoteBinaryObject -Id 'page-id' -Destination $destination

        $result | Should-BeNull
        Test-Path -LiteralPath ($destination | Join-Path -ChildPath 'page-id') | Should-BeFalse
      }
    }
  }
  Describe 'Import-OneNoteHierarchy' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
      $xmlFile = Export-OneNoteHierarchy -Id $notebookId -Destination $notebookPath
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'imports hierarchy by Path' {
        Import-OneNoteHierarchy -Path $xmlFile.FullName
      }
      It 'imports hierarchy by LiteralPath' {
        Import-OneNoteHierarchy -LiteralPath $xmlFile.FullName
      }
      It 'imports hierarchy by Path with ValueFromPipeline' {
        $xmlFile.FullName | Import-OneNoteHierarchy
      }
      It 'imports hierarchy by Path with ValueFromPipelineByPropertyName' {
        [PSCustomObject]@{ LiteralPath = $xmlFile.FullName } | Import-OneNoteHierarchy
      }
    }
    Context 'SupportsShouldProcess' {
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Import-OneNoteHierarchy' -Tag 'CI', 'Mock' {
    BeforeAll {
      function New-OneNoteHierarchyImportMockApplication {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function only constructs an in-memory mock application and does not change system state')]
        param ()
        $app = [PSCustomObject]@{ Updates = [List[string]]::new(); Calls = [List[object]]::new() }
        $app | Add-Member -MemberType ScriptMethod -Name UpdateHierarchy -Value {
          param (
            [string]
            $Content,
            [object]
            $Schema
          )
          [void]$this.Updates.Add($Content)
          [void]$this.Calls.Add([PSCustomObject]@{
              Content = $Content
              Schema  = $Schema
            })
        }
        return $app
      }
    }
    BeforeEach {
      $app = New-OneNoteHierarchyImportMockApplication
      Mock -CommandName Get-OneNoteApplication -MockWith { $app }

      $workPath = $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
      New-Item -Path $workPath -ItemType Directory -Force | Out-Null
      $file1 = $workPath | Join-Path -ChildPath 'a.hierarchy.xml'
      $file2 = $workPath | Join-Path -ChildPath 'b.hierarchy.xml'
      '<Hierarchy id="a" />' | Out-File -LiteralPath $file1 -Encoding UTF8
      '<Hierarchy id="b" />' | Out-File -LiteralPath $file2 -Encoding UTF8
    }
    AfterEach {
      if (Test-Path -LiteralPath $workPath) {
        Remove-Item -LiteralPath $workPath -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'accepts Path wildcard' {
        Import-OneNoteHierarchy -Path ($workPath | Join-Path -ChildPath '*.hierarchy.xml')
        $app.Updates.Count | Should-Be 2
      }
      It 'accepts LiteralPath array' {
        Import-OneNoteHierarchy -LiteralPath @($file1, $file2)
        $app.Updates.Count | Should-Be 2
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update hierarchy when WhatIf is specified' {
        Import-OneNoteHierarchy -LiteralPath $file1 -WhatIf
        $app.Updates.Count | Should-Be 0
      }
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Import-OneNotePageContent' {
    BeforeEach {
      $notebookPath = Get-NotebookPath
      $notebookId = New-OneNoteNotebook -Path $notebookPath
      $sectionPath = Get-SectionPath -NotebookPath $notebookPath
      $sectionId = New-OneNoteSection -Path $sectionPath
      $pageId = New-OneNotePage -SectionId $sectionId
      $xmlFile = Export-OneNotePageContent -Id $pageId -Destination $notebookPath
    }
    AfterEach {
      try {
        Close-OneNoteNotebook -NotebookId $notebookId
      } finally {
        if (Test-Path -Path $notebookPath) {
          Remove-Item -Path $notebookPath -Recurse -Force
        }
      }
    }
    Context 'ParameterSetName' {
      It 'imports page content by Path' {
        Import-OneNotePageContent -Path $xmlFile.FullName
      }
      It 'imports page content by LiteralPath' {
        Import-OneNotePageContent -LiteralPath $xmlFile.FullName
      }
      It 'imports page content by Path with ValueFromPipeline' {
        $xmlFile.FullName | Import-OneNotePageContent
      }
      It 'imports page content by Path with ValueFromPipelineByPropertyName' {
        [PSCustomObject]@{ LiteralPath = $xmlFile.FullName } | Import-OneNotePageContent
      }
    }
    Context 'SupportsShouldProcess' {
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
  Describe 'Import-OneNotePageContent' -Tag 'CI', 'Mock' {
    BeforeAll {
      function New-OneNotePageContentImportMockApplication {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function only constructs an in-memory mock application and does not change system state')]
        param ()
        $app = [PSCustomObject]@{ Updates = [List[string]]::new(); Calls = [List[object]]::new() }
        $app | Add-Member -MemberType ScriptMethod -Name UpdatePageContent -Value {
          param (
            [string]
            $Content,
            [object]
            $ExpectedLastModified,
            [object]
            $Schema,
            [bool]
            $Force
          )
          [void]$this.Updates.Add($Content)
          [void]$this.Calls.Add([PSCustomObject]@{
              Content              = $Content
              ExpectedLastModified = $ExpectedLastModified
              Schema               = $Schema
              Force                = $Force
            })
        }
        return $app
      }
    }
    BeforeEach {
      $app = New-OneNotePageContentImportMockApplication
      Mock -CommandName Get-OneNoteApplication -MockWith { $app }

      $workPath = $env:TEMP | Join-Path -ChildPath ([Guid]::NewGuid().ToString('N'))
      New-Item -Path $workPath -ItemType Directory -Force | Out-Null
      $file1 = $workPath | Join-Path -ChildPath 'a.content.xml'
      $file2 = $workPath | Join-Path -ChildPath 'b.content.xml'
      '<Page id="a" />' | Out-File -LiteralPath $file1 -Encoding UTF8
      '<Page id="b" />' | Out-File -LiteralPath $file2 -Encoding UTF8
    }
    AfterEach {
      if (Test-Path -LiteralPath $workPath) {
        Remove-Item -LiteralPath $workPath -Recurse -Force
      }
    }
    Context 'ParameterSetName' {
      It 'accepts Path wildcard' {
        Import-OneNotePageContent -Path ($workPath | Join-Path -ChildPath '*.content.xml')
        $app.Updates.Count | Should-Be 2

      }
      It 'accepts LiteralPath array' {
        Import-OneNotePageContent -LiteralPath @($file1, $file2)
        $app.Updates.Count | Should-Be 2
      }
    }
    Context 'SupportsShouldProcess' {
      It 'does not update page content when WhatIf is specified' {
        Import-OneNotePageContent -LiteralPath $file1 -WhatIf
        $app.Updates.Count | Should-Be 0
      }
    }
    Context 'Other parameters' {
    }
    Context 'Edge cases' {
    }
  }
}

---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: New-OneNoteSection
---

# New-OneNoteSection

## SYNOPSIS

Creates a new OneNote section file.

## SYNTAX

### PathSet (Default)

```
New-OneNoteSection [-Path] <string> [-WhatIf] [-Confirm]
```

### ParentIdSet

```
New-OneNoteSection [-Path] <string> -ParentId <string> [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Creates a new OneNote section (`.one`) file at the specified path using the OneNote COM API.
Returns the object ID of the created section.

## EXAMPLES

### EXAMPLE 1

```powershell

$sectionId = New-OneNoteSection -Path 'C:\Notes\MySection.one'
$sectionId
```

Creates a new OneNote section at the given path and returns its object ID.

### EXAMPLE 2

```powershell

$path = "$env:TEMP\Section.one"
$sectionId = New-OneNoteSection -Path $path
$pageId = New-OneNotePage -SectionId $sectionId
```

Creates a section (PathSet default), then creates a page within that section.

### EXAMPLE 3

```powershell

$notebookPath = "$env:TEMP\MyNotebook"
$notebookId = New-OneNoteNotebook -Path $notebookPath
$sectionPath = "$notebookPath\Section.one"
$sectionId = New-OneNoteSection -Path $sectionPath -ParentId $notebookId
```

Creates a section by ParentIdSet under the specified notebook.

## PARAMETERS

### -Confirm

Prompts you for confirmation before running the cmdlet.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
Aliases:
  - cf
ParameterSets:
  - Name: (All)
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -ParentId

Required only in ParentIdSet.
The OneNote object ID of the parent notebook or section group.
When specified, only the file name part of Path is used for OpenHierarchy.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: ParentIdSet
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Path

The file system path where the new section file will be created.
The path must be a valid leaf path.
The file does not need to exist beforehand.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: ParentIdSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: PathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -WhatIf

Runs the command in a mode that only reports what would happen without performing the actions.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
Aliases:
  - wi
ParameterSets:
  - Name: (All)
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.String

The OneNote object ID of the created section.

{{ Fill in the Description }}

### System.String

{{ Fill in the Description }}

## NOTES

This function uses the OneNote COM API (OpenHierarchy) internally.
COM objects are released after execution via garbage collection.

## RELATED LINKS

{{ Fill in the related links here }}

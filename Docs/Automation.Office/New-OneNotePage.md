---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: New-OneNotePage
---

# New-OneNotePage

## SYNOPSIS

Creates a new page in a OneNote section.

## SYNTAX

### __AllParameterSets

```
New-OneNotePage [-SectionId] <string> [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Creates a new page in the specified OneNote section using the OneNote COM API.
Returns the object ID of the created page.

## EXAMPLES

### EXAMPLE 1

```powershell

$path = "$env:TEMP\NoteBook.one"
$sectionId = New-OneNoteSection -Path $path
$pageId = New-OneNotePage -SectionId $sectionId
$pageId
```

Creates a new section, then creates a page within it and returns its object ID.

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

### -SectionId

The object ID of the section in which the new page will be created.
Obtain this value from New-OneNoteSection or Get-OneNoteHierarchy.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
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

The OneNote object ID of the created page.

{{ Fill in the Description }}

### System.String

{{ Fill in the Description }}

## NOTES

This function uses the OneNote COM API (CreateNewPage) internally.
The page is created with the default page style (npsDefault).
COM objects are released after execution via garbage collection.

## RELATED LINKS

{{ Fill in the related links here }}

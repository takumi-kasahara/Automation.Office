---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-OneNoteHierarchy
---

# Get-OneNoteHierarchy

## SYNOPSIS

Retrieves the OneNote hierarchy as XML nodes.

## SYNTAX

### __AllParameterSets

```
Get-OneNoteHierarchy [[-Id] <string>] [[-HierarchyScope] <HierarchyScope>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Retrieves the hierarchy of notebooks, section groups, sections, or pages
starting from the specified object using the OneNote COM API.
Returns the parsed XML document element.

## EXAMPLES

### EXAMPLE 1

```powershell

$notebookId = New-OneNoteNotebook -Path 'C:\Notes\Work'
$hierarchy = Get-OneNoteHierarchy -Id $notebookId
$hierarchy
```

Retrieves the hierarchy of the specified notebook.

### EXAMPLE 2

```powershell

$hierarchy = Get-OneNoteHierarchy -Id $notebookId -HierarchyScope hsPages
$hierarchy.Section.Page | Select-Object -ExpandProperty name
```

Retrieves the full page hierarchy under the specified notebook.

## PARAMETERS

### -HierarchyScope

The depth of the hierarchy to retrieve.
Defaults to hsSelf (only the specified object).
See HierarchyScope enumeration for available values (e.g., hsNotebooks, hsSections, hsPages).

```yaml
Type: Microsoft.Office.Interop.OneNote.HierarchyScope
DefaultValue: hsSelf
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: 1
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Id

The OneNote object ID of the starting node.
If omitted or empty, the hierarchy starts from the root.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: 0
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

### System.Xml.XmlNode

The XML node representing the requested hierarchy element.

{{ Fill in the Description }}

### System.Xml.XmlNode

{{ Fill in the Description }}

## NOTES

This function uses the OneNote COM API (GetHierarchy) internally.
COM objects are released after execution via garbage collection.

## RELATED LINKS

{{ Fill in the related links here }}

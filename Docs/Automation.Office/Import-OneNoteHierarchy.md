---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Import-OneNoteHierarchy
---

# Import-OneNoteHierarchy

## SYNOPSIS

Imports a OneNote hierarchy XML file into OneNote.

## SYNTAX

### PathSet

```
Import-OneNoteHierarchy [-Path] <string[]> [-WhatIf] [-Confirm]
```

### LiteralPathSet

```
Import-OneNoteHierarchy -LiteralPath <string[]> [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Reads a hierarchy XML file previously exported by Export-OneNoteHierarchy and applies it to the OneNote application using the UpdateHierarchy COM API method.
Accepts files by path (with wildcard support) or by literal path.

## EXAMPLES

### EXAMPLE 1

```powershell

$notebookId = New-OneNoteNotebook -Path 'C:\Notes\Work'
$xmlFile = Export-OneNoteHierarchy -Id $notebookId -Destination $env:TEMP
Import-OneNoteHierarchy -LiteralPath $xmlFile.FullName
```

Exports the hierarchy of a notebook and re-imports it from the XML file.

### EXAMPLE 2

```powershell

Get-ChildItem -Path "$env:TEMP\*.hierarchy.xml" | Import-OneNoteHierarchy
```

Imports all hierarchy XML files in the specified directory via pipeline.

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

### -LiteralPath

Required in LiteralPathSet.
One or more literal file paths to the hierarchy XML files to import.
No wildcard expansion is performed.
Accepts pipeline input by property name.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases:
  - PSPath
  - LP
ParameterSets:
  - Name: LiteralPathSet
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Path

Required in PathSet.
One or more file paths to the hierarchy XML files to import.
Accepts wildcards.
Accepts pipeline input by value and property name.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: PathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: true
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

### System.String[]

{{ Fill in the Description }}

## OUTPUTS

### System.Void

{{ Fill in the Description }}

## NOTES

This function uses the OneNote COM API (UpdateHierarchy) internally.
COM objects are released after execution via garbage collection.

## RELATED LINKS

{{ Fill in the related links here }}

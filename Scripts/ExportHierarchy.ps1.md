---
document type: cmdlet
external help file: ExportHierarchy.ps1-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: ''
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: ExportHierarchy.ps1
---

# ExportHierarchy.ps1

## SYNOPSIS

Exports OneNote hierarchy to CSV and XML files.

## SYNTAX

### __AllParameterSets

```
ExportHierarchy.ps1
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

This script exports the OneNote hierarchy structure to both CSV and XML formats.
It retrieves the hierarchy using Get-OneNoteHierarchy, filters out recycle bin items, and exports the data to a CSV file and individual XML files for each node.

## EXAMPLES

## PARAMETERS

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

The script expects a Config.psd1 file with Output and StartNodeId properties.

## RELATED LINKS

{{ Fill in the related links here }}

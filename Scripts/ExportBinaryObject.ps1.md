---
document type: cmdlet
external help file: ExportBinaryObject.ps1-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: ''
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: ExportBinaryObject.ps1
---

# ExportBinaryObject.ps1

## SYNOPSIS

Exports binary objects from OneNote pages to files.

## SYNTAX

### __AllParameterSets

```
ExportBinaryObject.ps1
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

This script exports binary objects (attachments, images, etc.) from OneNote pages to a specified output directory.
It uses the Get-OneNoteHierarchy cmdlet to retrieve the page hierarchy and then exports binary objects from each page.

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

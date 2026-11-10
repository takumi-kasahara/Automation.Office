---
document type: cmdlet
external help file: EditPageContent.ps1-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: ''
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: EditPageContent.ps1
---

# EditPageContent.ps1

## SYNOPSIS

Edits page content in OneNote documents by removing language attributes.

## SYNTAX

### __AllParameterSets

```
EditPageContent.ps1
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

This script processes OneNote XML files to remove language attributes from Page elements and OE elements.
It extracts the XML files from a compressed archive, modifies them to remove language attributes, and then recreates the archive with the updated files.

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

The script expects a Config.psd1 file in the same directory with an Output property specifying the base directory for processing.

## RELATED LINKS

{{ Fill in the related links here }}

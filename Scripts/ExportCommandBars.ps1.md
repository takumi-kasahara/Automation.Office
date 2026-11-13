---
document type: cmdlet
external help file: ExportCommandBars.ps1-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: ''
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: ExportCommandBars.ps1
---

# ExportCommandBars.ps1

## SYNOPSIS

Exports CommandBar configurations from Office applications to CSV files.

## SYNTAX

### __AllParameterSets

```
ExportCommandBars.ps1
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

This script exports CommandBar configurations from Access, Excel, PowerPoint, and Word applications to CSV files.
It creates COM objects for each application, retrieves the CommandBars and their controls, and exports the data to CSV files in the `.\Out` directory.
For Access, it also exports VBA CommandBars if available.

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

The script creates the `.\Out` directory if it doesn't exist.

## RELATED LINKS

{{ Fill in the related links here }}

---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-PowerPointVersionInfo
---

# Get-PowerPointVersionInfo

## SYNOPSIS

Retrieves the version information of the PowerPoint application.

## SYNTAX

### __AllParameterSets

```
Get-PowerPointVersionInfo
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Returns the FileVersionInfo object for the PowerPoint executable (POWERPNT.EXE) located in the PowerPoint installation directory.
This helps determine which version of PowerPoint is currently installed and available for automation.

## EXAMPLES

## PARAMETERS

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.Diagnostics.FileVersionInfo

{{ Fill in the Description }}

## NOTES

Requires PowerPoint to be installed and accessible via COM automation.
The function uses the New-PowerPointObject helper to locate the PowerPoint executable.

## RELATED LINKS

{{ Fill in the related links here }}

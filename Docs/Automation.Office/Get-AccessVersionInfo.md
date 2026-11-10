---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-AccessVersionInfo
---

# Get-AccessVersionInfo

## SYNOPSIS

Retrieves the version information of the Access application.

## SYNTAX

### __AllParameterSets

```
Get-AccessVersionInfo
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Returns the FileVersionInfo object for the Access executable (ACCESS.EXE) located in the Access installation directory.
This helps determine which version of Access is currently installed and available for automation.

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

Requires Access to be installed and accessible via COM automation.
The function uses the New-AccessObject helper to locate the Access executable.

## RELATED LINKS

{{ Fill in the related links here }}

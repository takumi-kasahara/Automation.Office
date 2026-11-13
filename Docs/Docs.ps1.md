---
document type: cmdlet
external help file: Docs.ps1-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: ''
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Docs.ps1
---

# Docs.ps1

## SYNOPSIS

Generates Markdown command documentation for the repository PowerShell scripts.

## SYNTAX

### __AllParameterSets

```
Docs.ps1
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Generates or refreshes Markdown help files using PlatyPS (New-MarkdownCommandHelp) for all *.ps1 scripts found in the repository root and in selected skill script folders.
Then applies Prettier formatting and markdownlint-cli2 auto-fixes to the generated Markdown files.
The script changes its working directory to the script root before processing.

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

Requires the platyPS (New-MarkdownCommandHelp) command to be available.
Requires Node.js tooling for Prettier and markdownlint-cli2.
Regenerated files are written in-place and may overwrite existing help files.

## RELATED LINKS

{{ Fill in the related links here }}

---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Export-WordVBProject
---

# Export-WordVBProject

## SYNOPSIS

Exports VBA project metadata and components from a Word document.

## SYNTAX

### __AllParameterSets

```
Export-WordVBProject [-Path] <string> -Destination <string> [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-ComponentRoot <string>] [-Force] [-NoClobber] [-WhatIf]
 [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens a file and exports references and components into a unified `VBProject.json` format.
The companion component folder is created next to the destination JSON unless `-ComponentRoot` is specified.

## EXAMPLES

### EXAMPLE 1

```powershell

Export-WordVBProject -Path "$env:TEMP\Document.docm" -Destination "$env:TEMP\VBProject.json"
```

Exports VBA project metadata and components from a Word document to VBProject.json and creates a companion component folder.

### EXAMPLE 2

```powershell

$password = Read-Host -AsSecureString
Export-WordVBProject -Path "$env:TEMP\Document.docm" -PasswordToOpen $password -Destination "$env:TEMP\VBProject.json"
```

Exports VBA project from a document protected with an open password.

### EXAMPLE 3

```powershell

Export-WordVBProject -Path "$env:TEMP\Document.docm" -PasswordToModify $password -Destination "$env:TEMP\VBProject.json"
```

Exports VBA project from a document protected with a modify password.

### EXAMPLE 4

```powershell

Export-WordVBProject -Path "$env:TEMP\Document.docm" -Destination "$env:TEMP\VBProject.json" -WhatIf
```

Shows what would happen if the command were executed without actually exporting the VBA project.

## PARAMETERS

### -ComponentRoot

Specifies the directory path where component files are exported.
When omitted, the directory is derived from `-Destination` by removing its extension.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
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

### -Destination

Specifies the export file path for `VBProject.json`.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Force

Overrides read-only destination file attributes.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
SupportsWildcards: false
Aliases: []
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

### -NoClobber

Produces an error if -Destination already exists.
If omitted, existing output is overwritten.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
SupportsWildcards: false
Aliases: []
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

### -PasswordToModify

Specifies the password required to modify the document.

```yaml
Type: System.Security.SecureString
DefaultValue: ''
SupportsWildcards: false
Aliases: []
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

### -PasswordToOpen

Specifies the password required to open the document.

```yaml
Type: System.Security.SecureString
DefaultValue: ''
SupportsWildcards: false
Aliases: []
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

### -Path

Specifies the document path literally.
Wildcards are not interpreted.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
  - FilePath
  - FullName
ParameterSets:
  - Name: (All)
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

### System.String

{{ Fill in the Description }}

## OUTPUTS

### System.IO.FileInfo

Returns the created export file.

{{ Fill in the Description }}

### System.IO.FileInfo

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

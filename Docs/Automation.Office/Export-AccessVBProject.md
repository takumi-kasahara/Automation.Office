---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Export-AccessVBProject
---

# Export-AccessVBProject

## SYNOPSIS

Exports VBA project metadata and components from an Access database.

## SYNTAX

### __AllParameterSets

```
Export-AccessVBProject [-Path] <string> -Destination <string> [-Password <securestring>]
 [-ComponentRoot <string>] [-Force] [-NoClobber] [-IncludeAcObjectType <AcObjectType[]>]
 [-ExcludeAcObjectType <AcObjectType[]>] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens a database and exports references and components into a unified `VBProject.json` format.
The companion component folder is created next to the destination JSON unless `-ComponentRoot` is specified.

## EXAMPLES

### EXAMPLE 1

```powershell

Export-AccessVBProject -Path "$env:TEMP\Database.accdb" -Destination "$env:TEMP\VBProject.json"
```

Exports the VBA project from the Access database to VBProject.json and creates a companion component folder.

### EXAMPLE 2

```powershell

Export-AccessVBProject -Path "$env:TEMP\Database.accdb" -Destination "$env:TEMP\VBProject.json" -ComponentRoot "$env:TEMP\Components"
```

Exports the VBA project to VBProject.json and places component files in the specified ComponentRoot directory.

### EXAMPLE 3

```powershell

Export-AccessVBProject -Path "$env:TEMP\Database.accdb" -Destination "$env:TEMP\VBProject.json" -IncludeAcObjectType acForm
```

Exports only form components from the VBA project.

### EXAMPLE 4

```powershell

Export-AccessVBProject -Path "$env:TEMP\Database.accdb" -Destination "$env:TEMP\VBProject.json" -ExcludeAcObjectType acReport
```

Exports all components except reports from the VBA project.

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

### -ExcludeAcObjectType

Specifies an array of AcObjectType values to exclude when exporting the VBProject.
When specified, the specified object types are omitted from the export.
Corresponds to the ExcludeAcObjectType parameter in the Access VBA object model.

```yaml
Type: Microsoft.Office.Interop.Access.AcObjectType[]
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

### -IncludeAcObjectType

Specifies an array of AcObjectType values to include when exporting the VBProject.
When specified, only the specified object types are included in the export.
Corresponds to the IncludeAcObjectType parameter in the Access VBA object model.

```yaml
Type: Microsoft.Office.Interop.Access.AcObjectType[]
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

### -Password

Specifies the password required to open the database.

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

Specifies the database path literally.
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

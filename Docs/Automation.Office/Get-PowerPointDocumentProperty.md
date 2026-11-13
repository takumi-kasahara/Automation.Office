---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-PowerPointDocumentProperty
---

# Get-PowerPointDocumentProperty

## SYNOPSIS

Gets document properties from presentations.

## SYNTAX

### PathSet (Default)

```
Get-PowerPointDocumentProperty [-Path] <string[]> [[-Name] <string[]>]
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-Custom]
```

### LiteralPathSet

```
Get-PowerPointDocumentProperty [[-Name] <string[]>] -LiteralPath <string[]>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-Custom]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more presentations and returns document properties as a `PSCustomObject`.
The cmdlet supports both `-Path` and `-LiteralPath` parameter sets, optional name filtering with `-Name`, and protected presentations via `-PasswordToOpen` and `-PasswordToModify`.
When `-Custom` is specified, the cmdlet reads custom document properties instead of built-in properties.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Name Title
```

Gets the built-in `Title` property from a file.

### EXAMPLE 2

```powershell

$password = Read-Host -AsSecureString
Get-PowerPointDocumentProperty -LiteralPath "$env:TEMP\Presentation.pptx" -PasswordToOpen $password -Name Title
```

Gets properties from a file protected with an open password.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
Get-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -PasswordToModify $password -Name Title
```

Gets properties from a file protected with a modify password.

### EXAMPLE 4

```powershell

Get-PowerPointDocumentProperty -Path "$env:TEMP\Presentation.pptx" -Custom -Name MyCustomProperty
```

Gets a custom document property.

## PARAMETERS

### -Custom

Gets custom document properties.

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

### -LiteralPath

Specifies presentation paths literally.
Wildcards are not interpreted.

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

### -Name

Filters returned properties by name.
Throws if any specified property does not exist.
If not specified, all properties are returned.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: 1
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -PasswordToModify

Specifies the password required to modify the presentation.

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

Specifies the password required to open the presentation.

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

Specifies presentation paths.
Wildcards are supported.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: true
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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String[]

{{ Fill in the Description }}

## OUTPUTS

### System.Management.Automation.PSObject

Each NoteProperty corresponds to a file property.

{{ Fill in the Description }}

### System.Management.Automation.PSObject

{{ Fill in the Description }}

## NOTES

If a file cannot be opened (for example, incorrect password), the cmdlet throws an error record and continues processing remaining items.

## RELATED LINKS

{{ Fill in the related links here }}

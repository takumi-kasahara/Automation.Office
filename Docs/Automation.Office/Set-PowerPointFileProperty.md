---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Set-PowerPointFileProperty
---

# Set-PowerPointFileProperty

## SYNOPSIS

Sets presentation file properties.

## SYNTAX

### ValuePathSet (Default)

```
Set-PowerPointFileProperty [-Path] <string[]> [-Name] <string> [-Value] <Object>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-PassThru] [-WhatIf]
 [-Confirm]
```

### PSObjectPathSet

```
Set-PowerPointFileProperty [-Path] <string[]> [-InputObject] <psobject>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-PassThru] [-WhatIf]
 [-Confirm]
```

### PSObjectLiteralPathSet

```
Set-PowerPointFileProperty [-InputObject] <psobject> -LiteralPath <string[]>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-PassThru] [-WhatIf]
 [-Confirm]
```

### ValueLiteralPathSet

```
Set-PowerPointFileProperty [-Name] <string> [-Value] <Object> -LiteralPath <string[]>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-PassThru] [-WhatIf]
 [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more presentations and updates a file property.
The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
If a target is read-only, the cmdlet throws an exception.

## EXAMPLES

### EXAMPLE 1

```powershell

Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final -Value $true
```

Sets a single file property.

### EXAMPLE 2

```powershell

Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -InputObject ([PSCustomObject]@{ Final = $true })
```

Updates multiple file properties in a single call.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final -Value $true -PasswordToOpen $password
```

Updates file properties on a protected presentation.

### EXAMPLE 4

```powershell

Set-PowerPointFileProperty -Path "$env:TEMP\Presentation.pptx" -Name Final -Value $true -PassThru
```

Updates the property and returns the updated value.

## PARAMETERS

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

### -InputObject

Specifies a `PSObject` whose properties are mapped to file property names and values.
Use this parameter to update multiple properties at once.

```yaml
Type: System.Management.Automation.PSObject
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: PSObjectLiteralPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: PSObjectPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: true
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
  - Name: PSObjectLiteralPathSet
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
  - Name: ValueLiteralPathSet
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

Specifies the file property name to update.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: ValueLiteralPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: ValuePathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -PassThru

Returns the updated file property object when specified.

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
  - Name: PSObjectPathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
  - Name: ValuePathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Value

Specifies the value assigned to the property.

```yaml
Type: System.Object
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: ValueLiteralPathSet
    Position: 2
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: ValuePathSet
    Position: 2
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
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

### System.String[]

{{ Fill in the Description }}

### System.String

{{ Fill in the Description }}

### System.Management.Automation.PSObject

{{ Fill in the Description }}

## OUTPUTS

### System.Management.Automation.PSObject

When `-PassThru` is specified.

{{ Fill in the Description }}

### System.Management.Automation.PSObject

{{ Fill in the Description }}

### System.Void

{{ Fill in the Description }}

## NOTES

If a file cannot be updated (for example, read-only state, invalid property name, or file access issues), the cmdlet throws an exception and continues processing remaining items.

## RELATED LINKS

{{ Fill in the related links here }}

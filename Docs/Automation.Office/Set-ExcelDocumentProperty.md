---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Set-ExcelDocumentProperty
---

# Set-ExcelDocumentProperty

## SYNOPSIS

Sets built-in or custom document properties on workbooks.

## SYNTAX

### ValuePathSet (Default)

```
Set-ExcelDocumentProperty [-Path] <string[]> [-Name] <string> [-Value] <Object>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-Custom] [-PassThru] [-Force]
 [-WhatIf] [-Confirm]
```

### PSObjectPathSet

```
Set-ExcelDocumentProperty [-Path] <string[]> [-InputObject] <psobject>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-Custom] [-PassThru] [-Force]
 [-WhatIf] [-Confirm]
```

### PSObjectLiteralPathSet

```
Set-ExcelDocumentProperty [-InputObject] <psobject> -LiteralPath <string[]>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-Custom] [-PassThru] [-Force]
 [-WhatIf] [-Confirm]
```

### ValueLiteralPathSet

```
Set-ExcelDocumentProperty [-Name] <string> [-Value] <Object> -LiteralPath <string[]>
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-Custom] [-PassThru] [-Force]
 [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more workbooks and updates a built-in or custom document property.
The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
Use `-Name` and `-Value` to update a single property, or `-InputObject` to update multiple properties in one operation.
By default, read-only items and read-only recommended workbooks throw an exception.
Use `-Force` to ignore the read-only recommended flag when opening a file.

## EXAMPLES

### EXAMPLE 1

```powershell

Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title -Value 'Monthly report'
```

Sets the `Title` built-in document property.

### EXAMPLE 2

```powershell

Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Custom -Name MyProperty -Value 'Custom value'
```

Sets the `MyProperty` custom document property.

### EXAMPLE 3

```powershell

Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -InputObject ([PSCustomObject]@{ Title = 'Monthly report'; Subject = 'Sales' })
```

Updates multiple built-in document properties in a single call.

### EXAMPLE 4

```powershell

$password = Read-Host -AsSecureString
Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title -Value 'Protected book' -PasswordToOpen $password -Force
```

Updates a protected workbook.

### EXAMPLE 5

```powershell

Set-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -Name Title -Value 'Monthly report' -PassThru
```

Updates the `Title` built-in document property and returns the updated property object.

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

### -Custom

Updates a custom document property instead of a built-in property.

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

### -Force

Forces update processing when a file is marked as read-only recommended.

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

### -InputObject

Specifies a `PSObject` whose properties are mapped to document property names and values.
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

Specifies workbook paths literally.
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

Specifies the document property name to update.
When `-Custom` is specified, the cmdlet updates a custom document property.

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

Returns the updated document property object when specified.

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

Specifies the password required to modify the workbook.

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

Specifies the password required to open the workbook.

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

Specifies workbook paths.
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

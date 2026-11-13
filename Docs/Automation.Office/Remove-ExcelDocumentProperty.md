---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Remove-ExcelDocumentProperty
---

# Remove-ExcelDocumentProperty

## SYNOPSIS

Removes document properties from workbooks.

## SYNTAX

### PathSet (Default)

```
Remove-ExcelDocumentProperty [-Path] <string[]> [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-RemoveDocInfoType <XlRemoveDocInfoType>] [-Force] [-WhatIf]
 [-Confirm]
```

### LiteralPathSet

```
Remove-ExcelDocumentProperty -LiteralPath <string[]> [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-RemoveDocInfoType <XlRemoveDocInfoType>] [-Force] [-WhatIf]
 [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more workbooks and removes document information.
The cmdlet supports both `-Path` and `-LiteralPath` parameter sets.
By default, read-only items and read-only recommended workbooks throw an exception.
Use `-Force` to ignore the read-only recommended flag when opening a file.

## EXAMPLES

### EXAMPLE 1

```powershell

Remove-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx"
```

Removes document properties from a file.

### EXAMPLE 2

```powershell

$password = Read-Host -AsSecureString
Remove-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -PasswordToOpen $password -Force
```

Removes document properties from a protected workbook.

### EXAMPLE 3

```powershell

Remove-ExcelDocumentProperty -Path "$env:TEMP\Workbook.xlsx" -RemoveDocInfoType xlRDIAll
```

Removes all supported document information categories.

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

### -Force

Forces removal processing when a file is marked as read-only recommended.

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

### -RemoveDocInfoType

Specifies the document information type removed by Excel.
The default value is `xlRDIDocumentProperties`.

```yaml
Type: Microsoft.Office.Interop.Excel.XlRemoveDocInfoType
DefaultValue: xlRDIDocumentProperties
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

## OUTPUTS

### System.Void

{{ Fill in the Description }}

## NOTES

If a file cannot be updated (for example, read-only state or file access issues), the cmdlet throws an exception and continues processing remaining items.

## RELATED LINKS

{{ Fill in the related links here }}

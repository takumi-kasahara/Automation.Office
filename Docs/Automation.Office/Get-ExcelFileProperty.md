---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-ExcelFileProperty
---

# Get-ExcelFileProperty

## SYNOPSIS

Gets workbook file properties.

## SYNTAX

### PathSet (Default)

```
Get-ExcelFileProperty [-Path] <string[]> [[-Name] <string[]>] [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-Depth <int>]
```

### LiteralPathSet

```
Get-ExcelFileProperty [[-Name] <string[]>] -LiteralPath <string[]> [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-Depth <int>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more workbooks and returns file properties as a `PSCustomObject`.
Some properties are skipped with a warning at every depth because their getters can block automation.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -Name SaveLinkValues
```

Gets the `SaveLinkValues` workbook property from a file.

### EXAMPLE 2

```powershell

$password = Read-Host -AsSecureString
Get-ExcelFileProperty -LiteralPath "$env:TEMP\Book.xlsx" -PasswordToOpen $password -Name SaveLinkValues
```

Gets file properties from a file protected with an open password.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
Get-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -PasswordToModify $password -Name SaveLinkValues
```

Gets file properties from a file protected with a modify password.

### EXAMPLE 4

```powershell

Get-ExcelFileProperty -Path "$env:TEMP\Book.xlsx" -Depth 2
```

Gets file properties including nested COM object properties up to 2 levels deep.

## PARAMETERS

### -Depth

Specifies the number of nested COM object levels to retrieve.
The default value is `1`.

```yaml
Type: System.Int32
DefaultValue: 1
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

### -Name

Filters returned properties by name.
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

---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-AccessFileProperty
---

# Get-AccessFileProperty

## SYNOPSIS

Gets database file properties.

## SYNTAX

### PathSet (Default)

```
Get-AccessFileProperty [-Path] <string[]> [[-Name] <string[]>] [-Password <securestring>]
 [-Depth <int>]
```

### LiteralPathSet

```
Get-AccessFileProperty [[-Name] <string[]>] -LiteralPath <string[]> [-Password <securestring>]
 [-Depth <int>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more databases and returns database file properties as a `PSCustomObject`.
Some properties are skipped with a warning at every depth because their getters can block automation.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-AccessFileProperty -Path "$env:TEMP\Database.accdb" -Name RemovePersonalInformation
```

Gets the `RemovePersonalInformation` file property from a database.

### EXAMPLE 2

```powershell

$password = Read-Host -AsSecureString
Get-AccessFileProperty -LiteralPath "$env:TEMP\Database.accdb" -Password $password -Name RemovePersonalInformation
```

Gets file properties from a file protected with a password.

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

Specifies database paths literally.
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
If not specified, all file properties are returned.

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

Specifies database paths.
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

If a database cannot be opened (for example, incorrect password), the cmdlet throws an error record and continues processing remaining items.

## RELATED LINKS

{{ Fill in the related links here }}

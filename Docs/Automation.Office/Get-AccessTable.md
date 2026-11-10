---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-AccessTable
---

# Get-AccessTable

## SYNOPSIS

Gets tables in Access database files.

## SYNTAX

### PathSet (Default)

```
Get-AccessTable [-Path] <string[]> [-Password <securestring>]
```

### LiteralPathSet

```
Get-AccessTable -LiteralPath <string[]> [-Password <securestring>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more Access database files (.accdb, .mdb) through OLE DB/ODBC and returns table entries.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-AccessTable -Path '$env:TEMP\Database.accdb'
```

Gets all tables in the specified Access database file.

### EXAMPLE 2

```powershell

Get-AccessTable -LiteralPath '$env:TEMP\Database.accdb'
```

Gets all tables using a literal path (wildcards not interpreted).

### EXAMPLE 3

```powershell

'$env:TEMP\Database.accdb' | Get-AccessTable
```

Gets all tables by piping the file path to the cmdlet.

### EXAMPLE 4

```powershell

Get-AccessTable -Path '$env:TEMP\*.accdb'
```

Gets all tables from all .accdb files in the specified directory using wildcards.

### EXAMPLE 5

```powershell

Get-AccessTable -LiteralPath '$env:TEMP\Database.mdb'
```

Gets all tables from an .mdb format database file.

### EXAMPLE 6

```powershell

$password = Read-Host -AsSecureString
Get-AccessTable -LiteralPath '$env:TEMP\Database.accdb' -Password $password
```

Gets all tables from a password-protected database.

## PARAMETERS

### -LiteralPath

Specifies Access database file paths literally.
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

### -Password

Specifies the password required to open a protected Access database.

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

Specifies Access database file paths.
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

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Open-AccessFile
---

# Open-AccessFile

## SYNOPSIS

Opens an Access database file.

## SYNTAX

### __AllParameterSets

```
Open-AccessFile [-Path] <string> [-Application <__ComObject>] [-Password <securestring>]
 [-ActionDb <scriptblock>] [-ActionProject <scriptblock>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens a database file by automating Access through COM.

## EXAMPLES

### EXAMPLE 1

```powershell

Open-AccessFile -Path "$env:TEMP\Database.accdb"
```

Opens a database and returns the CurrentProject object.

### EXAMPLE 2

```powershell

Open-AccessFile -Path "$env:TEMP\Database.accdb" -ActionDb {
  param ($Database)
  $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
}
```

Opens a database, creates a table using DAO, and closes the database.

### EXAMPLE 3

```powershell

Open-AccessFile -Path "$env:TEMP\Database.accdb" -ActionProject {
  param ($Project)
  return $Project.Connection
}
```

Opens a database and returns the connection string.

## PARAMETERS

### -ActionDb

Specifies a script block to execute with the DAO database object.
The script block receives the database object as its first argument.
When this parameter is specified, the database is automatically closed after the action completes.

```yaml
Type: System.Management.Automation.ScriptBlock
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

### -ActionProject

Specifies a script block to execute with the CurrentProject object.
The script block receives the CurrentProject object as its first argument.
When this parameter is specified, the database is automatically closed after the action completes.

```yaml
Type: System.Management.Automation.ScriptBlock
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

### -Application

Specifies the Access Application COM object to use for opening the database.
If not specified, a new Access Application object is created automatically.

```yaml
Type: System.__ComObject
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

Specifies the path to the database file to open.
This parameter does not support wildcards because it represents an existing file path.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String

{{ Fill in the Description }}

## OUTPUTS

### System.__ComObject

{{ Fill in the Description }}

## NOTES

If neither ActionDb nor ActionProject is specified, do not return anything.

## RELATED LINKS

{{ Fill in the related links here }}

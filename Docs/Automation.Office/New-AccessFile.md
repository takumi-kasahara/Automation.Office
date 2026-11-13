---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: New-AccessFile
---

# New-AccessFile

## SYNOPSIS

Creates a new Access database file.

## SYNTAX

### __AllParameterSets

```
New-AccessFile [-Path] <string> [-FileFormat <AcNewDatabaseFormat>] [-Password <securestring>]
 [-Force] [-RemovePersonalInformation] [-InitializeDb <scriptblock>]
 [-InitializeProject <scriptblock>] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Creates a database at the specified path by automating Access through COM.
If the destination file already exists, the command stops unless `-Force` is specified.

## EXAMPLES

### EXAMPLE 1

```powershell

New-AccessFile -Path "$env:TEMP\Database.accdb"
```

Creates a new database in the temporary directory.

### EXAMPLE 2

```powershell

New-AccessFile -Path "$env:TEMP\Database.mdb" -FileFormat acNewDatabaseFormatAccess2000 -Force
```

Creates or overwrites a legacy Access database.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
New-AccessFile -Path "$env:TEMP\Database.accdb" -Password $password
```

Creates a file protected with an open password.

### EXAMPLE 4

```powershell

New-AccessFile -Path "$env:TEMP\Database.accdb" -RemovePersonalInformation
```

Creates a new database and removes personal information from a file.

### EXAMPLE 5

```powershell

New-AccessFile -Path "$env:TEMP\Database.accdb" -InitializeDb {
  param ($Database)
  $Database.Execute('CREATE TABLE Employees (Id INTEGER, Name TEXT(255))')
}
```

Creates a database and creates a table using DAO.

### EXAMPLE 6

```powershell

New-AccessFile -Path "$env:TEMP\Database.accdb" -InitializeProject {
  param ($Project)
  return $Project.Connection
}
```

Returns the connection string of the new database project.

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

### -FileFormat

Specifies the Access file format used when creating the database.
The default value is `acNewDatabaseFormatUserDefault`.

```yaml
Type: Microsoft.Office.Interop.Access.AcNewDatabaseFormat
DefaultValue: acNewDatabaseFormatUserDefault
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

Overwrites an existing file at `-Path`.
Without this switch, the cmdlet stops when the destination file already exists.

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

### -InitializeDb

Specifies a script block to initialize the database before saving.
The script block receives the DAO database object as its first argument.

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

### -InitializeProject

Specifies a script block to initialize the database project before saving.
The script block receives the CurrentProject object as its first argument.

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

Specifies the destination path of the database file to create.
This parameter does not support wildcards because it represents a new file path.

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

### -RemovePersonalInformation

Removes personal information from the new database file.

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

{{ Fill in the Description }}

## NOTES

The -Password and -RemovePersonalInformation parameters cannot be specified together.
Setting RemovePersonalInformation triggers a save operation that displays a password dialog.

## RELATED LINKS

{{ Fill in the related links here }}

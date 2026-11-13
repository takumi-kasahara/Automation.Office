---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Invoke-AccessQuery
---

# Invoke-AccessQuery

## SYNOPSIS

Executes a SQL statement against Access database files and returns rows or affected-row metadata.

## SYNTAX

### TablePathSet (Default)

```
Invoke-AccessQuery [-Path] <string[]> [-Table] <string[]> [-Columns <string[]>]
 [-Password <securestring>]
```

### QueryPathSet

```
Invoke-AccessQuery [-Path] <string[]> [-Query] <string> [-Password <securestring>]
```

### ViewPathSet

```
Invoke-AccessQuery [-Path] <string[]> [-View] <string[]> [-Columns <string[]>]
 [-Password <securestring>]
```

### QueryLiteralPathSet

```
Invoke-AccessQuery [-Query] <string> -LiteralPath <string[]> [-Password <securestring>]
```

### ViewLiteralPathSet

```
Invoke-AccessQuery [-View] <string[]> -LiteralPath <string[]> [-Columns <string[]>]
 [-Password <securestring>]
```

### TableLiteralPathSet

```
Invoke-AccessQuery [-Table] <string[]> -LiteralPath <string[]> [-Columns <string[]>]
 [-Password <securestring>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more Access database files (.accdb, .mdb) through OLE DB and executes the specified SQL statement.
For `SELECT` statements, rows are returned as PSCustomObject with Path, ObjectType, ObjectName, and column properties.
For `INSERT`, `UPDATE`, and `DELETE` statements, a single PSCustomObject with Path, ObjectType, ObjectName, and RecordsAffected is returned.

## EXAMPLES

### EXAMPLE 1

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -Table 'Employees'
```

Gets all rows from the Employees table in the specified Access database.

### EXAMPLE 2

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -Table 'Employees' -Columns Id, Name, Department
```

Gets only the Id, Name, and Department columns from the Employees table.

### EXAMPLE 3

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -Query 'SELECT [Id], [Name] FROM Employees WHERE [Department] = "Sales"'
```

Executes a custom SQL query against the database.

### EXAMPLE 4

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -Query "INSERT INTO [Employees] (Id, Name, Department) VALUES (4, 'David', 'Marketing')"
```

Inserts a new row and returns the number of affected rows.

### EXAMPLE 5

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -Query "UPDATE [Employees] SET [Department] = 'Sales' WHERE [Id] = 4"
```

Updates existing rows and returns the number of affected rows.

### EXAMPLE 6

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -Query 'DELETE FROM [Employees] WHERE [Id] = 4'
```

Deletes rows and returns the number of affected rows.

### EXAMPLE 7

```powershell

Invoke-AccessQuery -LiteralPath '$env:TEMP\Database.accdb' -View vwSalesEmployees
```

Gets all rows from the vwSalesEmployees view.

### EXAMPLE 8

```powershell

$password = Read-Host -AsSecureString
Invoke-AccessQuery -LiteralPath '$env:TEMP\Protected.accdb' -Query 'SELECT * FROM Employees' -Password $password
```

Executes a query against a password-protected database.

## PARAMETERS

### -Columns

Specifies the columns to return when querying tables or views by name.

When omitted, all columns are returned.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: ViewLiteralPathSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: ViewPathSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: TableLiteralPathSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: TablePathSet
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
  - Name: QueryLiteralPathSet
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
  - Name: ViewLiteralPathSet
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
  - Name: TableLiteralPathSet
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
  - Name: QueryPathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
  - Name: ViewPathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
  - Name: TablePathSet
    Position: 0
    IsRequired: true
    ValueFromPipeline: true
    ValueFromPipelineByPropertyName: true
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Query

Specifies a SQL statement to execute against the database.
`SELECT`, `INSERT`, `UPDATE`, and `DELETE` are supported.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: QueryLiteralPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: QueryPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Table

Specifies one or more table names to query.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: TableLiteralPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: TablePathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -View

Specifies one or more view names to query.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: ViewLiteralPathSet
    Position: 1
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: ViewPathSet
    Position: 1
    IsRequired: true
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

### System.Management.Automation.PSObject

{{ Fill in the Description }}

## NOTES

SQL syntax is based on the Microsoft Access SQL dialect.
For a complete SQL reference, see:
<https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/microsoft-access-sql-reference>

For connection string details, see:
<https://www.connectionstrings.com/access/>

## RELATED LINKS

{{ Fill in the related links here }}

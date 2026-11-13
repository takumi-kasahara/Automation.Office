---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Invoke-ExcelQuery
---

# Invoke-ExcelQuery

## SYNOPSIS

Executes a SQL statement against Excel workbook files and returns rows or affected-row metadata.

## SYNTAX

### TablePathSet (Default)

```
Invoke-ExcelQuery [-Path] <string[]> [-Table] <string[]> [-Address <string[]>] [-Columns <string[]>]
 [-NoHeader] [-NoIMEX] [-ReadOnly]
```

### QueryPathSet

```
Invoke-ExcelQuery [-Path] <string[]> [-Query] <string> [-NoHeader] [-NoIMEX] [-ReadOnly]
```

### QueryLiteralPathSet

```
Invoke-ExcelQuery [-Query] <string> -LiteralPath <string[]> [-NoHeader] [-NoIMEX] [-ReadOnly]
```

### TableLiteralPathSet

```
Invoke-ExcelQuery [-Table] <string[]> -LiteralPath <string[]> [-Address <string[]>]
 [-Columns <string[]>] [-NoHeader] [-NoIMEX] [-ReadOnly]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more Excel workbook files (.xlsx, .xls, .xlsm, .xlsb) through OLE DB/ODBC and executes the specified SQL statement.
For `SELECT` statements, rows are returned as PSCustomObject with Path, ObjectType, ObjectName, and column properties.
For `INSERT` and `UPDATE` statements, a single PSCustomObject with Path, ObjectType, ObjectName, and RecordsAffected is returned.
`DELETE` statements are not supported by the Excel OLE DB provider.

## EXAMPLES

### EXAMPLE 1

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Table 'Employees$'
```

Gets all rows from the Employees$ table in the specified Excel workbook.

### EXAMPLE 2

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Table 'Employees$' -Columns Id, Name
```

Gets only the Id and Name columns from the Employees$ table.

### EXAMPLE 3

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Query 'SELECT [Id], [Name] FROM [Employees$] WHERE [Id] = 1'
```

Executes a custom SQL query against the workbook.

### EXAMPLE 4

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Query "INSERT INTO [Employees$] (Id, Name, Department) VALUES (3, 'Carol', 'Marketing')"
```

Inserts a new row and returns the number of affected rows.

### EXAMPLE 5

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Query "UPDATE [Employees$] SET [Department] = 'Marketing' WHERE [Id] = 1"
```

Updates existing rows and returns the number of affected rows.

### EXAMPLE 6

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Table 'Employees$' -Address 'A1:C2'
```

Queries a specific cell range (A1:C2) within the Employees$ table.

### EXAMPLE 7

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Table 'Employees$' -NoHeader
```

Queries the table treating the first row as data (no header row), resulting in column names F1, F2, F3.

### EXAMPLE 8

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsx" -Table 'Employees$' -ReadOnly
```

Queries the table in read-only mode to prevent any write operations.

### EXAMPLE 9

```powershell

Invoke-ExcelQuery -LiteralPath "$env:TEMP\Workbook.xlsm" -Table 'Employees$'
```

Queries a macro-enabled workbook (.xlsm format).

## PARAMETERS

### -Address

Specifies one or more cell ranges to query within the corresponding table.
Use A1-style notation such as A1:B10.
The range is appended to the table name in the form [Sheet1$A1:B10].
This parameter must have the same number of elements as -Table.
Reference: [Import from Excel or Export to Excel with SQL Server Integration Services (SSIS)](https://learn.microsoft.com/en-us/sql/integration-services/load-data-to-from-excel-with-ssis?view=sql-server-ver17)

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
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

### -Columns

Specifies the columns to return when querying tables by name.
When omitted, all columns are returned.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
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

Specifies Excel workbook file paths literally.
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

### -NoHeader

Specifies that the first row of the Excel range does not contain column names.
When this switch is specified, the OLE DB connection uses `HDR=NO` instead of the default `HDR=YES`.
Column names appear as F1, F2, F3, and so on.
Reference: [Initializing the Microsoft Excel driver](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/initializing-the-microsoft-excel-driver)

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: QueryLiteralPathSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: QueryPathSet
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

### -NoIMEX

Disables IMEX mode for the OLE DB connection.
When this switch is specified, the OLE DB connection uses `IMEX=0` instead of the default `IMEX=1`.
With IMEX disabled, the driver may return null for cells whose data type does not match the guessed column type.
DML statements automatically use `IMEX=0` because the Excel provider requires an updateable connection.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: QueryLiteralPathSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: QueryPathSet
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

### -Path

Specifies Excel workbook file paths.
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
`SELECT`, `INSERT`, and `UPDATE` are supported.
`DELETE` is not supported by the Excel OLE DB provider.

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

### -ReadOnly

Opens the connection in read-only mode.
When this switch is specified, the OLE DB/ODBC connection uses read-only mode to prevent any write operations.
This switch is ignored when executing DML statements.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: QueryLiteralPathSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
  - Name: QueryPathSet
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

SQL syntax is based on the Jet/ACE SQL dialect used by the Microsoft Access database engine.
For a complete SQL reference, see:
<https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/microsoft-access-sql-reference>

Excel does not support `DELETE` statements.
`UPDATE` requires an updateable connection (`IMEX=0`).

DDL statements (`CREATE TABLE`, `DROP TABLE`, `ALTER TABLE`) are not supported.
To create or delete tables, use the Excel COM object or the OpenXML SDK.

For connection string details, see:
<https://www.connectionstrings.com/excel/>

## RELATED LINKS

{{ Fill in the related links here }}

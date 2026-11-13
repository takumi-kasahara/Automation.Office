---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Import-AccessDatabase
---

# Import-AccessDatabase

## SYNOPSIS

Imports data from a text file or spreadsheet into an Access database table.

## SYNTAX

### TextSet (Default)

```
Import-AccessDatabase [-Path] <string> -Source <string> -TableName <string>
 [-Password <securestring>] [-TransferType <AcTextTransferType>] [-CodePage <int>] [-WhatIf]
 [-Confirm]
```

### SpreadsheetSet

```
Import-AccessDatabase [-Path] <string> -Source <string> -TableName <string>
 [-Password <securestring>] [-SpreadsheetType <AcSpreadSheetType>] [-Range <string>] [-WhatIf]
 [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Imports data from a text file (CSV, delimited, HTML) or spreadsheet (Excel) into an Access database table by automating Access through COM.
The cmdlet supports two parameter sets:

- **TextSet** (default): Imports from text formats using `DoCmd.TransferText`.
- **SpreadsheetSet**: Imports from Excel formats using `DoCmd.TransferSpreadsheet`.
  If the destination database is read-only, the cmdlet throws an error.

## EXAMPLES

### EXAMPLE 1

```powershell

Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.csv" -TableName 'Employees'
```

Imports data from a CSV file into the Employees table.

### EXAMPLE 2

```powershell

Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.csv" -TableName 'Employees' -HasFieldNames
```

Imports data from a CSV file with field names in the first row.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.csv" -TableName 'Employees' -Password $password
```

Imports data into a password-protected database.

### EXAMPLE 4

```powershell

Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.xlsx" -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml
```

Imports data from an Excel spreadsheet into the Employees table.

### EXAMPLE 5

```powershell

Import-AccessDatabase -Path "$env:TEMP\Database.accdb" -Source "$env:TEMP\Employees.xlsx" -TableName 'Employees' -SpreadsheetType acSpreadsheetTypeExcel12Xml -Range 'A1:C10'
```

Imports data from a specific range in an Excel spreadsheet.

## PARAMETERS

### -CodePage

Specifies the code page to use for the imported text file.
This parameter is only valid with the **TextSet** parameter set.
The default value is `1200` (Unicode).

```yaml
Type: System.Int32
DefaultValue: 1200
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: TextSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

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

Specifies the path to the Access database file to import into.
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

### -Range

Specifies the range of cells to import from the spreadsheet.
This parameter is only valid with the **SpreadsheetSet** parameter set.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: SpreadsheetSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Source

Specifies the path to the source file to import from.
This parameter does not support wildcards because it represents an existing file path.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -SpreadsheetType

Specifies the spreadsheet type for the import.
This parameter is only valid with the **SpreadsheetSet** parameter set.
See <https://learn.microsoft.com/en-us/office/vba/api/access.acspreadsheettype> for valid values.

```yaml
Type: Microsoft.Office.Interop.Access.AcSpreadSheetType
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: SpreadsheetSet
    Position: Named
    IsRequired: false
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -TableName

Specifies the name of the Access table to import data into.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: Named
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -TransferType

Specifies the text transfer type for the import.
This parameter is only valid with the **TextSet** parameter set.
The default value is `acImportDelim`.
See <https://learn.microsoft.com/en-us/office/vba/api/access.actexttransfertype> for valid values.

```yaml
Type: Microsoft.Office.Interop.Access.AcTextTransferType
DefaultValue: acImportDelim
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: TextSet
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

### System.Void

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

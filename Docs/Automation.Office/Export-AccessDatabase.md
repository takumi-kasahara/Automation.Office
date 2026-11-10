---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Export-AccessDatabase
---

# Export-AccessDatabase

## SYNOPSIS

Exports data from an Access database table to a text file or spreadsheet.

## SYNTAX

### TextSet (Default)

```
Export-AccessDatabase [-Path] <string> -TableName <string> -Destination <string>
 [-Password <securestring>] [-Force] [-NoClobber] [-TransferType <AcTextTransferType>]
 [-CodePage <int>] [-WhatIf] [-Confirm]
```

### SpreadsheetSet

```
Export-AccessDatabase [-Path] <string> -TableName <string> -Destination <string>
 [-Password <securestring>] [-Force] [-NoClobber] [-SpreadsheetType <AcSpreadSheetType>]
 [-Range <string>] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Exports the contents of a table from an Access database to a text file (CSV, delimited, HTML) or spreadsheet (Excel) by automating Access through COM.
The cmdlet supports two parameter sets:

- **TextSet** (default): Exports to text formats using `DoCmd.TransferText`.
- **SpreadsheetSet**: Exports to Excel formats using `DoCmd.TransferSpreadsheet`.
  If the destination file already exists, the cmdlet stops unless `-Force` is specified.
  If `-NoClobber` is specified, the cmdlet throws an error when the destination exists.

## EXAMPLES

### EXAMPLE 1

```powershell

Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.txt"
```

Exports the Employees table to a delimited text file.

### EXAMPLE 2

```powershell

Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.txt" -HasFieldNames
```

Exports the Employees table to a text file with field names as the first row.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.txt" -Password $password
```

Exports data from a password-protected database.

### EXAMPLE 4

```powershell

Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.html" -TransferType acExportHTML
```

Exports the Employees table to an HTML file.

### EXAMPLE 5

```powershell

Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.xlsx" -SpreadsheetType acSpreadsheetTypeExcel12Xml
```

Exports the Employees table to an Excel spreadsheet.

### EXAMPLE 6

```powershell

Export-AccessDatabase -Path "$env:TEMP\Database.accdb" -TableName 'Employees' -Destination "$env:TEMP\Employees.xlsx" -SpreadsheetType acSpreadsheetTypeExcel12Xml -Range 'A1:C10'
```

Exports the Employees table to an Excel spreadsheet with a specific range.

## PARAMETERS

### -CodePage

Specifies the code page to use for the exported text file.
This parameter is only valid with the **TextSet** parameter set.
The default value is `1200` (Unicode).
See <https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers> for valid code page values.

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

### -Destination

Specifies the destination file path for the exported data.
This parameter does not support wildcards because it represents a new file path.

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

### -Force

Overwrites an existing file at `-Destination`.
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

### -NoClobber

Prevents overwriting an existing file at `-Destination`.
If the destination file exists, the cmdlet throws an error.
This switch takes precedence over `-Force`.

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

Specifies the path to the Access database file to export from.
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

Specifies the range of cells to export in the spreadsheet.
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

### -SpreadsheetType

Specifies the spreadsheet type for the export.
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

Specifies the name of the table to export.

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

Specifies the text transfer type for the export.
This parameter is only valid with the **TextSet** parameter set.
The default value is `acExportDelim`.
See <https://learn.microsoft.com/en-us/office/vba/api/access.actexttransfertype> for valid values.

```yaml
Type: Microsoft.Office.Interop.Access.AcTextTransferType
DefaultValue: acExportDelim
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

### System.IO.FileInfo

Returns a FileInfo object representing the exported file.

{{ Fill in the Description }}

### System.IO.FileInfo

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

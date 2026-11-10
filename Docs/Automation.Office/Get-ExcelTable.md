---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-ExcelTable
---

# Get-ExcelTable

## SYNOPSIS

Gets tables in Excel workbook files.

## SYNTAX

### PathSet (Default)

```
Get-ExcelTable [-Path] <string[]> [-Password <securestring>] [-ReadOnly]
```

### LiteralPathSet

```
Get-ExcelTable -LiteralPath <string[]> [-Password <securestring>] [-ReadOnly]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more Excel workbook files (.xlsx, .xls, .xlsm, .xlsb) through OLE DB/ODBC and returns table entries.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-ExcelTable -Path "$env:TEMP\Workbook.xlsx"
```

Gets all tables (worksheets) in the specified Excel workbook file.

### EXAMPLE 2

```powershell

Get-ExcelTable -LiteralPath "$env:TEMP\Workbook.xlsx"
```

Gets all tables using a literal path (wildcards not interpreted).

### EXAMPLE 3

```powershell

"$env:TEMP\Workbook.xlsx" | Get-ExcelTable
```

Gets all tables by piping the file path to the cmdlet.

### EXAMPLE 4

```powershell

Get-ExcelTable -Path "$env:TEMP\*.xlsx"
```

Gets all tables from all .xlsx files in the specified directory using wildcards.

### EXAMPLE 5

```powershell

Get-ExcelTable -LiteralPath "$env:TEMP\Workbook.xls"
```

Gets all tables from an .xls format workbook file.

### EXAMPLE 6

```powershell

Get-ExcelTable -LiteralPath '$env:TEMP\Workbook.xlsm' -ReadOnly
```

Gets all tables from a macro-enabled workbook in read-only mode.

## PARAMETERS

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

{{ Fill Password Description }}

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

Specifies Excel workbook file paths.
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

### -ReadOnly

Opens the connection in read-only mode.
When this switch is specified, the OLE DB/ODBC connection uses read-only mode to prevent any write operations.

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

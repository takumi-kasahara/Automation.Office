---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-ExcelTableColumn
---

# Get-ExcelTableColumn

## SYNOPSIS

Gets table column metadata in Excel workbook files.

## SYNTAX

### PathSet (Default)

```
Get-ExcelTableColumn [-Path] <string[]> [-Table] <string[]> [-Password <securestring>] [-ReadOnly]
```

### LiteralPathSet

```
Get-ExcelTableColumn [-Table] <string[]> -LiteralPath <string[]> [-Password <securestring>]
 [-ReadOnly]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more Excel workbook files (.xlsx, .xls, .xlsm, .xlsb) through OLE DB/ODBC and returns column metadata for specified tables.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-ExcelTableColumn -Path "$env:TEMP\Workbook.xlsx" -Table 'Employees$'
```

Gets column metadata for the Employees$ table in the specified Excel workbook file.

### EXAMPLE 2

```powershell

Get-ExcelTableColumn -LiteralPath "$env:TEMP\Workbook.xlsx" -Table 'Employees$'
```

Gets column metadata using a literal path (wildcards not interpreted).

### EXAMPLE 3

```powershell

"$env:TEMP\Workbook.xlsx" | Get-ExcelTableColumn -Table 'Employees$'
```

Gets column metadata by piping the file path to the cmdlet.

### EXAMPLE 4

```powershell

Get-ExcelTableColumn -Path "$env:TEMP\*.xlsx" -Table 'Employees$'
```

Gets column metadata from all .xlsx files in the specified directory using wildcards.

### EXAMPLE 5

```powershell

Get-ExcelTableColumn -LiteralPath "$env:TEMP\Workbook.xls" -Table 'Employees$'
```

Gets column metadata from an .xls format workbook file.

### EXAMPLE 6

```powershell

Get-ExcelTableColumn -LiteralPath "$env:TEMP\Workbook.xlsm" -Table 'Employees$' -ReadOnly
```

Gets column metadata from a macro-enabled workbook in read-only mode.

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

### -Table

Specifies one or more table names whose column metadata is returned.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
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

## RELATED LINKS

{{ Fill in the related links here }}

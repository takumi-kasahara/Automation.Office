---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Open-ExcelFile
---

# Open-ExcelFile

## SYNOPSIS

Opens an Excel workbook file.

## SYNTAX

### __AllParameterSets

```
Open-ExcelFile [-Path] <string> [-Application <__ComObject>] [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-ReadOnly] [-Force] [-Action <scriptblock>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens a workbook file by automating Excel through COM.

## EXAMPLES

### EXAMPLE 1

```powershell

Open-ExcelFile -Path "$env:TEMP\Book.xlsx"
```

Opens a workbook and returns the workbook object.

### EXAMPLE 2

```powershell

Open-ExcelFile -Path "$env:TEMP\Book.xlsx" -Action {
  param ($Workbook)
  $Workbook.Worksheets.Item(1).Name = 'Sheet1'
}
```

Opens a workbook, renames the first worksheet, and closes the workbook.

### EXAMPLE 3

```powershell

$app = New-ExcelObject
Open-ExcelFile -Application $app -Path "$env:TEMP\Book.xlsx" -Action {
  param ($Workbook)
  $Workbook.Worksheets.Item(1).Name = 'Sheet1'
}
$app.Quit()
```

Opens a workbook using an existing Excel Application object.

## PARAMETERS

### -Action

Specifies a script block to execute with the opened workbook.
The script block receives the workbook object as its first argument.
When this parameter is specified, the workbook is automatically closed after the action completes.

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

Specifies the Excel Application COM object to use for opening the workbook.
If not specified, a new Excel Application object is created automatically.

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

### -Force

Ignores the read-only recommended flag when opening the workbook.

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

### -PasswordToModify

Specifies the password required to modify the workbook.

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

### -PasswordToOpen

Specifies the password required to open the workbook.

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

Specifies the path to the workbook file to open.
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

### -ReadOnly

Opens the workbook in read-only mode.

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

### System.String

{{ Fill in the Description }}

## OUTPUTS

### System.__ComObject

Returns the opened workbook object.

{{ Fill in the Description }}

### System.__ComObject

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

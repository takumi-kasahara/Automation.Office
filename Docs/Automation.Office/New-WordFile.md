---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: New-WordFile
---

# New-WordFile

## SYNOPSIS

Creates a new Word document file.

## SYNTAX

### __AllParameterSets

```
New-WordFile [-Path] <string> [-FileFormat <WdSaveFormat>] [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-ReadOnlyRecommended] [-Force] [-RemovePersonalInformation]
 [-Initialize <scriptblock>] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Creates a file at the specified path by automating Word through COM.
If the destination file already exists, the command stops unless `-Force` is specified.

## EXAMPLES

### EXAMPLE 1

```powershell

New-WordFile -Path "$env:TEMP\Document.docx"
```

Creates a new document.

### EXAMPLE 2

```powershell

New-WordFile -Path "$env:TEMP\Document.docm" -FileFormat wdFormatXMLDocumentMacroEnabled -Force
```

Creates or overwrites a macro-enabled document.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
New-WordFile -Path "$env:TEMP\Document.docx" -PasswordToOpen $password
```

Creates a file protected with an open password.

### EXAMPLE 4

```powershell

New-WordFile -Path "$env:TEMP\Document.docx" -ReadOnlyRecommended
```

Creates a file that recommends opening as read-only.

### EXAMPLE 5

```powershell

New-WordFile -Path "$env:TEMP\Document.docx" -Initialize {
  param ($Document)
  $Document.Range().Text = 'Hello, World!'
}
```

Creates a document and adds text to it.

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

Specifies the Word file format used when saving the document.
The default value is `wdFormatXMLDocument`.

```yaml
Type: Microsoft.Office.Interop.Word.WdSaveFormat
DefaultValue: wdFormatXMLDocument
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

### -Initialize

Specifies a script block to initialize the document before saving.
The script block receives the document object as its first argument.

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

### -PasswordToModify

Specifies the password required to modify the document.

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

Specifies the password required to open the document.

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

Specifies the destination path of the document file to create.
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

### -ReadOnlyRecommended

Saves the document with the read-only recommended flag set.

When this switch is specified, Word recommends that users open the document as read-only.

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

### -RemovePersonalInformation

Removes personal information from the new document file.

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

Returns the created document.

{{ Fill in the Description }}

### System.IO.FileInfo

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

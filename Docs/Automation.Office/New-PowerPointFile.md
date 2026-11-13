---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: New-PowerPointFile
---

# New-PowerPointFile

## SYNOPSIS

Creates a new PowerPoint presentation file.

## SYNTAX

### __AllParameterSets

```
New-PowerPointFile [-Path] <string> [-FileFormat <PpSaveAsFileType>]
 [-PasswordToOpen <securestring>] [-PasswordToModify <securestring>] [-ReadOnlyRecommended] [-Force]
 [-RemovePersonalInformation] [-Initialize <scriptblock>] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Creates a file at the specified path by automating PowerPoint through COM.
If the destination file already exists, the command stops unless `-Force` is specified.

## EXAMPLES

### EXAMPLE 1

```powershell

New-PowerPointFile -Path "$env:TEMP\Presentation.pptx"
```

Creates a new presentation.

### EXAMPLE 2

```powershell

New-PowerPointFile -Path "$env:TEMP\Presentation.pptm" -FileFormat ppSaveAsOpenXMLPresentationMacroEnabled -Force
```

Creates or overwrites a macro-enabled presentation.

### EXAMPLE 3

```powershell

$password = Read-Host -AsSecureString
New-PowerPointFile -Path "$env:TEMP\Presentation.pptx" -PasswordToOpen $password
```

Creates a file protected with an open password.

### EXAMPLE 4

```powershell

New-PowerPointFile -Path "$env:TEMP\Presentation.pptx" -ReadOnlyRecommended
```

Creates a file that recommends opening as read-only.

### EXAMPLE 5

```powershell

New-PowerPointFile -Path "$env:TEMP\Presentation.pptx" -Initialize {
  param ($Presentation)
  $Presentation.Slides.Add(1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutTitleOnly) | Out-Null
}
```

Creates a presentation and adds a title slide.

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

Specifies the PowerPoint file format used when saving the presentation.
The default value is `ppSaveAsOpenXMLPresentation`.

```yaml
Type: Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType
DefaultValue: ppSaveAsOpenXMLPresentation
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

Specifies a script block to initialize the presentation before saving.
The script block receives the presentation object as its first argument.

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

Specifies the password required to modify the presentation.

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

Specifies the password required to open the presentation.

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

Specifies the destination path of the presentation file to create.
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

Saves the presentation with the read-only recommended flag set.

When this switch is specified, PowerPoint recommends that users open the presentation as read-only.

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

Removes personal information from the new presentation file.

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

Returns the created presentation.

{{ Fill in the Description }}

### System.IO.FileInfo

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

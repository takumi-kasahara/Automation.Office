---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Open-PowerPointFile
---

# Open-PowerPointFile

## SYNOPSIS

Opens a PowerPoint presentation file.

## SYNTAX

### __AllParameterSets

```
Open-PowerPointFile [-Path] <string> [-Application <__ComObject>] [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-ReadOnly] [-Action <scriptblock>]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens a presentation file by automating PowerPoint through COM.

## EXAMPLES

### EXAMPLE 1

```powershell

Open-PowerPointFile -Path "$env:TEMP\Presentation.pptx"
```

Opens a presentation and returns the presentation object.

### EXAMPLE 2

```powershell

Open-PowerPointFile -Path "$env:TEMP\Presentation.pptx" -Action {
  param ($Presentation)
  $Presentation.Slides.Add(1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutTitleOnly) | Out-Null
}
```

Opens a presentation, adds a title slide, and closes the presentation.

## PARAMETERS

### -Action

Specifies a script block to execute with the opened presentation.
The script block receives the presentation object as its first argument.
When this parameter is specified, the presentation is automatically closed after the action completes.

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

Specifies the PowerPoint Application COM object to use for opening the presentation.
If not specified, a new PowerPoint Application object is created automatically.

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

Specifies the path to the presentation file to open.
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

Opens the presentation in read-only mode.

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

Returns the opened presentation object.

{{ Fill in the Description }}

### System.__ComObject

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}

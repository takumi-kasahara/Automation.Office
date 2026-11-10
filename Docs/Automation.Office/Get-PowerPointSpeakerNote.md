---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Get-PowerPointSpeakerNote
---

# Get-PowerPointSpeakerNote

## SYNOPSIS

Gets speaker notes from one or more PowerPoint presentation files.

## SYNTAX

### PathSet (Default)

```
Get-PowerPointSpeakerNote [-Path] <string[]> [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-Range <string>] [-Force]
```

### LiteralPathSet

```
Get-PowerPointSpeakerNote -LiteralPath <string[]> [-PasswordToOpen <securestring>]
 [-PasswordToModify <securestring>] [-Range <string>] [-Force]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Opens one or more PowerPoint presentations in read-only mode and extracts speaker notes from each slide.
The cmdlet supports both `-Path` and `-LiteralPath` parameter sets and can open protected presentations by using `-PasswordToOpen` and `-PasswordToModify`.
Each output object contains the source file path and a collection of speaker notes, where each note includes the slide number and the corresponding text.

## EXAMPLES

### EXAMPLE 1

```powershell

Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx"
```

Gets speaker notes from the specified presentation.

### EXAMPLE 2

```powershell

Get-PowerPointSpeakerNote -LiteralPath "$env:TEMP\Presentation.pptx"
```

Gets speaker notes from the specified presentation by using a literal path.

### EXAMPLE 3

```powershell

$note = Get-PowerPointSpeakerNote -LiteralPath "$env:TEMP\Presentation.pptx"
$note.Items |
ForEach-Object { $_.Text -split "`r" } |
Where-Object { $_.Trim() -ne [string]::Empty } |
Out-File -LiteralPath "$([IO.Path]::GetFileName($note.Path)).txt" -Append -Encoding utf8
```

Exports each speaker note to a separate text file named after the source presentation and slide number.

### EXAMPLE 4

```powershell

$password = Read-Host -AsSecureString
Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx" -PasswordToOpen $password
```

Gets speaker notes from a presentation protected with an open password.

### EXAMPLE 5

```powershell

Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx" -Force
```

Gets speaker notes including notes on hidden slides.

### EXAMPLE 6

```powershell

Get-PowerPointSpeakerNote -Path "$env:TEMP\Presentation.pptx" -Range '2, 4-6'
```

Gets speaker notes only for slides 2, 4, 5, and 6.

## PARAMETERS

### -Force

Includes speaker notes from hidden slides.
If omitted, speaker notes from hidden slides are not included in the output.

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

### -LiteralPath

Specifies presentation paths literally.
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

Specifies presentation paths.
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

### -Range

Specifies slide indexes to include.

Supports discrete indexes and inclusive ranges separated by commas.
Spaces are allowed.
For example, `1-3, 5, 7-10` includes slides 1, 2, 3, 5, 7, 8, 9, and 10.

```yaml
Type: System.String
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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String[]

{{ Fill in the Description }}

## OUTPUTS

### SpeakerNotes

Returns an object that contains the source path and a collection of `SpeakerNote` objects.

{{ Fill in the Description }}

### SpeakerNotes

{{ Fill in the Description }}

## NOTES

The presentation is always opened in read-only mode and the PowerPoint COM object is released after processing.

## RELATED LINKS

{{ Fill in the related links here }}

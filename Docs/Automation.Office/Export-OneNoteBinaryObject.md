---
document type: cmdlet
external help file: Automation.Office-Help.xml
HelpUri: ''
Locale: ja-JP
Module Name: Automation.Office
ms.date: 11/09/2026
PlatyPS schema version: 2024-05-01
title: Export-OneNoteBinaryObject
---

# Export-OneNoteBinaryObject

## SYNOPSIS

Exports binary objects embedded in a OneNote page.

## SYNTAX

### __AllParameterSets

```
Export-OneNoteBinaryObject [-Id] <string> [-Destination] <string> [-Force] [-NoClobber] [-UseName]
 [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
{{Insert list of aliases}}

## DESCRIPTION

Retrieves all binary objects (images, attachments) from the specified OneNote page and saves each one as a file in a subdirectory under the destination directory.
The subdirectory is named `<Id>` by default, or `<name>` when -UseName is specified.
Each file is named by its callback ID, with the extension determined by the file signature.
File timestamps (CreationTime, LastWriteTime) are set from the parent OE element when available.
If the page contains no binary objects, no files are created and nothing is returned.

## EXAMPLES

### EXAMPLE 1

```powershell

$pageId = New-OneNotePage -SectionId $sectionId
Export-OneNoteBinaryObject -Id $pageId -Destination $env:TEMP
```

Exports all binary objects from the page to $env:TEMP\<pageId>\.

### EXAMPLE 2

```powershell

Export-OneNoteBinaryObject -Id $pageId -Destination $env:TEMP -UseName
```

Exports binary objects to a subdirectory named after the page's display name.

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

### -Destination

The parent directory path where the binary object subdirectory will be created.
The directory will be created if it does not exist.

```yaml
Type: System.String
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

### -Force

Overwrites read-only output files when present.

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

### -Id

The OneNote object ID of the page from which binary objects will be exported.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
  - Name: (All)
    Position: 0
    IsRequired: true
    ValueFromPipeline: false
    ValueFromPipelineByPropertyName: false
    ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -NoClobber

Throws an error if an output file already exists.

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

### -UseName

When specified, the subdirectory is named using the page's display name
instead of its object ID.

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

## OUTPUTS

### System.IO.FileInfo

One FileInfo object per exported binary object file.
Nothing is returned if the page contains no binary objects.

{{ Fill in the Description }}

### System.IO.FileInfo

{{ Fill in the Description }}

## NOTES

This function uses the OneNote COM API (GetBinaryPageContent) internally.
Supported file signatures: PNG, JPEG, GIF, BMP, TIFF.
Files without a recognized signature are saved with a .bin extension.
COM objects are released after execution via garbage collection.

## RELATED LINKS

{{ Fill in the related links here }}

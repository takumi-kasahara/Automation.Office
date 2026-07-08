---
name: powershell-cmdlet
description: 'This skill creates or updates PowerShell cmdlets for this repository. Use when: create commandlet, update commandlet, add parameter, comment-based help.'
argument-hint: 'Enter the target cmdlet and the behavior you want to add or change.'
user-invocable: true
---

# How to Edit PowerShell Cmdlets

## Guidelines

Use this checklist by section:

### Cmdlet Pattern

1. Pick the correct cmdlet pattern (`New-*`, `Get-*`, `Set-*`, `Remove-*`, `Export-*`, `Import-*`).

### Parameters and Behavior

2. Apply parameter and attribute rules from the parameter section.
3. If `SupportsShouldProcess` is enabled, verify `-WhatIf` and `-Confirm` behavior.

### Error Handling

4. Add clear, user-friendly terminating errors for invalid or unsupported parameter values, including valid alternatives.

- Follow parameter naming and attribute usage conventions from the official PowerShell docs: https://learn.microsoft.com/powershell/.
- If using `[CmdletBinding(SupportsShouldProcess)]`, always implement `-WhatIf`/`-Confirm` behavior.
- Refer to existing similar functions and tests.
- If a parameter value is invalid or unsupported, throw a clear terminating error that states the issue and valid alternatives.

### Parameter Guidelines

- When adding `[CmdletBinding(SupportsShouldProcess)]`, follow official documentation
  - Always call `$PSCmdlet.ShouldProcess()`
  - Add `-WhatIf:$WhatIfPreference` to any command that supports `-WhatIf`

#### `New-*`

- `[CmdletBinding(SupportsShouldProcess)]`
  - `[switch]$Force`
    - Overwrite existing items when present
    - If the target is a folder or another mismatched type when creating a file, throw an exception
- `[string]$Path`
  - `[Alias('FullName')]`
  - `[Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]`
  - `[ValidateScript({ Test-Path -LiteralPath $_ -IsValid })]`

#### `Get-*`

- `[string[]]$Path`
  - `[Parameter(Mandatory, Position = 0)]`
  - `[SupportsWildcards()]`
  - `[ValidateScript({ Test-Path -Path $_ })]`
- `[string[]]$LiteralPath`
  - `[Alias('PSPath', 'LP')]`
  - `[Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]`
  - `[ValidateScript({ Test-Path -LiteralPath $_ })]`

#### `Set-*` / `Remove-*`

- `[CmdletBinding(SupportsShouldProcess)]`
  - `[switch]$Force`: ignore read-only attributes and overwrite
- `[string[]]$Path`
  - `[Parameter(Mandatory, ParameterSetName = 'PathSet', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]`
  - `[SupportsWildcards()]`
  - `[ValidateScript({ Test-Path -Path $_ })]`
- `[string[]]$LiteralPath`
  - `[Alias('PSPath', 'LP')]`
  - `[Parameter(Mandatory, ParameterSetName = 'LiteralPathSet', ValueFromPipelineByPropertyName)]`
  - `[ValidateScript({ Test-Path -LiteralPath $_ })]`

#### `Export-*`

- `[CmdletBinding(SupportsShouldProcess)]`
  - `[switch]$Force`: ignore read-only attributes and overwrite
  - `[switch]$NoClobber`: error if the file already exists
    - If not specified, overwrite
- `[string]$Path`
  - `[Alias('FilePath', 'FullName')]`
  - `[Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]`
  - `[ValidateScript({ Test-Path -LiteralPath $_ })]`

#### `Import-*`

- `[CmdletBinding(SupportsShouldProcess)]`
  - `[switch]$Force`: ignore read-only attributes and overwrite
- `[string]$Path`
  - `[Alias('FilePath', 'FullName')]`
  - `[Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]`
  - `[ValidateScript({ Test-Path -LiteralPath $_ })]`

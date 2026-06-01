---
name: powershell-instructions
description: This file describes the workspace guidelines for editing PowerShell scripts and modules.
applyTo: '**/*.ps1, **/*.psm1'
---

# PowerShell Coding Guidelines

When editing PowerShell scripts and modules in this workspace, follow the following guidelines:

## Critical Rules

- For edit commandlet, use [powershell-cmdlet/SKILL.md](../skills/powershell-cmdlet/SKILL.md).
- For edit comment-based help, use [comment-based-help/SKILL.md](../skills/comment-based-help/SKILL.md).
- Use `[CmdletBinding()]` for every function.
- Use `Set-StrictMode -Version Latest` in every script.

## Best Practices

### Error Handling

- Do not use `Write-Error` for error handling. It writes to the error stream but does not stop execution.
- Use `$PSCmdlet.ThrowTerminatingError()` for error handling.

### Null Handling and Type Conversion Caveats

- Use `$null -eq $value` or `$null -ne $value` for null comparisons to avoid potential parsing issues.
- For null checks, use `[string]::IsNullOrEmpty($value)` for strings.

#### References

- [Everything you wanted to know about $null](https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-null?view=powershell-5.1)

### Constructor

For .NET object construction, prefer using the `new` method of the class over `New-Object` cmdlet.

- Example:
  - OK: `[System.IO.FileInfo]::new('C:\folder\file.txt')`
  - NG: `New-Object System.IO.FileInfo 'C:\folder\file.txt'`

#### References

- [about_Object_Creation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_object_creation?view=powershell-5.1)

### String Formatting

For simple string formatting, prefer string interpolation (expansion) syntax.
Do not use `+` for string concatenation.

- Example:
  - OK: `"value: $value"`
  - OK: `"value: {0:0.0}" -f $value`
  - NG: `"value: {0}" -f $value`
  - NG: `"value: " + $value`

### Path Resolution

- For paths that may not exist at runtime, prefer `System.IO.Path` methods.
- `Convert-Path` and `Resolve-Path` throw exceptions when the target path does not exist.
- Use these cmdlets only when the path is guaranteed to exist.

#### Examples

  - OK: `if (Test-Path -LiteralPath $existingPath) { Convert-Path -Path $existingPath }`
  - OK: `$item = Get-Item -LiteralPath $existingPath; Resolve-Path -LiteralPath $item.FullName`
  - OK: `[Path]::GetFullPath($mayNotExistPath)`
  - NG: `Convert-Path -Path $mayNotExistPath`
  - NG: `Resolve-Path -Path $mayNotExistPath`

#### References

  - [Convert-Path](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/convert-path?view=powershell-5.1)
  - [Resolve-Path](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/resolve-path?view=powershell-5.1)

### Path Joining

- When joining paths, prefer passing the path through the pipeline into `Join-Path -ChildPath`.
- This style is especially useful when concatenating multiple segments or chaining path construction.

#### Examples

    - OK: `$PSScriptRoot | Join-Path -ChildPath 'src' | Join-Path -ChildPath 'module.psm1'`
    - NG: `Join-Path $PSScriptRoot 'src\module.psm1'`
    - NG: `Join-Path -Path $PSScriptRoot -ChildPath 'src\module.psm1'`

### Wildcard Handling

- `Split-Path` accepts wildcard patterns through `-Path`, and passing an uncertain input via the pipeline can cause unexpected behavior.
- Escape path values with `[WildcardPattern]::Escape(String)` before calling `Split-Path`.

#### Examples

- OK: `[WildcardPattern]::Escape($path) | Split-Path -Extension`
- NG: `Split-Path -Path $path -Extension`

#### References

- [Split-Path](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/split-path?view=powershell-5.1)
- [WildcardPattern.Escape(String) Method](https://learn.microsoft.com/en-us/dotnet/api/system.management.automation.wildcardpattern.escape?view=powershellsdk-1.1.0)

#### References

- [Join-Path](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/join-path?view=powershell-5.1)

## References

- [Documentation](https://learn.microsoft.com/en-us/powershell/)
- [PowerShell Module Browser](https://learn.microsoft.com/en-us/powershell/module/)
- [PowerShell Gallery](https://www.powershellgallery.com/)

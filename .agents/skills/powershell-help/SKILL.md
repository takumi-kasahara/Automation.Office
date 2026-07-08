---
name: powershell-help
description: 'This skill creates or updates PowerShell comment-based help for this repository. Use when documenting a function, updating help, or aligning help with implementation.'
argument-hint: 'Provide: (1) target function or cmdlet name, and (2) help content to add or update. If either is missing, ask for the missing item before editing.'
user-invocable: true
---

# PowerShell Comment-Based Help

## Guidelines

Use these steps in order:

1. Core sections: update `.SYNOPSIS` and `.DESCRIPTION`.
2. Parameters: update each `.PARAMETER <Name>` in function-signature order.
3. Supporting sections: add or refresh `.EXAMPLE`, `.OUTPUTS`, and `.NOTES`.
4. Validation: verify with `Get-Help [-Path <Path>] -Name <Name> -Full`.

- Follow the official PowerShell comment-based help documentation: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_comment_based_help.
- `.SYNOPSIS` should describe the function purpose concisely.
- `.DESCRIPTION` should describe behavior details, notes, and overall `ShouldProcess` behavior.
- `.PARAMETER <Name>` should match the parameter definition in the function signature, including order, type, and required/optional behavior.
- `.EXAMPLE` should include usage examples, scenarios, and results.
- `.OUTPUTS` should include the output type FQCN (Fully Qualified Class Name) and meaning.
  - Example:
    - `[OutputType([void])]`: `.OUTPUTS None.`
    - `[OutputType([string])]`: `.OUTPUTS System.String`
    - `[OutputType([string[]])]`: `.OUTPUTS System.String[]`
    - `[OutputType([System.IO.FileInfo])]`: `.OUTPUTS System.IO.FileInfo`
- `.NOTES` should include supplementary information or caveats.
- After editing, always verify output with `Get-Help [-Path <Path>] -Name <Name> -Full`.

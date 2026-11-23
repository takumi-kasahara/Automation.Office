# Project Guidelines

This repository is a collection of PowerShell modules for automating Microsoft Office applications (Access / Excel / PowerPoint / Word / OneNote).

## Overview

- [Apps/](./Apps/) contains VBA source components (`.bas`, `.cls`, `.frm`, `.vba`) extracted from Office files.
- [Bin/](./Bin/) contains generated Office files produced by compile workflows.
- [Modules/](./Modules/) contains PowerShell module manifests and implementation.
- [Modules/*.psd1](./Modules/) exports commands from [Modules/Functions/*.psm1](./Modules/Functions/).
- [Modules/Functions/](./Modules/Functions/) contains modules and tests.
- [Scripts/](./Scripts/) contains standalone program helper scripts.
- [Tests/](./Tests/) contains Office fixtures for VBProject scenarios.

## Workflow

### PowerShell module development cycle

1. Edit files under [Modules/Functions/](./Modules/Functions/).
2. When editing `Module.ps1`, also update or create the corresponding test file `Module.Tests.ps1` and ensure tests cover your changes.
3. When add or remove cmdlets, also update the module manifest files [Modules/*.psd1](./Modules/) to reflect the changes.

#### PSScriptAnalyzer Invocation

Prefer the `psscriptanalyzer` task defined in [tasks.json](./.vscode/tasks.json) when using VS Code.

Run PSScriptAnalyzer with [analyze.cmd](./.agents/skills/powershell-5/scripts/analyze.cmd) only when invoking directly outside of VS Code tasks:

```bat
".agents\skills\powershell-5\scripts\analyze.cmd" -Path "Modules" -Recurse
```

```powershell
& '.\.agents\skills\powershell-5\scripts\analyze.cmd' -Path '.\Modules\' -Recurse
```

#### Pester Invocation

Prefer the `pester` task defined in [tasks.json](./.vscode/tasks.json) when using VS Code.

Run Pester with [pester.cmd](./.agents/skills/powershell-5/scripts/pester.cmd) only when invoking directly outside of VS Code tasks:

```bat
".agents\skills\powershell-5\scripts\pester.cmd" -Path "Modules\Functions" -Parallel
```

```powershell
& '.\.agents\skills\powershell-5\scripts\pester.cmd' -Path '.\Modules\Functions\' -Parallel
```

### VBA development cycle

1. Edit files under [Apps/](./Apps/).
2. Run `Office Macros: compile` task.

### Office Scripts development cycle

1. Edit the TypeScript source file under [Office Scripts](./Office%20Scripts/).
2. Run `Office Scripts: compile` task.

## Coding Conventions

- Avoid breaking changes unless necessary, and always seek review for such changes.
- Keep changes minimal and consistent with existing file patterns.
- Keep functions and modules small and single-responsibility.
- Prefer explicit error handling and fail-fast behavior.
- Prefer self-documenting code; use comments for intent, not for restating code.

## Command Execution Guidelines

- Do not run one-line PowerShell commands directly for auditability and reproducibility.
- Write the command content to `./.temp/execute.ps1` first.
- Run PowerShell scripts only with:
  - `powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "./.temp/execute.ps1"`

# Project Guidelines

This repository is a collection of PowerShell modules for automating Microsoft Office applications (Access / Excel / PowerPoint / Word / OneNote).

## Overview

- `Apps/` contains VBA source components (`.bas`, `.cls`, `.frm`, `.vba`) extracted from Office files.
- `Bin/` contains generated Office files produced by compile workflows.
- `Modules/` contains PowerShell module manifests and implementation.
- `Modules/*.psd1` exports commands from `Modules/Sources/*.psm1`.
- `Modules/Sources/` contains modules and tests.
- `Scripts/` contains standalone program helper scripts.
- `Tests/` contains Office fixtures for VBProject scenarios.

## Workflow

### PowerShell module development cycle

1. Edit files under `Modules/Sources/`.
2. When editing `Modules/Sources/Module.psm1`, also update or create the corresponding test file `Modules/Tests/Module.Tests.ps1` and ensure tests cover your changes.
3. When add or remove cmdlets, also update the module manifest files (`Modules/*.psd1`) to reflect the changes.

### VBA development cycle

1. Edit files under `Apps/`.
2. Run `Compile.ps1`.

### Office Scripts development cycle

1. Edit the TypeScript source file under `Office Scripts/` (e.g., `Script.ts`).
2. Run the compile script: `powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "./Compile.OfficeScripts.ps1"`.
   - This converts `.ts` to `.osts` format using the JSON schema defined in `Script.osts`.
3. To decompile, run: `powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "./Decompile.OfficeScripts.ps1"`.
   - This extracts the TypeScript source from `.osts` back to `.ts`.

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
  - `powershell.exe -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "./.temp/execute.ps1"`

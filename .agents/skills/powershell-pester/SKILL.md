---
name: powershell-pester
description: 'This skill creates or updates Pester tests. Use to create *.Tests.ps1. Follow the context checklist in order.'
argument-hint: 'Enter the target module or test file and the test cases you want to validate.'
user-invocable: true
---

# How to create Pester test

## Guidelines

- Organize Pester tests with Describe, Context, and It blocks, creating one Describe block per function.
- Put helper functions in the Describe block's BeforeAll.
- Divide Context into these four categories:
  - ParameterSetName
    - Cover all ParameterSetNames.
    - If there is no ParameterSetName, cover mandatory parameters.
    - For parameters with `SupportsWildcards()`, test values containing wildcards.
  - SupportsShouldProcess
    - Apply this section only when `CmdletBinding(SupportsShouldProcess)` is present.
    - Required cases (run in order):
      1. Verify no side effects occur with `-WhatIf`.
      2. Verify the command runs when `-Confirm` is accepted.
      3. Verify the command does not run when `-Confirm` is declined.
    - Notes:
      - Since mocking `$PSCmdlet.ShouldProcess()` is difficult, review this in code review.
      - When `ConfirmImpact` is `High`, suppress the dialog by testing with `-Confirm:$false`.
    - Optional cases:
      - Verify `-Force` takes precedence over `-Confirm`.
      - Verify existing files are not overwritten with `-NoClobber`.
  - Other parameter
    - Parameters not covered by `ParameterSetName` or `SupportsShouldProcess`.
    - Add tests for defaults, aliases, and accepted input shapes.
  - Edge case
    - Add boundary and error-case tests.
    - Example: minimum/maximum numeric values, read-only target file behavior.
    - Do not test constraints defined by `Parameter` or `Validate*` attributes.
- Put helper functions in the Describe block's BeforeAll.
- When changing existing code, run only related tests and verify behavior.

## Best Practices for Pester Mock

### Pay attention to scope

- Declare Mocks inside BeforeAll or It blocks.
- A Mock declared in BeforeAll applies to all test cases in the same Describe block.
- A Mock declared in an It block applies only to that test case.

### Always mock validation and side-effect commands

- Always mock functions called by ValidationScript or operations with significant side effects, such as file operations or external program execution.
  - Example: always mock validation functions like `Test-Path`.
- Use primitive types or objects created with `[PSCustomObject]` for Mock return values.
  - Example:
    - OK: `Mock -CommandName Test-Path -MockWith { $true }`
    - OK: `Mock -CommandName Get-Item -MockWith { [PSCustomObject]@{ FullName = 'C:\file.txt' } }`
    - NG: `Mock -CommandName Get-Item -MockWith { [System.IO.FileInfo]::new('C:\file.txt') }`

### Avoid unnecessary parameter declarations

- Use `-ParameterFilter` to simulate behavior for specific arguments.
- This avoids needing `param()` inside `-MockWith`.
  - Example: `Mock -CommandName Get-Item -ParameterFilter { $Path -eq '*.txt' }`

## Tips

- Use [Pester.ps1](./scripts/Pester.ps1) for tests.

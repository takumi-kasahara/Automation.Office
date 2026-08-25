---
name: office-scripts
description: "Create, migrate, and debug Excel Office Scripts. Use when converting VBA macros to Office Scripts, fixing runtime errors, generating scripts, and validating by compiling to `.osts`."
argument-hint: "Target macro/script and expected behavior"
user-invocable: true
---

# Office Scripts Workflow

Refer to the Office Scripts documentation on Microsoft Learn to create or update Office Scripts (`*.ts`) that run in Excel.

## When to Use

- Convert VBA macro procedures to Office Scripts.
- Add new scripts under `Office Scripts/`.
- Debug Office Scripts runtime errors.
- Resolve API mismatch issues (for example, `Range` vs `RangeAreas`).
- Rebuild `.osts` artifacts after TypeScript changes.

## Inputs

1. Target function/procedure name.
2. Source file path (if converting from VBA).
3. Required behavior to preserve.
4. Acceptable behavior changes (if Office Scripts lacks equivalent UI/runtime features).

## Step-by-Step Procedure

1. Read the target implementation and summarize expected behavior.
2. Classify each operation:
   - **Direct mapping**: Worksheet/Range APIs available in Office Scripts.
   - **Alternative required**: similar API but different shape.
   - **Unsupported**: no Office Scripts equivalent (for example, VBA dialog prompts).
3. Verify APIs using Microsoft Learn before coding.
4. Implement script in `Office Scripts/<Name>.ts`:
   - Use `function main(workbook: ExcelScript.Workbook)`.
   - Prefer active worksheet/selection unless explicitly required otherwise.
   - Add null guards for methods that may return `undefined`.
5. Apply branching logic for shape differences:
   - If API returns `Range`, call methods directly.
   - If API returns `RangeAreas`, iterate `getAreas()` and apply to each `Range`.
6. Reconcile behavior differences and document them briefly.
7. Compile with `Compile.OfficeScripts.ps1` to update `.osts` outputs.
8. Report completion with preserved behavior, changes, and edge cases.

## Quality Criteria (Completion Checks)

- Correct Office Scripts object type usage (`Range` vs `RangeAreas`).
- No unsupported method calls on returned object types.
- Null/undefined guards present where needed.
- Script compiles through repository Office Scripts compile workflow.
- `.osts` artifact generated/updated.
- Any unavoidable behavior gap is called out explicitly.

## Common Pitfalls

- Calling `Range` methods on a `RangeAreas` object.
- Assuming VBA UI interactions (`MsgBox`) can be replicated directly.
- Forgetting to compile after editing `Office Scripts/*.ts`.
- Ignoring first-row `R1C1` edge cases that may produce `#REF!`.

## Repository-Specific Notes

- Office Scripts source: `Office Scripts/*.ts`
- Generated artifact: `Office Scripts/*.osts`
- Compile entrypoint: [`Compile.OfficeScripts.ps1`](../../../Compile.OfficeScripts.ps1)
- Decompile entrypoint: [`Decompile.OfficeScripts.ps1`](../../../Decompile.OfficeScripts.ps1)
- Project conventions: [`AGENTS.md`](../../../AGENTS.md)

---
name: office-scripts-to-vba
description: "Convert Excel Office Scripts (TypeScript) to VBA macros with Microsoft Learn API checks. Use when asked to migrate Office Scripts procedures back to VBA for desktop Excel workflows."
argument-hint: "Target Office Script name(s), output module location, and behavior constraints"
user-invocable: true
---

# Office Scripts to VBA Conversion

Convert Office Scripts into runnable Excel VBA procedures.

## When to Use

- User asks to convert Office Scripts to VBA (reverse of Office Scripts migration).
- Target source is `Office Scripts/*.ts` and behavior is worksheet/range oriented.
- You need Microsoft Learn API confirmation before writing VBA equivalents.
- You need to place converted procedures under `Apps/*/VBProject/` and validate through macro compile workflow.

## Inputs to Collect

1. Office Script function name and source location.
2. Target VBA module/workbook (`Normal.dotm`, `Personal.xlsb`, or another VBA host in this repo).
3. Expected behavior (must-preserve logic, acceptable simplifications).
4. Runtime constraints (desktop-only APIs allowed, user prompts, clipboard, dialogs, compatibility level).

## Conversion Workflow

1. Read the Office Script `main(...)` and summarize behavior in 3–5 bullets.
2. Identify Office Script dependencies and classify each as:
   - Directly mappable to Excel VBA object model
   - Requires VBA alternative or helper procedure
   - Not meaningful in desktop VBA context (for example, Power Automate return contract)
3. Ground API usage with Microsoft Learn:
   - Confirm Office Scripts method intent and equivalent Excel VBA API names.
   - Prefer official examples for both Office Scripts and VBA when available.
4. Implement VBA procedure(s) in the target VBA module:
   - Use `Sub` by default; use `Function` only when caller needs a return value.
   - Resolve workbook/sheet scope explicitly (`ActiveWorkbook` / `ThisWorkbook` / `ActiveSheet`) based on requirements.
   - Add guards for empty ranges and `Nothing` objects.
5. Handle multi-area ranges correctly:
   - Office Scripts `RangeAreas` often maps to VBA `Range.Areas` collection.
   - Iterate each `Area` when applying formulas, formatting, or value updates.
6. Preserve behavior intent, then document unavoidable differences.
7. Validate with repository macro workflow:
   - Update source under `Apps/*/VBProject/`.
   - Run macro compile workflow to verify no VBProject export/compile regressions.

## Mapping Patterns

- `workbook.getActiveWorksheet()` $\rightarrow$ `ActiveSheet`
- `workbook.getSelectedRange()` $\rightarrow$ `Selection`
- `worksheet.getUsedRange()` $\rightarrow$ `ActiveSheet.UsedRange`
- `range.getFormat().autofitColumns()` / `autofitRows()` $\rightarrow$ `range.EntireColumn.AutoFit` / `range.EntireRow.AutoFit`
- `getSpecialCells(ExcelScript.SpecialCellType.blanks)` (returns `RangeAreas`) $\rightarrow$ `Selection.SpecialCells(xlCellTypeBlanks)` (returns `Range` with `.Areas`)
- `area.setFormulaR1C1("=R[-1]C")` $\rightarrow$ `area.FormulaR1C1 = "=R[-1]C"`
- `range.getMergedAreas().getAreas()` + `unmerge()` pattern $\rightarrow$ iterate merged `Range` blocks, capture top-left value, call `UnMerge`, then restore value

## Quality Checks

- Converted VBA compiles and integrates with repository macro build/export workflow.
- Range vs multi-area behavior is preserved (`Range.Areas` iteration where needed).
- Empty selection/range cases are handled safely.
- Workbook/sheet scope assumptions are explicit.
- Comments and naming remain concise; behavior differences are explicit.

## Known Pitfalls

- Office Scripts APIs are strongly typed, while VBA is variant-heavy; type coercion can change behavior.
- Office Scripts `undefined` checks map to VBA `Nothing`, `IsEmpty`, and `Len(...) = 0` patterns depending on value type.
- `Selection.SpecialCells(...)` raises runtime error when no matching cells exist; guard with error handling.
- Script parameters and return values used by Power Automate have no direct VBA automation contract unless separately implemented.
- `ActiveSheet`/`Selection` assumptions can drift from original script intent; prefer explicit sheet references when possible.

## Output Contract

- Add or update VBA code under `Apps/*/VBProject/`.
- Keep procedure names aligned with source script intent (or document renamed procedures).
- Run macro compile workflow to validate generated Office artifacts.
- Report:
  - What was preserved
  - What changed (and why)
  - Any non-equivalent Office Scripts features left as explicit gaps

## Repository References

- Agent/project workflow: [AGENTS.md](../../../AGENTS.md)
- Source Office Scripts: [Office Scripts/](../../../Office Scripts/)
- Macro compile script: [Compile.Macros.ps1](../../../Compile.Macros.ps1)
- Macro decompile script: [Decompile.Macros.ps1](../../../Decompile.Macros.ps1)

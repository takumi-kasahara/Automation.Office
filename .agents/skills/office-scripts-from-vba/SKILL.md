---
name: office-scripts-from-vba
description: "Convert Excel VBA macros to Office Scripts (TypeScript) with Microsoft Learn API checks. Use when asked to migrate VBA procedures."
argument-hint: "Target VBA procedure name(s) and expected behavior"
user-invocable: true
---

# Office Scripts from VBA Conversion

Convert VBA macros into runnable Office Scripts.

## When to Use

- User asks to convert VBA to Office Scripts (TypeScript).
- Target is in `*.bas` or `*.cls` and behavior is worksheet/range oriented.
- You need Microsoft Learn API confirmation before writing Office Scripts.
- You need to generate scripts under `Office Scripts/` and compile to `.osts`.

## Inputs to Collect

1. VBA procedure name and source location.
2. Expected behavior (must-preserve logic, acceptable simplifications).
3. Runtime constraints (clipboard, dialogs, confirmations).

## Conversion Workflow

1. Read the VBA source procedure and summarize behavior in 3–5 bullets.
2. Identify VBA dependencies and classify each as:
   - Directly mappable (Range/Worksheet operations)
   - Requires Office Scripts alternative
   - Not supported in Office Scripts runtime (for example, MsgBox confirmation, COM API calls)
3. Ground API usage with Microsoft Learn:
   - Confirm method/class exists via docs search.
   - Prefer official Office Scripts samples when available.
4. Implement a new script in `Office Scripts/<ProcedureName>.ts`:
   - Use `function main(workbook: ExcelScript.Workbook)`.
   - Start from active sheet/range unless requirements specify otherwise.
   - Add null/undefined guards for methods that can return no range.
5. Handle `RangeAreas` correctly:
   - `getSpecialCells(...)` returns `RangeAreas`.
   - Apply updates by iterating `rangeAreas.getAreas()`.
6. Preserve behavior intent, then document unavoidable differences.
7. Compile with repository workflow (`Compile.OfficeScripts.ps1`) to produce/update `.osts`.

## Mapping Patterns

- `WorksheetExtensions.GetUsedRange` $\rightarrow$ `workbook.getActiveWorksheet().getUsedRange()`
- `WorksheetExtensions.GetSelectedRange` $\rightarrow$ `workbook.getSelectedRange()`
- `Range.EntireColumn.AutoFit` / `EntireRow.AutoFit` $\rightarrow$ `range.getFormat().autofitColumns()` / `autofitRows()`
- `SpecialCells(xlCellTypeBlanks)` + `FormulaR1C1`:
  - Use `selectedRange.getSpecialCells(ExcelScript.SpecialCellType.blanks)`
  - Iterate `getAreas()` and call `area.setFormulaR1C1('=R[-1]C')`
- `MergeArea.UnMerge` + value restore:
  - Get merged areas
  - Cache top-left value
  - `unmerge()` then `setValue(value)` on the former merged area

## Quality Checks

- Script compiles through project Office Scripts compile workflow.
- No unsupported API assumptions (especially `Range` vs `RangeAreas`).
- Handles empty selections/ranges safely.
- Comments and naming remain concise; behavior differences are explicit.

## Known Pitfalls

- `getSpecialCells(...)` may return `undefined` if no match.
- `RangeAreas` does not expose all `Range` instance methods directly.
- **`MsgBox` is not supported in Office Scripts.** Treat it as a non-convertible VBA UI dependency and explicitly call out the behavior gap.
- **COM API calls are not supported in Office Scripts.** Mark them as non-convertible unless the behavior can be replaced with an Office Scripts API.
- R1C1 formulas at top row can yield `#REF!`; consider guard logic if needed.

## Output Contract

- Add or update `Office Scripts/<ProcedureName>.ts`.
- Run compile workflow to update matching `.osts` artifact.
- Report:
  - What was preserved
  - What changed (and why)
  - Any non-convertible VBA features left as explicit gaps, especially `MsgBox` and COM API usage

## Repository References

- Agent/project workflow: [AGENTS.md](../../../AGENTS.md)
- Compile script: [Compile.OfficeScripts.ps1](../../../Compile.OfficeScripts.ps1)
- Decompile script: [Decompile.OfficeScripts.ps1](../../../Decompile.OfficeScripts.ps1)

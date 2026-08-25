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
   - **Unsupported**: no Office Scripts equivalent (for example, VBA `MsgBox` with Yes/No buttons, file picker dialogs).
3. Verify APIs using Microsoft Learn before coding.
4. Implement script in `Office Scripts/<Name>.ts`:
5. Compile with `Compile.OfficeScripts.ps1` to update `.osts` outputs.

### Basic `main` Function Example

```typescript
function main(workbook: ExcelScript.Workbook) {
  const selectedCell = workbook.getActiveCell();
  const selectedSheet = workbook.getActiveWorksheet();

  selectedCell.getFormat().getFill().setColor("yellow");
}
```

## Parameterized Scripts: Input and Output

Adding parameters and return values to `main` unlocks powerful capabilities beyond standalone execution.

### User Input Prompts

When a script is run from the Excel ribbon or a button, parameters (after `workbook: ExcelScript.Workbook`) trigger input dialogs:

```typescript
/**
 * Highlights cells exceeding a threshold value.
 * @param highlightThreshold The numeric threshold for comparison.
 * @param color Background color in #RRGGBB format or named HTML color (e.g., "orange").
 */
function main(
  workbook: ExcelScript.Workbook,
  highlightThreshold: number,
  color: string,
) {
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getUsedRange();
  const values = range.getValues();
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      if (values[row][col] >= highlightThreshold) {
        range.getCell(row, col).getFormat().getFill().setColor(color);
      }
    }
  }
}
```

### Parameter Features

- [Get user input for scripts](https://learn.microsoft.com/en-us/office/dev/scripts/develop/user-input)

| Feature             | Syntax                          | Example                                                 |
| ------------------- | ------------------------------- | ------------------------------------------------------- |
| **Optional**        | `?` modifier                    | `name?: string`                                         |
| **Default value**   | `= value` in signature          | `location: string = "Seattle"`                          |
| **Dropdown list**   | Union of literals               | `region: "North" \| "South" \| "East" \| "West"`        |
| **Workbook import** | 2D array `string[][]` per param | `productData: string[][], salesData: string[][]`        |
| **Nested objects**  | Interface with supported types  | `{ name: string,  job: { id: number; title: string } }` |

### Return Values: Output to Power Automate

- [Pass data to and from scripts in Power Automate](https://learn.microsoft.com/en-us/office/dev/scripts/develop/power-automate-parameters-returns)

Adding a return type enables scripts to pass data back to Power Automate flows:

```typescript
/**
 * Counts rows with data in the active worksheet.
 * @returns The number of used rows (excluding header).
 */
function main(workbook: ExcelScript.Workbook): number {
  const sheet = workbook.getActiveWorksheet();
  return sheet.getUsedRange().getRows().length - 1;
}
```

**Acceptable return types:** `string`, `number`, `boolean`, `unknown`, `object`, and arrays of these. Nested objects are allowed if all properties use supported types.

### Type Restrictions

1. First parameter **must** be `ExcelScript.Workbook`.
2. Supported primitive types: `string`, `number`, `boolean`, `unknown`, `object`.
3. Arrays (`[]` or `Array<T>`) of supported types, including nested arrays.
4. Union types allowed only if all literals share a single type (e.g., `"Left" | "Right"`, not `"Left" | 5`).
5. Object interfaces must be defined in the script; inline anonymous objects are also supported.

### JSDoc Documentation

JSDoc comments are displayed to users when running the script. Always document:

- Script purpose
- Each parameter's meaning and format
- Return value description (if applicable)

```typescript
/**
 * Applies tax rate to sales figures.
 * @param taxRate The current sales tax rate as a decimal (enter 12% as .12).
 * @returns The total tax amount calculated.
 */
function main(workbook: ExcelScript.Workbook, taxRate: number): number {
  // ...
}
```

### Power Automate Integration

When parameters/return types change, the "Run script" action in Power Automate must be reconfigured. The returned value appears as dynamic content named `result` in the flow.

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

## References

- [Office Scripts documentation](https://learn.microsoft.com/en-us/office/dev/scripts/)

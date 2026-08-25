/**
 * Behavior:
 * - Targets blank cells in the current selected range.
 * - Fills each blank cell with the value from the previous row in the same column.
 *
 * References:
 * - Range.getSpecialCells (blanks):
 *   https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/range-samples#get-groups-of-cells-based-on-special-criteria
 * - Range.setFormulaR1C1:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts
 */
/* global ExcelScript */
function main(workbook: ExcelScript.Workbook) {
  const selectedRange = workbook.getSelectedRange();

  const blankCells = selectedRange.getSpecialCells(
    ExcelScript.SpecialCellType.blanks,
  );
  if (!blankCells) {
    return;
  }

  const areas = blankCells.getAreas();
  for (const area of areas) {
    area.setFormulaR1C1("=R[-1]C");
  }
}

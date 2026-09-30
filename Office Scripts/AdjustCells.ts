/**
 * References:
 * - ExcelScript.RangeFormat.autofitColumns/autofitRows:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeformat?view=office-scripts
 */
/* global ExcelScript */
function main(workbook: ExcelScript.Workbook) {
  const worksheet = workbook.getActiveWorksheet();

  const usedRange = worksheet.getUsedRange();
  if (!usedRange) {
    return;
  }

  const format = usedRange.getFormat();
  format.autofitColumns();
  format.autofitRows();
}

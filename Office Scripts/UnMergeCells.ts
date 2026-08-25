/**
 * Behavior:
 * - Targets the current selected range.
 * - For each merged area in the selection:
 *   1) keep the top-left value,
 *   2) unmerge,
 *   3) fill the whole area with that value.
 *
 * References:
 * - ExcelScript.Range.unmerge:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts
 * - ExcelScript.Range.getMergedAreas / RangeAreas.getAreas:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.rangeareas?view=office-scripts
 */
/* global ExcelScript */
function main(workbook: ExcelScript.Workbook) {
  const selectedRange = workbook.getSelectedRange();

  const mergedAreas = selectedRange.getMergedAreas();
  if (!mergedAreas) {
    return;
  }

  const areas = mergedAreas.getAreas();
  for (const area of areas) {
    const value = area.getCell(0, 0).getValue();
    area.unmerge();
    area.setValue(value);
  }
}

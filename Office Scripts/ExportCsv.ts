type ScriptCellValue = string | number | boolean | Date | null;

/**
 * Behavior:
 * - Uses the current selection as the export target.
 * - Uses non-empty cells in the first selected row as export columns.
 * - Exports each selected area row-by-row as quoted comma-separated values.
 * - Replaces non-breaking spaces (U+00A0) with regular spaces before download.
 *
 * Notes:
 * - OfficeScript.downloadFile supports ".txt" or ".pdf" extensions.
 *   This script downloads CSV content as a ".txt" file.
 *
 * References:
 * - OfficeScript.downloadFile:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/officescript?view=office-scripts#officescript-officescript-downloadfile-function(1)
 * - OfficeScript.DownloadFileProperties:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/officescript/officescript.downloadfileproperties?view=office-scripts
 */
/* global ExcelScript, OfficeScript */
function main(workbook: ExcelScript.Workbook) {
  const selectedRange = workbook.getSelectedRange();
  if (!selectedRange) {
    return;
  }

  const values: ScriptCellValue[][] = selectedRange.getValues();
  const firstRow: ScriptCellValue[] = values[0] ?? [];

  const columnIndices: number[] = [];
  for (let columnIndex = 0; columnIndex < firstRow.length; columnIndex += 1) {
    const cellValue: ScriptCellValue = firstRow[columnIndex];
    if (toTrimmedText(cellValue) === '') {
      continue;
    }
    columnIndices.push(columnIndex);
  }

  const rows: string[] = [];
  for (const row of values) {
    const columns: string[] = [];

    for (const columnIndex of columnIndices) {
      const cellValue: ScriptCellValue = row[columnIndex];
      columns.push(`"${toTrimmedText(cellValue)}"`);
    }

    if (columns.length > 0) {
      rows.push(columns.join(','));
    }
  }

  const content = rows.join('\n').replace(/\u00A0/g, ' ');
  const sheetName = selectedRange.getWorksheet().getName();

  OfficeScript.downloadFile({
    name: `${sheetName}.txt`,
    content,
  });
}

function toTrimmedText(value: ScriptCellValue | undefined): string {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

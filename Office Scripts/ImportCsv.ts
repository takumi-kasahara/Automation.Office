/**
 * Behavior:
 * - Converts CSV lines into worksheet cells.
 * - Writes data to the active worksheet, starting at the selected range's top-left cell.
 * - Auto-fits used range columns and rows after import.
 *
 * Notes:
 * - Accepts CSV as an array of lines instead of a single string to avoid
 *   newline loss when copying from Excel cells.
 * - This follows the Microsoft Learn CSV parsing sample and keeps the same
 *   limitation for embedded newline characters inside quoted values.
 *
 * References:
 * - Convert CSV files to Excel workbooks:
 *   https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/convert-csv
 * - ExcelScript.Workbook.getActiveWorksheet:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts
 * - ExcelScript.Workbook.getSelectedRange:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.workbook?view=office-scripts
 * - ExcelScript.Range.getCell / setValue:
 *   https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range?view=office-scripts
 */
/* global ExcelScript */
function main(workbook: ExcelScript.Workbook, csvLines: string[]) {
  if (!csvLines || csvLines.length === 0) {
    return;
  }

  const sheet = workbook.getActiveWorksheet();
  const selectedRange = workbook.getSelectedRange();
  const csvMatchRegex = /(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g;

  const parsedRows: string[][] = [];
  let maxColumnCount = 0;

  for (const line of csvLines) {
    if (!line || line.length === 0) {
      continue;
    }

    const row = line.match(csvMatchRegex);
    if (!row || row.length === 0) {
      continue;
    }

    // Check for blanks at the start of the row.
    if (row[0].charAt(0) === ',') {
      row.unshift('');
    }

    // Remove the preceding comma and surrounding quotation marks.
    row.forEach((cell, index) => {
      let normalized = cell.indexOf(',') === 0 ? cell.substring(1) : cell;
      normalized
        = normalized.indexOf('"') === 0
          && normalized.lastIndexOf('"') === normalized.length - 1
          ? normalized.substring(1, normalized.length - 1)
          : normalized;
      row[index] = normalized;
    });

    parsedRows.push(row);
    if (row.length > maxColumnCount) {
      maxColumnCount = row.length;
    }
  }

  if (parsedRows.length === 0 || maxColumnCount === 0) {
    return;
  }

  // setValue writes cell-by-cell, so pad short rows to preserve CSV column alignment.
  const values: string[][] = parsedRows.map(row => {
    if (row.length === maxColumnCount) {
      return row;
    }

    const padded = [...row];
    while (padded.length < maxColumnCount) {
      padded.push('');
    }

    return padded;
  });

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      const targetCell = selectedRange.getCell(rowIndex, columnIndex);
      targetCell.setValue(row[columnIndex]);
    }
  }

  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    return;
  }

  const format = usedRange.getFormat();
  format.autofitColumns();
  format.autofitRows();
}

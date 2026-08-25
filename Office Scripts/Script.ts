function main(workbook: ExcelScript.Workbook) {
  // Get the active cell and worksheet.
  const selectedCell = workbook.getActiveCell();
  const selectedSheet = workbook.getActiveWorksheet();

  // Set fill color to yellow for the selected cell.
  selectedCell.getFormat().getFill().setColor('yellow');

  // TODO: Write code or use the Insert action button below.
}

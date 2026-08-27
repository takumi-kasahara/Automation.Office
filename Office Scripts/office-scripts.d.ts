/* eslint-disable */

declare namespace ExcelScript {
  type CellValue = string | number | boolean | Date | null;

  enum SpecialCellType {
    blanks = "blanks",
  }

  interface Workbook {
    getActiveWorksheet(): Worksheet;
    getSelectedRange(): Range;
    getSelectedRanges(): RangeAreas;
  }

  interface Worksheet {
    getName(): string;
    getUsedRange(valuesOnly?: boolean): Range | undefined;
    getRange(address: string): Range;
  }

  interface Range {
    getWorksheet(): Worksheet;
    getFormat(): RangeFormat;
    getUsedRange(valuesOnly?: boolean): Range | undefined;
    getValues(): CellValue[][];
    getSpecialCells(
      cellType: SpecialCellType,
      cellValueType?: unknown,
    ): RangeAreas | undefined;
    getMergedAreas(): RangeAreas | undefined;
    getCell(rowIndex: number, columnIndex: number): Range;
    getValue(): CellValue;
    setValue(value: CellValue): void;
    setFormulaR1C1(formulaR1C1: string): void;
    unmerge(): void;
  }

  interface RangeAreas {
    getAreas(): Range[];
    getFormat(): RangeFormat;
  }

  interface RangeFormat {
    autofitColumns(): void;
    autofitRows(): void;
  }
}

declare namespace OfficeScript {
  interface DownloadFileProperties {
    content: string;
    name: string;
  }

  function downloadFile(fileProperties: DownloadFileProperties): void;
}

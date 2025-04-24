export interface ExcelData {
  headers: { [sheetName: string]: string[] };
  rows: { [sheetName: string]: (string | number | Date)[][] };
  sheetNames: string[];
  activeSheet: string;
  formats?: { [sheetName: string]: string[][] };
}

export interface CellPosition {
  row: number;
  col: number;
  sheet: string;
}

export interface CellEdit {
  position: CellPosition;
  value: string | number;
}
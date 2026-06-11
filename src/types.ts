/** Cell update: cell address → value map */
export interface CellUpdates {
  [cellRef: string]: string | number | boolean | null;
}

/** Result from listing sheets */
export interface SheetInfo {
  name: string;
  rows: number;
  cols: number;
}

/** Result from reading sheet data */
export interface ReadResult {
  sheet: string;
  rows: number;
  cols: number;
  data: Record<string, unknown>[] | unknown[][];
  format: "json" | "array";
}

/** Result from creating a file */
export interface CreateResult {
  file_path: string;
  sheet_name: string;
  rows: number;
  cols: number;
}

/** Result from adding a sheet */
export interface AddSheetResult {
  file_path: string;
  sheet_name: string;
  rows: number;
  cols: number;
  sheets: string[];
}

/** Result from updating cells */
export interface UpdateCellsResult {
  file_path: string;
  sheet: string;
  updated: number;
}

/** Result from listing sheets */
export interface ListSheetsResult {
  file_path: string;
  sheets: SheetInfo[];
}

/** xlsx read format */
export enum ReadFormat {
  JSON = "json",
  ARRAY = "array",
}

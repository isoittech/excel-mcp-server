/**
 * Excel MCPサーバーの型定義
 */

import ExcelJS from 'exceljs';

/**
 * ツールの引数の型定義
 */
export interface ReadExcelArgs {
  filePath: string;
  sheetName?: string;
  range?: string;
  previewOnly?: boolean;
}

export interface WriteExcelArgs {
  filePath: string;
  sheetName?: string;
  data: any[];
  startCell?: string;
  writeHeaders?: boolean;
}

export interface CreateSheetArgs {
  filePath: string;
  sheetName: string;
}

export interface CreateExcelArgs {
  filePath: string;
  sheetName?: string;
}

export interface GetWorkbookMetadataArgs {
  filePath: string;
  includeRanges?: boolean;
}

export interface RenameWorksheetArgs {
  filePath: string;
  oldName: string;
  newName: string;
}

export interface DeleteWorksheetArgs {
  filePath: string;
  sheetName: string;
}

export interface CopyWorksheetArgs {
  filePath: string;
  sourceSheet: string;
  targetSheet: string;
}

export interface ApplyFormulaArgs {
  filePath: string;
  sheetName: string;
  cell: string;
  formula: string;
}

export interface ValidateFormulaSyntaxArgs {
  filePath: string;
  sheetName: string;
  cell: string;
  formula: string;
}

export interface FormatRangeArgs {
  filePath: string;
  sheetName: string;
  startCell: string;
  endCell?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  bgColor?: string;
  borderStyle?: string;
  borderColor?: string;
  numberFormat?: string;
  alignment?: string;
  wrapText?: boolean;
  mergeCells?: boolean;
  protection?: Record<string, any>;
  conditionalFormat?: Record<string, any>;
}

export interface MergeCellsArgs {
  filePath: string;
  sheetName: string;
  startCell: string;
  endCell: string;
}

export interface UnmergeCellsArgs {
  filePath: string;
  sheetName: string;
  startCell: string;
  endCell: string;
}

export interface CopyRangeArgs {
  filePath: string;
  sheetName: string;
  sourceStart: string;
  sourceEnd: string;
  targetStart: string;
  targetSheet?: string;
}

export interface DeleteRangeArgs {
  filePath: string;
  sheetName: string;
  startCell: string;
  endCell: string;
  shiftDirection?: 'up' | 'left';
}

export interface ValidateExcelRangeArgs {
  filePath: string;
  sheetName: string;
  startCell: string;
  endCell?: string;
}

export interface CreateChartArgs {
  filePath: string;
  sheetName: string;
  dataRange: string;
  chartType: string;
  targetCell: string;
  title?: string;
  xAxis?: string;
  yAxis?: string;
}

export interface CreatePivotTableArgs {
  filePath: string;
  sheetName: string;
  dataRange: string;
  rows: string[];
  values: string[];
  columns?: string[];
  aggFunc?: string;
}

/**
 * ツールの戻り値の型定義
 */
export interface ToolResponse {
  content: {
    type: string;
    text: string;
  }[];
  isError?: boolean;
}

/**
 * セル座標の型定義
 */
export interface CellCoordinate {
  row: number;
  col: number;
}

/**
 * セル範囲の型定義
 */
export interface CellRange {
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

/**
 * ワークブックキャッシュの型定義
 */
export interface WorkbookCache {
  [filePath: string]: ExcelJS.Workbook;
}
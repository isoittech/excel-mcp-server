/**
 * データ操作ハンドラー
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { ReadExcelArgs, WriteExcelArgs, ToolResponse } from '../types/index.js';
import { parseCellOrRange } from '../utils/cellUtils.js';
import { loadWorkbook, getWorksheet, getExcelPath } from '../utils/fileUtils.js';
import { WorkbookCache } from '../types/index.js';

/**
 * Excelファイルからデータを読み込む
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleReadExcel(args: ReadExcelArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, range, previewOnly = false } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 範囲を解析
    let startCol = 1;
    let startRow = 1;
    let endCol = worksheet.columnCount || 100; // 列数が取得できない場合は100とする
    let endRow = worksheet.rowCount || 100; // 行数が取得できない場合は100とする
    
    if (range) {
      const parsedRange = parseCellOrRange(range);
      startCol = parsedRange.startCol;
      startRow = parsedRange.startRow;
      endCol = parsedRange.endCol;
      endRow = parsedRange.endRow;
    }
    
    // プレビューモードの場合は行数と列数を制限
    if (previewOnly) {
      endRow = Math.min(endRow, startRow + 9); // 最大10行
      endCol = Math.min(endCol, startCol + 9); // 最大10列
    }
    
    // データを読み込む
    const data = [];
    const headers = [];
    
    // ヘッダー行を取得（最初の行）
    for (let col = startCol; col <= endCol; col++) {
      const cell = worksheet.getCell(startRow, col);
      headers.push(cell.value);
    }
    
    // データ行を取得
    for (let row = startRow + 1; row <= endRow; row++) {
      const rowData: Record<string, any> = {};
      let hasData = false;
      
      for (let col = startCol; col <= endCol; col++) {
        const cell = worksheet.getCell(row, col);
        const header = headers[col - startCol];
        const headerKey = header ? String(header) : `Column${col - startCol + 1}`;
        
        if (cell.value !== null && cell.value !== undefined) {
          rowData[headerKey] = cell.value;
          hasData = true;
        }
      }
      
      // 空の行はスキップ
      if (hasData) {
        data.push(rowData);
      }
    }
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(data, null, 2)
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `データ読み込みエラー: ${error.message}` : 'データ読み込み中に不明なエラーが発生しました'
    );
  }
}

/**
 * Excelファイルにデータを書き込む
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleWriteExcel(args: WriteExcelArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, data, startCell = 'A1', writeHeaders = true } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 開始セルの座標を取得
    const { row: startRow, col: startCol } = parseCellOrRange(startCell).startRow 
      ? { row: parseCellOrRange(startCell).startRow, col: parseCellOrRange(startCell).startCol }
      : { row: 1, col: 1 };
    
    // データが配列の配列の場合（2次元配列）
    if (Array.isArray(data) && data.every(item => Array.isArray(item))) {
      for (let i = 0; i < data.length; i++) {
        const rowData = data[i];
        for (let j = 0; j < rowData.length; j++) {
          worksheet.getCell(startRow + i, startCol + j).value = rowData[j];
        }
      }
    }
    // データがオブジェクトの配列の場合
    else if (Array.isArray(data) && data.every(item => typeof item === 'object' && item !== null)) {
      // ヘッダーを書き込む
      if (writeHeaders && data.length > 0) {
        const headers = Object.keys(data[0]);
        for (let i = 0; i < headers.length; i++) {
          worksheet.getCell(startRow, startCol + i).value = headers[i];
        }
      }
      
      // データを書き込む
      const rowOffset = writeHeaders ? 1 : 0;
      for (let i = 0; i < data.length; i++) {
        const rowData = data[i];
        const keys = Object.keys(rowData);
        for (let j = 0; j < keys.length; j++) {
          worksheet.getCell(startRow + i + rowOffset, startCol + j).value = rowData[keys[j]];
        }
      }
    } else {
      throw new McpError(
        ErrorCode.InvalidParams,
        'データは配列の配列またはオブジェクトの配列である必要があります'
      );
    }
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: 'データが正常に保存されました'
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `データ書き込みエラー: ${error.message}` : 'データ書き込み中に不明なエラーが発生しました'
    );
  }
}
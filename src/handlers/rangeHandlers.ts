/**
 * セル範囲操作ハンドラー
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { 
  CopyRangeArgs, 
  DeleteRangeArgs, 
  ValidateExcelRangeArgs, 
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, getWorksheet } from '../utils/fileUtils.js';
import { cellRefToCoordinate, parseCellOrRange } from '../utils/cellUtils.js';

/**
 * セル範囲をコピー
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCopyRange(args: CopyRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, sourceStart, sourceEnd, targetStart, targetSheet } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const sourceWorksheet = getWorksheet(workbook, sheetName);
    const targetWorksheet = targetSheet ? getWorksheet(workbook, targetSheet) : sourceWorksheet;
    
    // 範囲を解析
    const sourceRange = parseCellOrRange(`${sourceStart}:${sourceEnd}`);
    const targetStartCoord = cellRefToCoordinate(targetStart);
    
    // ソース範囲のサイズを計算
    const rowCount = sourceRange.endRow - sourceRange.startRow + 1;
    const colCount = sourceRange.endCol - sourceRange.startCol + 1;
    
    // セル値と書式をコピー
    for (let i = 0; i < rowCount; i++) {
      for (let j = 0; j < colCount; j++) {
        const sourceCell = sourceWorksheet.getCell(
          sourceRange.startRow + i,
          sourceRange.startCol + j
        );
        
        const targetCell = targetWorksheet.getCell(
          targetStartCoord.row + i,
          targetStartCoord.col + j
        );
        
        // 値をコピー
        targetCell.value = sourceCell.value;
        
        // 書式をコピー
        targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        
        // 数値書式をコピー
        if (sourceCell.numFmt) {
          targetCell.numFmt = sourceCell.numFmt;
        }
      }
    }
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `範囲 ${sourceStart}:${sourceEnd} が ${targetSheet || sheetName} シートの ${targetStart} に正常にコピーされました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `範囲コピーエラー: ${error.message}` : '範囲コピー中に不明なエラーが発生しました'
    );
  }
}

/**
 * セル範囲を削除
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleDeleteRange(args: DeleteRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell, shiftDirection = 'up' } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 範囲を解析
    const range = parseCellOrRange(`${startCell}:${endCell}`);
    
    // 範囲内のセルをクリア
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = worksheet.getCell(row, col);
        cell.value = null;
        
        // スタイルもクリア
        cell.style = {};
        cell.numFmt = '';
      }
    }
    
    // セルをシフト
    if (shiftDirection === 'up') {
      // 上方向へのシフト
      // 削除範囲の下にあるセルを上に移動
      const colCount = range.endCol - range.startCol + 1;
      const rowCount = range.endRow - range.startRow + 1;
      
      for (let col = range.startCol; col <= range.endCol; col++) {
        for (let row = range.endRow + 1; row <= worksheet.rowCount; row++) {
          const sourceCell = worksheet.getCell(row, col);
          const targetCell = worksheet.getCell(row - rowCount, col);
          
          // 値と書式をコピー
          targetCell.value = sourceCell.value;
          targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
          if (sourceCell.numFmt) {
            targetCell.numFmt = sourceCell.numFmt;
          }
          
          // 元のセルをクリア
          sourceCell.value = null;
          sourceCell.style = {};
          sourceCell.numFmt = '';
        }
      }
    } else if (shiftDirection === 'left') {
      // 左方向へのシフト
      // 削除範囲の右にあるセルを左に移動
      const colCount = range.endCol - range.startCol + 1;
      
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.endCol + 1; col <= worksheet.columnCount; col++) {
          const sourceCell = worksheet.getCell(row, col);
          const targetCell = worksheet.getCell(row, col - colCount);
          
          // 値と書式をコピー
          targetCell.value = sourceCell.value;
          targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
          if (sourceCell.numFmt) {
            targetCell.numFmt = sourceCell.numFmt;
          }
          
          // 元のセルをクリア
          sourceCell.value = null;
          sourceCell.style = {};
          sourceCell.numFmt = '';
        }
      }
    }
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `範囲 ${startCell}:${endCell} が正常に削除され、セルが${shiftDirection === 'up' ? '上' : '左'}方向にシフトされました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `範囲削除エラー: ${error.message}` : '範囲削除中に不明なエラーが発生しました'
    );
  }
}

/**
 * Excelの範囲を検証
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleValidateExcelRange(args: ValidateExcelRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 範囲を解析
    let range;
    try {
      if (endCell) {
        range = parseCellOrRange(`${startCell}:${endCell}`);
      } else {
        range = parseCellOrRange(startCell);
      }
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: `範囲形式が無効です: ${error instanceof Error ? error.message : '不明なエラー'}`
        }]
      };
    }
    
    // 範囲がワークシートの境界内にあるか確認
    const maxRow = worksheet.rowCount || 1048576; // Excelの最大行数
    const maxCol = worksheet.columnCount || 16384; // Excelの最大列数
    
    if (range.startRow < 1 || range.startCol < 1 || range.endRow > maxRow || range.endCol > maxCol) {
      return {
        content: [{
          type: 'text',
          text: `範囲がワークシートの境界外です: 行(${range.startRow}-${range.endRow}), 列(${range.startCol}-${range.endCol})`
        }]
      };
    }
    
    // 範囲内のデータ情報を取得
    let cellCount = 0;
    let nonEmptyCellCount = 0;
    
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        cellCount++;
        const cell = worksheet.getCell(row, col);
        if (cell.value !== null && cell.value !== undefined) {
          nonEmptyCellCount++;
        }
      }
    }
    
    return {
      content: [{
        type: 'text',
        text: `範囲 ${startCell}${endCell ? `:${endCell}` : ''} は有効です。` +
              `セル数: ${cellCount}, データを含むセル: ${nonEmptyCellCount}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `範囲検証エラー: ${error.message}` : '範囲検証中に不明なエラーが発生しました'
    );
  }
}
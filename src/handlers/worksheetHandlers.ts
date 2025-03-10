/**
 * ワークシート操作ハンドラー
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { 
  RenameWorksheetArgs, 
  DeleteWorksheetArgs, 
  CopyWorksheetArgs, 
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath } from '../utils/fileUtils.js';

/**
 * ワークシートの名前を変更
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleRenameWorksheet(args: RenameWorksheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, oldName, newName } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // 元のシートが存在するか確認
    const worksheet = workbook.getWorksheet(oldName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `シート ${oldName} が見つかりません`);
    }
    
    // 新しい名前のシートが既に存在するか確認
    if (workbook.getWorksheet(newName)) {
      throw new McpError(ErrorCode.InvalidParams, `シート ${newName} は既に存在します`);
    }
    
    // シート名を変更
    worksheet.name = newName;
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `ワークシートの名前が ${oldName} から ${newName} に正常に変更されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `シート名変更エラー: ${error.message}` : 'シート名変更中に不明なエラーが発生しました'
    );
  }
}

/**
 * ワークシートを削除
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleDeleteWorksheet(args: DeleteWorksheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // シートが存在するか確認
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `シート ${sheetName} が見つかりません`);
    }
    
    // ワークブックに複数のシートがあるか確認
    if (workbook.worksheets.length === 1) {
      throw new McpError(
        ErrorCode.InvalidParams,
        'ワークブック内の唯一のシートは削除できません'
      );
    }
    
    // シートを削除
    workbook.removeWorksheet(worksheet.id);
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `ワークシート ${sheetName} が正常に削除されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `シート削除エラー: ${error.message}` : 'シート削除中に不明なエラーが発生しました'
    );
  }
}

/**
 * ワークシートをコピー
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCopyWorksheet(args: CopyWorksheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sourceSheet, targetSheet } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // 元のシートが存在するか確認
    const sourceWorksheet = workbook.getWorksheet(sourceSheet);
    if (!sourceWorksheet) {
      throw new McpError(ErrorCode.InvalidParams, `元のシート ${sourceSheet} が見つかりません`);
    }
    
    // 対象のシート名が既に存在するか確認
    if (workbook.getWorksheet(targetSheet)) {
      throw new McpError(ErrorCode.InvalidParams, `対象のシート ${targetSheet} は既に存在します`);
    }
    
    // 新しいシートを作成
    const targetWorksheet = workbook.addWorksheet(targetSheet);
    
    // プロパティをコピー
    targetWorksheet.properties = JSON.parse(JSON.stringify(sourceWorksheet.properties));
    targetWorksheet.properties.tabColor = sourceWorksheet.properties.tabColor;
    
    // 列のプロパティと幅をコピー
    sourceWorksheet.columns.forEach((column, index) => {
      if (column.width) {
        targetWorksheet.getColumn(index + 1).width = column.width;
      }
    });
    
    // 行の高さと値をコピー
    sourceWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const targetRow = targetWorksheet.getRow(rowNumber);
      targetRow.height = row.height;
      
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const targetCell = targetRow.getCell(colNumber);
        targetCell.value = cell.value;
        targetCell.style = JSON.parse(JSON.stringify(cell.style));
      });
      
      targetRow.commit();
    });
    
    // 結合セルをコピー
    // ExcelJSでは結合セルの直接アクセスが難しいため、
    // 代替手段として各セルを調べて結合セルを検出する
    sourceWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (cell.isMerged) {
          // このセルが結合セルの一部である場合、
          // 結合範囲の左上のセルを特定して結合を適用
          const master = cell.master;
          if (master.address === cell.address) {
            // このセルが結合の左上（マスター）セルの場合
            // 結合範囲を特定して適用
            // 注: 実際の実装では、結合範囲を正確に特定する方法が必要
            // ここでは簡易的な実装
            let endRow = rowNumber;
            let endCol = colNumber;
            
            // 結合範囲を推定（実際の実装ではより正確な方法が必要）
            for (let r = rowNumber; r <= sourceWorksheet.rowCount; r++) {
              const testCell = sourceWorksheet.getCell(r, colNumber);
              if (testCell.master && testCell.master.address === master.address) {
                endRow = r;
              } else if (r > rowNumber) {
                break;
              }
            }
            
            for (let c = colNumber; c <= sourceWorksheet.columnCount; c++) {
              const testCell = sourceWorksheet.getCell(rowNumber, c);
              if (testCell.master && testCell.master.address === master.address) {
                endCol = c;
              } else if (c > colNumber) {
                break;
              }
            }
            
            if (endRow > rowNumber || endCol > colNumber) {
              targetWorksheet.mergeCells(rowNumber, colNumber, endRow, endCol);
            }
          }
        }
      });
    });
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `ワークシート ${sourceSheet} が ${targetSheet} に正常にコピーされました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `シートコピーエラー: ${error.message}` : 'シートコピー中に不明なエラーが発生しました'
    );
  }
}
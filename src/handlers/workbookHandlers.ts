/**
 * ワークブック操作ハンドラー
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { 
  CreateExcelArgs, 
  CreateSheetArgs, 
  GetWorkbookMetadataArgs, 
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, ensureDirectoryExists } from '../utils/fileUtils.js';

/**
 * 新規Excelファイルを作成
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCreateExcel(args: CreateExcelArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName = 'Sheet1' } = args;
    const fullPath = getExcelPath(filePath);
    
    // ファイルが既に存在する場合はエラー
    if (fs.existsSync(fullPath)) {
      throw new McpError(ErrorCode.InvalidParams, `ファイル ${filePath} は既に存在します`);
    }
    
    // ディレクトリが存在しない場合は作成
    ensureDirectoryExists(fullPath);
    
    // 新規ワークブックを作成
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet(sheetName);
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    // キャッシュに追加
    workbookCache[fullPath] = workbook;
    
    return {
      content: [{
        type: 'text',
        text: `Excelファイルが ${fullPath} に正常に作成されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Excelファイル作成エラー: ${error.message}` : 'Excelファイル作成中に不明なエラーが発生しました'
    );
  }
}

/**
 * 新規シートを作成
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCreateSheet(args: CreateSheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // 同名のシートが既に存在する場合はエラー
    if (workbook.getWorksheet(sheetName)) {
      throw new McpError(ErrorCode.InvalidParams, `シート ${sheetName} は既に存在します`);
    }
    
    // 新規シートを作成
    workbook.addWorksheet(sheetName);
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `シート ${sheetName} が正常に作成されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `シート作成エラー: ${error.message}` : 'シート作成中に不明なエラーが発生しました'
    );
  }
}

/**
 * ワークブックのメタデータを取得
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleGetWorkbookMetadata(args: GetWorkbookMetadataArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, includeRanges = false } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    const metadata: any = {
      fileName: path.basename(fullPath),
      filePath: fullPath,
      sheets: []
    };
    
    // 各ワークシートの情報を取得
    workbook.eachSheet((worksheet, sheetId) => {
      const sheetInfo: any = {
        name: worksheet.name,
        id: sheetId,
        rowCount: worksheet.rowCount,
        columnCount: worksheet.columnCount,
        hidden: worksheet.state === 'hidden' || worksheet.state === 'veryHidden'
      };
      
      // 使用範囲の情報を含める場合
      if (includeRanges) {
        // 使用範囲を取得（内容のある最後のセルを探す）
        let maxRow = 0;
        let maxCol = 0;
        
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          maxRow = Math.max(maxRow, rowNumber);
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            maxCol = Math.max(maxCol, colNumber);
          });
        });
        
        if (maxRow > 0 && maxCol > 0) {
          sheetInfo.usedRange = {
            startRow: 1,
            startColumn: 1,
            endRow: maxRow,
            endColumn: maxCol
          };
        }
      }
      
      metadata.sheets.push(sheetInfo);
    });
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(metadata, null, 2)
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `メタデータ取得エラー: ${error.message}` : 'メタデータ取得中に不明なエラーが発生しました'
    );
  }
}
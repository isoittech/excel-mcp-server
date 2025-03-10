/**
 * ファイル操作ユーティリティ関数
 */

import fs from 'fs';
import path from 'path';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import ExcelJS from 'exceljs';
import { WorkbookCache } from '../types/index.js';

// Excelファイルのデフォルトの保存先
const DEFAULT_EXCEL_FILES_PATH = './excel_files';

// 環境変数からExcelファイルのパスを取得
export const EXCEL_FILES_PATH = process.env.EXCEL_FILES_PATH || DEFAULT_EXCEL_FILES_PATH;

// ディレクトリが存在しない場合は作成
if (!fs.existsSync(EXCEL_FILES_PATH)) {
  fs.mkdirSync(EXCEL_FILES_PATH, { recursive: true });
}

/**
 * Excelファイルのフルパスを取得
 * @param filename - Excelファイル名
 * @returns Excelファイルのフルパス
 */
export function getExcelPath(filename: string): string {
  // すでに絶対パスの場合はそのまま返す
  if (path.isAbsolute(filename)) {
    return filename;
  }
  
  // 設定されたExcelファイルパスを使用
  return path.join(EXCEL_FILES_PATH, filename);
}

/**
 * ワークブックをロード
 * @param filePath - Excelファイルのパス
 * @param workbookCache - ワークブックキャッシュ
 * @returns ロードされたワークブック
 */
export async function loadWorkbook(filePath: string, workbookCache: WorkbookCache): Promise<ExcelJS.Workbook> {
  // キャッシュにある場合はキャッシュから返す
  if (workbookCache[filePath]) {
    return workbookCache[filePath];
  }

  // ファイルが存在しない場合はエラー
  if (!fs.existsSync(filePath)) {
    throw new McpError(ErrorCode.InvalidParams, `ファイル ${filePath} が見つかりません`);
  }

  // ワークブックをロード
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  // キャッシュに追加
  workbookCache[filePath] = workbook;
  
  return workbook;
}

/**
 * ワークシートを取得
 * @param workbook - ワークブック
 * @param sheetName - シート名（省略時は最初のシート）
 * @returns ワークシート
 */
export function getWorksheet(workbook: ExcelJS.Workbook, sheetName?: string): ExcelJS.Worksheet {
  const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  
  if (!worksheet) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `シート ${sheetName || '(最初のシート)'} が見つかりません`
    );
  }
  
  return worksheet;
}

/**
 * ファイルの拡張子を確認
 * @param filePath - ファイルパス
 * @param expectedExt - 期待する拡張子（例：'.xlsx'）
 * @returns 拡張子が一致する場合はtrue、それ以外はfalse
 */
export function checkFileExtension(filePath: string, expectedExt: string): boolean {
  const ext = path.extname(filePath).toLowerCase();
  return ext === expectedExt.toLowerCase();
}

/**
 * Excelファイルの拡張子を確認
 * @param filePath - ファイルパス
 * @returns 拡張子が.xlsxの場合はtrue、それ以外はfalse
 */
export function isExcelFile(filePath: string): boolean {
  return checkFileExtension(filePath, '.xlsx');
}

/**
 * ファイルパスのディレクトリが存在しない場合は作成
 * @param filePath - ファイルパス
 */
export function ensureDirectoryExists(filePath: string): void {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}
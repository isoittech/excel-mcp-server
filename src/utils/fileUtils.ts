/**
 * File operation utility functions
 */

import fs from 'fs';
import path from 'path';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import ExcelJS from 'exceljs';
import { WorkbookCache } from '../types/index.js';

// Default directory for Excel files
const DEFAULT_EXCEL_FILES_PATH = './excel_files';

// Get Excel files path from environment variable
export const EXCEL_FILES_PATH = process.env.EXCEL_FILES_PATH || DEFAULT_EXCEL_FILES_PATH;

// Create directory if it doesn't exist
if (!fs.existsSync(EXCEL_FILES_PATH)) {
  fs.mkdirSync(EXCEL_FILES_PATH, { recursive: true });
}

/**
 * Get full path of Excel file
 * @param filename - Excel filename
 * @returns Full path of Excel file
 */
export function getExcelPath(filename: string): string {
  // Return as is if already an absolute path
  if (path.isAbsolute(filename)) {
    return filename;
  }
  
  // Use configured Excel files path
  return path.join(EXCEL_FILES_PATH, filename);
}

/**
 * Load workbook
 * @param filePath - Excel file path
 * @param workbookCache - Workbook cache
 * @returns Loaded workbook
 */
export async function loadWorkbook(filePath: string, workbookCache: WorkbookCache): Promise<ExcelJS.Workbook> {
  // Return from cache if available
  if (workbookCache[filePath]) {
    return workbookCache[filePath];
  }

  // Error if file doesn't exist
  if (!fs.existsSync(filePath)) {
    throw new McpError(ErrorCode.InvalidParams, `File ${filePath} not found`);
  }

  // Load workbook
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  // Add to cache
  workbookCache[filePath] = workbook;
  
  return workbook;
}

/**
 * Get worksheet
 * @param workbook - Workbook
 * @param sheetName - Sheet name (first sheet if omitted)
 * @returns Worksheet
 */
export function getWorksheet(workbook: ExcelJS.Workbook, sheetName?: string): ExcelJS.Worksheet {
  const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  
  if (!worksheet) {
    throw new McpError(
      ErrorCode.InvalidParams,
      `Sheet ${sheetName || '(first sheet)'} not found`
    );
  }
  
  return worksheet;
}

/**
 * Check file extension
 * @param filePath - File path
 * @param expectedExt - Expected extension (e.g., '.xlsx')
 * @returns true if extension matches, false otherwise
 */
export function checkFileExtension(filePath: string, expectedExt: string): boolean {
  const ext = path.extname(filePath).toLowerCase();
  return ext === expectedExt.toLowerCase();
}

/**
 * Check if file is an Excel file
 * @param filePath - File path
 * @returns true if extension is .xlsx, false otherwise
 */
export function isExcelFile(filePath: string): boolean {
  return checkFileExtension(filePath, '.xlsx');
}

/**
 * Create directory if it doesn't exist for file path
 * @param filePath - File path
 */
export function ensureDirectoryExists(filePath: string): void {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}
/**
 * Workbook operation handlers
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
 * Create new Excel file
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleCreateExcel(args: CreateExcelArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName = 'Sheet1' } = args;
    const fullPath = getExcelPath(filePath);
    
    // Error if file already exists
    if (fs.existsSync(fullPath)) {
      throw new McpError(ErrorCode.InvalidParams, `File ${filePath} already exists`);
    }
    
    // Create directory if it doesn't exist
    ensureDirectoryExists(fullPath);
    
    // Create new workbook
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet(sheetName);
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    // Add to cache
    workbookCache[fullPath] = workbook;
    
    return {
      content: [{
        type: 'text',
        text: `Excel file successfully created at ${fullPath}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Excel file creation error: ${error.message}` : 'Unknown error occurred while creating Excel file'
    );
  }
}

/**
 * Create new worksheet
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleCreateSheet(args: CreateSheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // Error if sheet with same name already exists
    if (workbook.getWorksheet(sheetName)) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} already exists`);
    }
    
    // Create new sheet
    workbook.addWorksheet(sheetName);
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Sheet ${sheetName} successfully created`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Sheet creation error: ${error.message}` : 'Unknown error occurred while creating sheet'
    );
  }
}

/**
 * Get workbook metadata
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
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
    
    // Get information for each worksheet
    workbook.eachSheet((worksheet, sheetId) => {
      const sheetInfo: any = {
        name: worksheet.name,
        id: sheetId,
        rowCount: worksheet.rowCount,
        columnCount: worksheet.columnCount,
        hidden: worksheet.state === 'hidden' || worksheet.state === 'veryHidden'
      };
      
      // Include used range information if requested
      if (includeRanges) {
        // Find the last cell with content (to determine used range)
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
      error instanceof Error ? `Metadata retrieval error: ${error.message}` : 'Unknown error occurred while retrieving metadata'
    );
  }
}
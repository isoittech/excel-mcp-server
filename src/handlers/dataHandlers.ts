/**
 * Data operation handlers
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { ReadExcelArgs, WriteExcelArgs, ToolResponse } from '../types/index.js';
import { parseCellOrRange } from '../utils/cellUtils.js';
import { loadWorkbook, getWorksheet, getExcelPath } from '../utils/fileUtils.js';
import { WorkbookCache } from '../types/index.js';

/**
 * Read data from Excel file
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleReadExcel(args: ReadExcelArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, range, previewOnly = false } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse range
    let startCol = 1;
    let startRow = 1;
    let endCol = worksheet.columnCount || 100; // Default to 100 if column count is not available
    let endRow = worksheet.rowCount || 100; // Default to 100 if row count is not available
    
    if (range) {
      const parsedRange = parseCellOrRange(range);
      startCol = parsedRange.startCol;
      startRow = parsedRange.startRow;
      endCol = parsedRange.endCol;
      endRow = parsedRange.endRow;
    }
    
    // Limit rows and columns in preview mode
    if (previewOnly) {
      endRow = Math.min(endRow, startRow + 9); // Maximum 10 rows
      endCol = Math.min(endCol, startCol + 9); // Maximum 10 columns
    }
    
    // Read data
    const data = [];
    const headers = [];
    
    // Get header row (first row)
    for (let col = startCol; col <= endCol; col++) {
      const cell = worksheet.getCell(startRow, col);
      headers.push(cell.value);
    }
    
    // Get data rows
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
      
      // Skip empty rows
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
      error instanceof Error ? `Data reading error: ${error.message}` : 'Unknown error occurred while reading data'
    );
  }
}

/**
 * Write data to Excel file
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleWriteExcel(args: WriteExcelArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, data, startCell = 'A1', writeHeaders = true } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Get start cell coordinates
    const { row: startRow, col: startCol } = parseCellOrRange(startCell).startRow
      ? { row: parseCellOrRange(startCell).startRow, col: parseCellOrRange(startCell).startCol }
      : { row: 1, col: 1 };
    
    // If data is an array of arrays (2D array)
    if (Array.isArray(data) && data.every(item => Array.isArray(item))) {
      for (let i = 0; i < data.length; i++) {
        const rowData = data[i];
        for (let j = 0; j < rowData.length; j++) {
          worksheet.getCell(startRow + i, startCol + j).value = rowData[j];
        }
      }
    }
    // If data is an array of objects
    else if (Array.isArray(data) && data.every(item => typeof item === 'object' && item !== null)) {
      // Write headers
      if (writeHeaders && data.length > 0) {
        const headers = Object.keys(data[0]);
        for (let i = 0; i < headers.length; i++) {
          worksheet.getCell(startRow, startCol + i).value = headers[i];
        }
      }
      
      // Write data
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
        'Data must be an array of arrays or an array of objects'
      );
    }
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: 'Data saved successfully'
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Data writing error: ${error.message}` : 'Unknown error occurred while writing data'
    );
  }
}
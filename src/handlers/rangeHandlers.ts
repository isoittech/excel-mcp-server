/**
 * Cell range operation handlers
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
 * Copy cell range
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleCopyRange(args: CopyRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, sourceStart, sourceEnd, targetStart, targetSheet } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const sourceWorksheet = getWorksheet(workbook, sheetName);
    const targetWorksheet = targetSheet ? getWorksheet(workbook, targetSheet) : sourceWorksheet;
    
    // Parse range
    const sourceRange = parseCellOrRange(`${sourceStart}:${sourceEnd}`);
    const targetStartCoord = cellRefToCoordinate(targetStart);
    
    // Calculate source range size
    const rowCount = sourceRange.endRow - sourceRange.startRow + 1;
    const colCount = sourceRange.endCol - sourceRange.startCol + 1;
    
    // Copy cell values and formatting
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
        
        // Copy value
        targetCell.value = sourceCell.value;
        
        // Copy formatting
        targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        
        // Copy number format
        if (sourceCell.numFmt) {
          targetCell.numFmt = sourceCell.numFmt;
        }
      }
    }
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Range ${sourceStart}:${sourceEnd} successfully copied to ${targetStart} in sheet ${targetSheet || sheetName}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Range copy error: ${error.message}` : 'Unknown error occurred while copying range'
    );
  }
}

/**
 * Delete cell range
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleDeleteRange(args: DeleteRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell, shiftDirection = 'up' } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse range
    const range = parseCellOrRange(`${startCell}:${endCell}`);
    
    // Clear cells in range
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = worksheet.getCell(row, col);
        cell.value = null;
        
        // Clear styles
        cell.style = {};
        cell.numFmt = '';
      }
    }
    
    // Shift cells
    if (shiftDirection === 'up') {
      // Shift up
      // Move cells below the deleted range upward
      const colCount = range.endCol - range.startCol + 1;
      const rowCount = range.endRow - range.startRow + 1;
      
      for (let col = range.startCol; col <= range.endCol; col++) {
        for (let row = range.endRow + 1; row <= worksheet.rowCount; row++) {
          const sourceCell = worksheet.getCell(row, col);
          const targetCell = worksheet.getCell(row - rowCount, col);
          
          // Copy values and formatting
          targetCell.value = sourceCell.value;
          targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
          if (sourceCell.numFmt) {
            targetCell.numFmt = sourceCell.numFmt;
          }
          
          // Clear original cell
          sourceCell.value = null;
          sourceCell.style = {};
          sourceCell.numFmt = '';
        }
      }
    } else if (shiftDirection === 'left') {
      // Shift left
      // Move cells to the right of the deleted range leftward
      const colCount = range.endCol - range.startCol + 1;
      
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.endCol + 1; col <= worksheet.columnCount; col++) {
          const sourceCell = worksheet.getCell(row, col);
          const targetCell = worksheet.getCell(row, col - colCount);
          
          // Copy values and formatting
          targetCell.value = sourceCell.value;
          targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
          if (sourceCell.numFmt) {
            targetCell.numFmt = sourceCell.numFmt;
          }
          
          // Clear original cell
          sourceCell.value = null;
          sourceCell.style = {};
          sourceCell.numFmt = '';
        }
      }
    }
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Range ${startCell}:${endCell} successfully deleted and cells shifted ${shiftDirection === 'up' ? 'up' : 'left'}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Range deletion error: ${error.message}` : 'Unknown error occurred while deleting range'
    );
  }
}

/**
 * Validate Excel range
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleValidateExcelRange(args: ValidateExcelRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse range
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
          text: `Invalid range format: ${error instanceof Error ? error.message : 'Unknown error'}`
        }]
      };
    }
    
    // Check if range is within worksheet boundaries
    const maxRow = worksheet.rowCount || 1048576; // Excel maximum rows
    const maxCol = worksheet.columnCount || 16384; // Excel maximum columns
    
    if (range.startRow < 1 || range.startCol < 1 || range.endRow > maxRow || range.endCol > maxCol) {
      return {
        content: [{
          type: 'text',
          text: `Range is outside worksheet boundaries: rows(${range.startRow}-${range.endRow}), columns(${range.startCol}-${range.endCol})`
        }]
      };
    }
    
    // Get data information within range
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
        text: `Range ${startCell}${endCell ? `:${endCell}` : ''} is valid. ` +
              `Cell count: ${cellCount}, Cells with data: ${nonEmptyCellCount}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Range validation error: ${error.message}` : 'Unknown error occurred during range validation'
    );
  }
}
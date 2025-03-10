/**
 * Worksheet operation handlers
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
 * Rename worksheet
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleRenameWorksheet(args: RenameWorksheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, oldName, newName } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // Check if original sheet exists
    const worksheet = workbook.getWorksheet(oldName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${oldName} not found`);
    }
    
    // Check if sheet with new name already exists
    if (workbook.getWorksheet(newName)) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${newName} already exists`);
    }
    
    // Rename sheet
    worksheet.name = newName;
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Worksheet successfully renamed from ${oldName} to ${newName}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Sheet rename error: ${error.message}` : 'Unknown error occurred while renaming sheet'
    );
  }
}

/**
 * Delete worksheet
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleDeleteWorksheet(args: DeleteWorksheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // Check if sheet exists
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} not found`);
    }
    
    // Check if workbook has multiple sheets
    if (workbook.worksheets.length === 1) {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Cannot delete the only sheet in the workbook'
      );
    }
    
    // Delete sheet
    workbook.removeWorksheet(worksheet.id);
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Worksheet ${sheetName} successfully deleted`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Sheet deletion error: ${error.message}` : 'Unknown error occurred while deleting sheet'
    );
  }
}

/**
 * Copy worksheet
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleCopyWorksheet(args: CopyWorksheetArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sourceSheet, targetSheet } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    
    // Check if source sheet exists
    const sourceWorksheet = workbook.getWorksheet(sourceSheet);
    if (!sourceWorksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Source sheet ${sourceSheet} not found`);
    }
    
    // Check if target sheet name already exists
    if (workbook.getWorksheet(targetSheet)) {
      throw new McpError(ErrorCode.InvalidParams, `Target sheet ${targetSheet} already exists`);
    }
    
    // Create new sheet
    const targetWorksheet = workbook.addWorksheet(targetSheet);
    
    // Copy properties
    targetWorksheet.properties = JSON.parse(JSON.stringify(sourceWorksheet.properties));
    targetWorksheet.properties.tabColor = sourceWorksheet.properties.tabColor;
    
    // Copy column properties and widths
    sourceWorksheet.columns.forEach((column, index) => {
      if (column.width) {
        targetWorksheet.getColumn(index + 1).width = column.width;
      }
    });
    
    // Copy row heights and values
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
    
    // Copy merged cells
    // Direct access to merged cells is difficult in ExcelJS,
    // so as an alternative, examine each cell to detect merged cells
    sourceWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (cell.isMerged) {
          // If this cell is part of a merged cell,
          // identify the top-left cell of the merge range and apply the merge
          const master = cell.master;
          if (master.address === cell.address) {
            // If this cell is the top-left (master) cell of the merge
            // Identify and apply the merge range
            // Note: In a real implementation, a method to accurately identify the merge range is needed
            // This is a simplified implementation
            let endRow = rowNumber;
            let endCol = colNumber;
            
            // Estimate merge range (a more accurate method would be needed in a real implementation)
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
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Worksheet ${sourceSheet} successfully copied to ${targetSheet}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Sheet copy error: ${error.message}` : 'Unknown error occurred while copying sheet'
    );
  }
}
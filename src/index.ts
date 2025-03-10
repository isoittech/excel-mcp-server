#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListResourcesRequestSchema,
  ListResourceTemplatesRequestSchema,
  ListToolsRequestSchema,
  McpError,
  ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

class ExcelServer {
  private server: Server;
  private workbookCache: Map<string, ExcelJS.Workbook> = new Map();

  constructor() {
    this.server = new Server(
      {
        name: 'excel-server',
        version: '0.1.0',
      },
      {
        capabilities: {
          resources: {},
          tools: {},
        },
      }
    );

    this.setupToolHandlers();
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'read_excel',
          description: 'Read data from Excel file',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              range: { type: 'string' }
            },
            required: ['filePath']
          }
        },
        {
          name: 'write_excel',
          description: 'Write data to Excel file',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              data: { type: 'array' }
            },
            required: ['filePath', 'data']
          }
        },
        {
          name: 'create_sheet',
          description: 'Create new sheet in Excel file',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' }
            },
            required: ['filePath', 'sheetName']
          }
        },
        {
          name: 'create_excel',
          description: 'Create new Excel file',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string', default: 'Sheet1' }
            },
            required: ['filePath']
          }
        },
        {
          name: 'get_workbook_metadata',
          description: 'Get metadata about workbook including sheets, ranges, etc.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              includeRanges: { type: 'boolean', default: false }
            },
            required: ['filePath']
          }
        },
        {
          name: 'rename_worksheet',
          description: 'Rename worksheet in workbook',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              oldName: { type: 'string' },
              newName: { type: 'string' }
            },
            required: ['filePath', 'oldName', 'newName']
          }
        },
        {
          name: 'delete_worksheet',
          description: 'Delete worksheet from workbook',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' }
            },
            required: ['filePath', 'sheetName']
          }
        },
        {
          name: 'copy_worksheet',
          description: 'Copy worksheet within workbook',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sourceSheet: { type: 'string' },
              targetSheet: { type: 'string' }
            },
            required: ['filePath', 'sourceSheet', 'targetSheet']
          }
        }
      ]
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          case 'read_excel':
            const readExcelResult = await this.handleReadExcel(request.params.arguments);
            return readExcelResult;
          case 'write_excel':
            const writeExcelResult = await this.handleWriteExcel(request.params.arguments);
            return writeExcelResult;
          case 'create_sheet':
            const createSheetResult = await this.handleCreateSheet(request.params.arguments);
            return createSheetResult;
          case 'create_excel':
            const createExcelResult = await this.handleCreateExcel(request.params.arguments);
            return createExcelResult;
          case 'get_workbook_metadata':
            const getWorkbookMetadataResult = await this.handleGetWorkbookMetadata(request.params.arguments);
            return getWorkbookMetadataResult;
          case 'rename_worksheet':
            const renameWorksheetResult = await this.handleRenameWorksheet(request.params.arguments);
            return renameWorksheetResult;
          case 'delete_worksheet':
            const deleteWorksheetResult = await this.handleDeleteWorksheet(request.params.arguments);
            return deleteWorksheetResult;
          case 'copy_worksheet':
            const copyWorksheetResult = await this.handleCopyWorksheet(request.params.arguments);
            return copyWorksheetResult;
          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(ErrorCode.InternalError, error instanceof Error ? error.message : 'Unknown error');
      }
    });
  }

  /**
   * Reads data from an Excel file
   * @param args - Object containing:
   *   - filePath: string - Path to the Excel file
   *   - sheetName?: string - Name of the sheet to read (default: first sheet)
   *   - range?: string - Range to read (e.g., "A1:C10")
   * @returns Object containing the read data in JSON format
   */
  private async handleReadExcel(args: any) {
    const { filePath, sheetName, range } = args;
    const workbook = await this.loadWorkbook(filePath);
    const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
    
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} not found`);
    }

    const data = [];
    const [startCol, startRow, endCol, endRow] = range ? this.parseRange(range) : [1, 1, worksheet.columnCount, worksheet.rowCount];

    for (let row = startRow; row <= endRow; row++) {
      const rowData = [];
      for (let col = startCol; col <= endCol; col++) {
        const cell = worksheet.getCell(row, col);
        rowData.push(cell.value);
      }
      data.push(rowData);
    }

    return {
      content: [{
        type: 'text',
        text: JSON.stringify(data, null, 2)
      }]
    };
  }

  private async handleWriteExcel(args: any) {
    const { filePath, sheetName, data } = args;
    const workbook = await this.loadWorkbook(filePath);
    const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
    
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} not found`);
    }

    for (let row = 0; row < data.length; row++) {
      for (let col = 0; col < data[row].length; col++) {
        worksheet.getCell(row + 1, col + 1).value = data[row][col];
      }
    }

    await workbook.xlsx.writeFile(filePath);
    return {
      content: [{
        type: 'text',
        text: 'File saved successfully'
      }]
    };
  }

  /**
   * Creates a new sheet in an existing Excel file
   * @param args - Object containing:
   *   - filePath: string - Path to the Excel file
   *   - sheetName: string - Name of the new sheet
   * @returns Object with success message
   */
  private async handleCreateSheet(args: any) {
    const { filePath, sheetName } = args;
    const workbook = await this.loadWorkbook(filePath);
    
    if (workbook.getWorksheet(sheetName)) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} already exists`);
    }

    workbook.addWorksheet(sheetName);
    await workbook.xlsx.writeFile(filePath);
    return {
      content: [{
        type: 'text',
        text: `Sheet ${sheetName} created successfully`
      }]
    };
  }

  /**
   * Creates a new Excel file
   * @param args - Object containing:
   *   - filePath: string - Path to create the new Excel file
   *   - sheetName?: string - Name of the sheet (default: "Sheet1")
   * @returns Object with success message
   */
  private async handleCreateExcel(args: any) {
    const { filePath, sheetName = 'Sheet1' } = args;

    if (fs.existsSync(filePath)) {
      throw new McpError(ErrorCode.InvalidParams, `File ${filePath} already exists`);
    }

    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet(sheetName);
    await workbook.xlsx.writeFile(filePath);
    
    // Add to cache
    this.workbookCache.set(filePath, workbook);
    
    return {
      content: [{
        type: 'text',
        text: `Excel file created successfully at ${filePath}`
      }]
    };
  }

  /**
   * Loads an Excel workbook from file
   * @param filePath - Path to the Excel file
   * @returns Promise resolving to the loaded workbook
   * @throws McpError if file not found
   */
  private async loadWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
    if (this.workbookCache.has(filePath)) {
      return this.workbookCache.get(filePath)!;
    }

    if (!fs.existsSync(filePath)) {
      throw new McpError(ErrorCode.InvalidParams, `File ${filePath} not found`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    this.workbookCache.set(filePath, workbook);
    return workbook;
  }

  /**
   * Parses Excel range string into numeric coordinates
   * @param range - Excel range string (e.g., "A1:C10")
   * @returns Array containing [startCol, startRow, endCol, endRow]
   * @throws McpError if range format is invalid
   */
  private parseRange(range: string): [number, number, number, number] {
    const rangeParts = range.split(':');
    if (rangeParts.length !== 2) {
      throw new McpError(ErrorCode.InvalidParams, 'Invalid range format. Expected format like "A1:C10"');
    }

    const [startCell, endCell] = rangeParts;

    // Get start cell column and row
    const startColMatch = startCell.match(/[A-Za-z]+/);
    const startRowMatch = startCell.match(/\d+/);
    if (!startColMatch || !startRowMatch) {
      throw new McpError(ErrorCode.InvalidParams, 'Invalid start cell format');
    }
    const startCol = this.columnLetterToNumber(startColMatch[0]);
    const startRow = parseInt(startRowMatch[0], 10);

    // Get end cell column and row
    const endColMatch = endCell.match(/[A-Za-z]+/);
    const endRowMatch = endCell.match(/\d+/);
    if (!endColMatch || !endRowMatch) {
      throw new McpError(ErrorCode.InvalidParams, 'Invalid end cell format');
    }
    const endCol = this.columnLetterToNumber(endColMatch[0]);
    const endRow = parseInt(endRowMatch[0], 10);

    return [startCol, startRow, endCol, endRow];
  }

  /**
   * Converts Excel column letters to numeric index
   * @param letters - Column letters (e.g., "A", "B", "AA")
   * @returns Numeric column index (1-based)
   */
  private columnLetterToNumber(letters: string): number {
    let column = 0;
    letters = letters.toUpperCase();
    for (let i = 0; i < letters.length; i++) {
      column = column * 26 + (letters.charCodeAt(i) - 64);
    }
    return column;
  }

  /**
   * Gets metadata about a workbook including sheets, ranges, etc.
   * @param args - Object containing:
   *   - filePath: string - Path to the Excel file
   *   - includeRanges?: boolean - Whether to include used ranges info (default: false)
   * @returns Object with workbook metadata
   */
  private async handleGetWorkbookMetadata(args: any) {
    const { filePath, includeRanges = false } = args;
    const workbook = await this.loadWorkbook(filePath);
    
    const metadata: any = {
      fileName: path.basename(filePath),
      filePath: filePath,
      sheets: []
    };

    // Get information about each worksheet
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
        // Get the used range by finding the last cell with content
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
  }

  /**
   * Renames a worksheet in a workbook
   * @param args - Object containing:
   *   - filePath: string - Path to the Excel file
   *   - oldName: string - Current name of the worksheet
   *   - newName: string - New name for the worksheet
   * @returns Object with success message
   */
  private async handleRenameWorksheet(args: any) {
    const { filePath, oldName, newName } = args;
    const workbook = await this.loadWorkbook(filePath);
    
    const worksheet = workbook.getWorksheet(oldName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${oldName} not found`);
    }

    // Check if new name already exists
    if (workbook.getWorksheet(newName)) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${newName} already exists`);
    }

    worksheet.name = newName;
    await workbook.xlsx.writeFile(filePath);
    
    return {
      content: [{
        type: 'text',
        text: `Worksheet renamed from ${oldName} to ${newName} successfully`
      }]
    };
  }

  /**
   * Deletes a worksheet from a workbook
   * @param args - Object containing:
   *   - filePath: string - Path to the Excel file
   *   - sheetName: string - Name of the worksheet to delete
   * @returns Object with success message
   */
  private async handleDeleteWorksheet(args: any) {
    const { filePath, sheetName } = args;
    const workbook = await this.loadWorkbook(filePath);
    
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} not found`);
    }

    // Ensure we're not deleting the only worksheet
    if (workbook.worksheets.length === 1) {
      throw new McpError(
        ErrorCode.InvalidParams,
        'Cannot delete the only worksheet in the workbook'
      );
    }

    workbook.removeWorksheet(worksheet.id);
    await workbook.xlsx.writeFile(filePath);
    
    return {
      content: [{
        type: 'text',
        text: `Worksheet ${sheetName} deleted successfully`
      }]
    };
  }

  /**
   * Copies a worksheet within a workbook
   * @param args - Object containing:
   *   - filePath: string - Path to the Excel file
   *   - sourceSheet: string - Name of the source worksheet
   *   - targetSheet: string - Name for the new worksheet
   * @returns Object with success message
   */
  private async handleCopyWorksheet(args: any) {
    const { filePath, sourceSheet, targetSheet } = args;
    const workbook = await this.loadWorkbook(filePath);
    
    const sourceWorksheet = workbook.getWorksheet(sourceSheet);
    if (!sourceWorksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Source sheet ${sourceSheet} not found`);
    }

    // Check if target sheet name already exists
    if (workbook.getWorksheet(targetSheet)) {
      throw new McpError(ErrorCode.InvalidParams, `Target sheet ${targetSheet} already exists`);
    }

    // Create new worksheet
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
        
        // Note: ExcelJS doesn't provide a direct way to check if a cell is part of a merge group
        // We'll skip the merge cell handling as it would require additional tracking
        // In a production implementation, you might want to track merged ranges separately
      });
      
      targetRow.commit();
    });
    
    await workbook.xlsx.writeFile(filePath);
    
    return {
      content: [{
        type: 'text',
        text: `Worksheet ${sourceSheet} copied to ${targetSheet} successfully`
      }]
    };
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
  }
}

const server = new ExcelServer();
server.run().catch(console.error);
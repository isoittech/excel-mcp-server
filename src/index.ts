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

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
  }
}

const server = new ExcelServer();
server.run().catch(console.error);
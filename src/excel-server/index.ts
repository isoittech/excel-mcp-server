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
    console.error('Starting Excel MCP server...'); // 追加: 起動ログ出力
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
        }
      ]
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          case 'read_excel':
            console.error('Handling read_excel tool...'); // 追加: ツールハンドラー開始ログ
            const readExcelResult = await this.handleReadExcel(request.params.arguments);
            console.error('read_excel tool handled.'); // 追加: ツールハンドラー終了ログ
            return readExcelResult;
          case 'write_excel':
            console.error('Handling write_excel tool...'); // 追加: ツールハンドラー開始ログ
            const writeExcelResult = await this.handleWriteExcel(request.params.arguments);
            console.error('write_excel tool handled.'); // 追加: ツールハンドラー終了ログ
            return writeExcelResult;
          case 'create_sheet':
            console.error('Handling create_sheet tool...'); // 追加: ツールハンドラー開始ログ
            const createSheetResult = await this.handleCreateSheet(request.params.arguments);
            console.error('create_sheet tool handled.'); // 追加: ツールハンドラー終了ログ
            return createSheetResult;
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

  private async handleReadExcel(args: any) {
    console.error('handleReadExcel started...'); // 追加: handleReadExcel 開始ログ
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
    console.error('handleReadExcel finished.'); // 追加: handleReadExcel 終了ログ

    return {
      content: [{
        type: 'text',
        text: JSON.stringify(data, null, 2)
      }]
    };
  }

  private async handleWriteExcel(args: any) {
    console.error('handleWriteExcel started...'); // 追加: handleWriteExcel 開始ログ
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
    console.error('handleWriteExcel finished.'); // 追加: handleWriteExcel 終了ログ
    return {
      content: [{
        type: 'text',
        text: 'File saved successfully'
      }]
    };
  }

  private async handleCreateSheet(args: any) {
    console.error('handleCreateSheet started...'); // 追加: handleCreateSheet 開始ログ
    const { filePath, sheetName } = args;
    const workbook = await this.loadWorkbook(filePath);
    
    if (workbook.getWorksheet(sheetName)) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} already exists`);
    }

    workbook.addWorksheet(sheetName);
    await workbook.xlsx.writeFile(filePath);
    console.error('handleCreateSheet finished.'); // 追加: handleCreateSheet 終了ログ
    return {
      content: [{
        type: 'text',
        text: `Sheet ${sheetName} created successfully`
      }]
    };
  }

  private async loadWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
    console.error('loadWorkbook started...', filePath); // 追加: loadWorkbook 開始ログ
    if (this.workbookCache.has(filePath)) {
      console.error('loadWorkbook finished - cache hit.', filePath); // 追加: loadWorkbook 終了ログ (キャッシュヒット)
      return this.workbookCache.get(filePath)!;
    }

    if (!fs.existsSync(filePath)) {
      console.error('loadWorkbook finished - file not found.', filePath); // 追加: loadWorkbook 終了ログ (ファイルNotFound)
      throw new McpError(ErrorCode.InvalidParams, `File ${filePath} not found`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    this.workbookCache.set(filePath, workbook);
    console.error('loadWorkbook finished - file read.', filePath); // 追加: loadWorkbook 終了ログ (ファイル読み込み)
    return workbook;
  }

  private parseRange(range: string): [number, number, number, number] {
    // Implement range parsing logic
    return [1, 1, 10, 10]; // Placeholder
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Transport connected.'); // 追加: 接続ログ出力
    console.error('Excel MCP server running on stdio');
  }
}

const server = new ExcelServer();
server.run().catch(console.error);
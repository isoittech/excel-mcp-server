/**
 * Excel MCP Server Class
 */

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
import { WorkbookCache } from './types/index.js';

// Import handlers
import { handleReadExcel, handleWriteExcel } from './handlers/dataHandlers.js';
import { handleCreateExcel, handleCreateSheet, handleGetWorkbookMetadata } from './handlers/workbookHandlers.js';
import { handleRenameWorksheet, handleDeleteWorksheet, handleCopyWorksheet } from './handlers/worksheetHandlers.js';
import { handleApplyFormula, handleValidateFormulaSyntax } from './handlers/formulaHandlers.js';
import { handleFormatRange, handleMergeCells, handleUnmergeCells } from './handlers/formatHandlers.js';
import { handleCopyRange, handleDeleteRange, handleValidateExcelRange } from './handlers/rangeHandlers.js';
import { handleCreateChart } from './handlers/chartHandlers.js';
import { handleCreatePivotTable } from './handlers/pivotHandlers.js';

/**
 * Excel MCP Server Class
 */
export class ExcelServer {
  private server: Server;
  private workbookCache: WorkbookCache = {};

  /**
   * Constructor
   */
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

  /**
   * Setup tool handlers
   */
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
              range: { type: 'string' },
              previewOnly: { type: 'boolean', default: false }
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
              data: { type: 'array' },
              startCell: { type: 'string', default: 'A1' },
              writeHeaders: { type: 'boolean', default: true }
            },
            required: ['filePath', 'data']
          }
        },
        {
          name: 'create_sheet',
          description: 'Create a new sheet in Excel file',
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
          description: 'Create a new Excel file',
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
          description: 'Get workbook metadata',
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
          description: 'Rename worksheet',
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
          description: 'Delete worksheet',
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
          description: 'Copy worksheet',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sourceSheet: { type: 'string' },
              targetSheet: { type: 'string' }
            },
            required: ['filePath', 'sourceSheet', 'targetSheet']
          }
        },
        {
          name: 'apply_formula',
          description: 'Apply formula to cell',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              cell: { type: 'string' },
              formula: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'cell', 'formula']
          }
        },
        {
          name: 'validate_formula_syntax',
          description: 'Validate formula syntax',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              cell: { type: 'string' },
              formula: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'cell', 'formula']
          }
        },
        {
          name: 'format_range',
          description: 'Format cell range',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              startCell: { type: 'string' },
              endCell: { type: 'string' },
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean' },
              fontSize: { type: 'number' },
              fontColor: { type: 'string' },
              bgColor: { type: 'string' },
              borderStyle: { type: 'string' },
              borderColor: { type: 'string' },
              numberFormat: { type: 'string' },
              alignment: { type: 'string' },
              wrapText: { type: 'boolean' },
              mergeCells: { type: 'boolean' }
            },
            required: ['filePath', 'sheetName', 'startCell']
          }
        },
        {
          name: 'merge_cells',
          description: 'Merge cells',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              startCell: { type: 'string' },
              endCell: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'startCell', 'endCell']
          }
        },
        {
          name: 'unmerge_cells',
          description: 'Unmerge cells',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              startCell: { type: 'string' },
              endCell: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'startCell', 'endCell']
          }
        },
        {
          name: 'copy_range',
          description: 'Copy cell range',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              sourceStart: { type: 'string' },
              sourceEnd: { type: 'string' },
              targetStart: { type: 'string' },
              targetSheet: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'sourceStart', 'sourceEnd', 'targetStart']
          }
        },
        {
          name: 'delete_range',
          description: 'Delete cell range',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              startCell: { type: 'string' },
              endCell: { type: 'string' },
              shiftDirection: { type: 'string', enum: ['up', 'left'], default: 'up' }
            },
            required: ['filePath', 'sheetName', 'startCell', 'endCell']
          }
        },
        {
          name: 'validate_excel_range',
          description: 'Validate Excel range',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              startCell: { type: 'string' },
              endCell: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'startCell']
          }
        },
        {
          name: 'create_chart',
          description: 'Create chart',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              dataRange: { type: 'string' },
              chartType: { type: 'string' },
              targetCell: { type: 'string' },
              title: { type: 'string' },
              xAxis: { type: 'string' },
              yAxis: { type: 'string' }
            },
            required: ['filePath', 'sheetName', 'dataRange', 'chartType', 'targetCell']
          }
        },
        {
          name: 'create_pivot_table',
          description: 'Create pivot table',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string' },
              sheetName: { type: 'string' },
              dataRange: { type: 'string' },
              rows: { type: 'array', items: { type: 'string' } },
              values: { type: 'array', items: { type: 'string' } },
              columns: { type: 'array', items: { type: 'string' } },
              aggFunc: { type: 'string', default: 'sum' }
            },
            required: ['filePath', 'sheetName', 'dataRange', 'rows', 'values']
          }
        }
      ]
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request): Promise<any> => {
      try {
        const args = request.params.arguments || {};
        
        switch (request.params.name) {
          case 'read_excel':
            return await handleReadExcel(args as any, this.workbookCache);
          case 'write_excel':
            return await handleWriteExcel(args as any, this.workbookCache);
          case 'create_sheet':
            return await handleCreateSheet(args as any, this.workbookCache);
          case 'create_excel':
            return await handleCreateExcel(args as any, this.workbookCache);
          case 'get_workbook_metadata':
            return await handleGetWorkbookMetadata(args as any, this.workbookCache);
          case 'rename_worksheet':
            return await handleRenameWorksheet(args as any, this.workbookCache);
          case 'delete_worksheet':
            return await handleDeleteWorksheet(args as any, this.workbookCache);
          case 'copy_worksheet':
            return await handleCopyWorksheet(args as any, this.workbookCache);
          case 'apply_formula':
            return await handleApplyFormula(args as any, this.workbookCache);
          case 'validate_formula_syntax':
            return await handleValidateFormulaSyntax(args as any, this.workbookCache);
          case 'format_range':
            return await handleFormatRange(args as any, this.workbookCache);
          case 'merge_cells':
            return await handleMergeCells(args as any, this.workbookCache);
          case 'unmerge_cells':
            return await handleUnmergeCells(args as any, this.workbookCache);
          case 'copy_range':
            return await handleCopyRange(args as any, this.workbookCache);
          case 'delete_range':
            return await handleDeleteRange(args as any, this.workbookCache);
          case 'validate_excel_range':
            return await handleValidateExcelRange(args as any, this.workbookCache);
          case 'create_chart':
            return await handleCreateChart(args as any, this.workbookCache);
          case 'create_pivot_table':
            return await handleCreatePivotTable(args as any, this.workbookCache);
          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
        }
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          error instanceof Error ? error.message : 'Unknown error'
        );
      }
    });
  }

  /**
   * Run the server
   */
  async run() {
    try {
      const transport = new StdioServerTransport();
      await this.server.connect(transport);
      console.error('Excel MCP server running on stdio');
    } catch (error) {
      console.error('Server error:', error);
      process.exit(1);
    }
  }
}
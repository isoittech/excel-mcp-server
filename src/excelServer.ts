/**
 * Excel MCP Server Class
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
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
        name: 'excel-mcp-server',
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path, e.g. 'excel_files/sample-data.xlsx'."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet to read. If omitted, the first worksheet is used. Example: 'Sheet1'."
              },
              range: {
                type: 'string',
                description: "Cell range to read in A1 notation. Example: 'A1:D10'. If omitted, reads the entire worksheet."
              },
              previewOnly: {
                type: 'boolean',
                default: false,
                description: "If true, returns only a preview (first few rows) of the data. If false, returns all data. Default: false."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file to write. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet to write to. If omitted, the first worksheet is used."
              },
              data: {
                type: 'array',
                description: "Array of data to write. Each element should represent a row (as an array or object)."
              },
              startCell: {
                type: 'string',
                default: 'A1',
                description: "Top-left cell where writing starts, in A1 notation. Example: 'B2'. Default: 'A1'."
              },
              writeHeaders: {
                type: 'boolean',
                default: true,
                description: "If true, writes column headers as the first row. If false, writes data only. Default: true."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the new worksheet to create. Example: 'Summary'."
              }
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
              filePath: {
                type: 'string',
                description: "Path for the new Excel file to create. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                default: 'Sheet1',
                description: "Name of the first worksheet to create in the new file. Default: 'Sheet1'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              includeRanges: {
                type: 'boolean',
                default: false,
                description: "If true, includes cell range information in the metadata. Default: false."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              oldName: {
                type: 'string',
                description: "Current name of the worksheet to rename. Example: 'Sheet1'."
              },
              newName: {
                type: 'string',
                description: "New name for the worksheet. Example: 'Summary'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet to delete. Example: 'Sheet2'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sourceSheet: {
                type: 'string',
                description: "Name of the worksheet to copy. Example: 'Sheet1'."
              },
              targetSheet: {
                type: 'string',
                description: "Name for the new copied worksheet. Example: 'Sheet1_Copy'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              cell: {
                type: 'string',
                description: "Cell address to apply the formula, in A1 notation. Example: 'C5'."
              },
              formula: {
                type: 'string',
                description: "Formula to apply to the cell. Example: '=SUM(A1:A10)'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              cell: {
                type: 'string',
                description: "Cell address to validate the formula, in A1 notation. Example: 'B2'."
              },
              formula: {
                type: 'string',
                description: "Formula to validate. Example: '=AVERAGE(B2:B10)'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              startCell: {
                type: 'string',
                description: "Top-left cell of the range to format, in A1 notation. Example: 'A1'."
              },
              endCell: {
                type: 'string',
                description: "Bottom-right cell of the range to format, in A1 notation. Example: 'D10'."
              },
              bold: {
                type: 'boolean',
                description: "If true, applies bold formatting to the range."
              },
              italic: {
                type: 'boolean',
                description: "If true, applies italic formatting to the range."
              },
              underline: {
                type: 'boolean',
                description: "If true, applies underline formatting to the range."
              },
              fontSize: {
                type: 'number',
                description: "Font size to apply to the range. Example: 12."
              },
              fontColor: {
                type: 'string',
                description: "Font color in hex or named color. Example: '#FF0000' or 'red'."
              },
              bgColor: {
                type: 'string',
                description: "Background color in hex or named color. Example: '#FFFF00' or 'yellow'."
              },
              borderStyle: {
                type: 'string',
                description: "Border style to apply. Example: 'thin', 'medium', 'dashed'."
              },
              borderColor: {
                type: 'string',
                description: "Border color in hex or named color. Example: '#000000' or 'black'."
              },
              numberFormat: {
                type: 'string',
                description: "Number format string. Example: '0.00', 'yyyy-mm-dd'."
              },
              alignment: {
                type: 'string',
                description: "Text alignment. Example: 'left', 'center', 'right'."
              },
              wrapText: {
                type: 'boolean',
                description: "If true, enables text wrapping in the range."
              },
              mergeCells: {
                type: 'boolean',
                description: "If true, merges the specified cell range."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              startCell: {
                type: 'string',
                description: "Top-left cell of the range to merge, in A1 notation. Example: 'A1'."
              },
              endCell: {
                type: 'string',
                description: "Bottom-right cell of the range to merge, in A1 notation. Example: 'B2'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              startCell: {
                type: 'string',
                description: "Top-left cell of the range to unmerge, in A1 notation. Example: 'A1'."
              },
              endCell: {
                type: 'string',
                description: "Bottom-right cell of the range to unmerge, in A1 notation. Example: 'B2'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet to copy from. Example: 'Sheet1'."
              },
              sourceStart: {
                type: 'string',
                description: "Top-left cell of the source range, in A1 notation. Example: 'A1'."
              },
              sourceEnd: {
                type: 'string',
                description: "Bottom-right cell of the source range, in A1 notation. Example: 'C10'."
              },
              targetStart: {
                type: 'string',
                description: "Top-left cell of the target range, in A1 notation. Example: 'E1'."
              },
              targetSheet: {
                type: 'string',
                description: "Name of the worksheet to copy to. If omitted, uses the same sheet."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              startCell: {
                type: 'string',
                description: "Top-left cell of the range to delete, in A1 notation. Example: 'A1'."
              },
              endCell: {
                type: 'string',
                description: "Bottom-right cell of the range to delete, in A1 notation. Example: 'B5'."
              },
              shiftDirection: {
                type: 'string',
                enum: ['up', 'left'],
                default: 'up',
                description: "Direction to shift remaining cells after deletion. 'up' shifts cells up, 'left' shifts cells left. Default: 'up'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet. Example: 'Sheet1'."
              },
              startCell: {
                type: 'string',
                description: "Top-left cell of the range to validate, in A1 notation. Example: 'A1'."
              },
              endCell: {
                type: 'string',
                description: "Bottom-right cell of the range to validate, in A1 notation. Example: 'C10'."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet to create the chart in. Example: 'Sheet1'."
              },
              dataRange: {
                type: 'string',
                description: "Cell range containing the data for the chart, in A1 notation. Example: 'A1:D10'."
              },
              chartType: {
                type: 'string',
                description: "Type of chart to create. Example: 'bar', 'line', 'pie'."
              },
              targetCell: {
                type: 'string',
                description: "Top-left cell where the chart will be placed, in A1 notation. Example: 'F5'."
              },
              title: {
                type: 'string',
                description: "Title of the chart. Optional."
              },
              xAxis: {
                type: 'string',
                description: "Label for the X axis. Optional."
              },
              yAxis: {
                type: 'string',
                description: "Label for the Y axis. Optional."
              }
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
              filePath: {
                type: 'string',
                description: "Path to the Excel file. Specify an absolute or workspace-relative path."
              },
              sheetName: {
                type: 'string',
                description: "Name of the worksheet to create the pivot table in. Example: 'Sheet1'."
              },
              dataRange: {
                type: 'string',
                description: "Cell range containing the source data for the pivot table, in A1 notation. Example: 'A1:F100'."
              },
              rows: {
                type: 'array',
                items: { type: 'string' },
                description: "Array of field names to use as rows in the pivot table. Example: ['Region', 'Product']."
              },
              values: {
                type: 'array',
                items: { type: 'string' },
                description: "Array of field names to aggregate as values. Example: ['Sales', 'Profit']."
              },
              columns: {
                type: 'array',
                items: { type: 'string' },
                description: "Array of field names to use as columns in the pivot table. Optional."
              },
              aggFunc: {
                type: 'string',
                default: 'sum',
                description: "Aggregation function to use for values. Example: 'sum', 'count', 'average'. Default: 'sum'."
              }
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
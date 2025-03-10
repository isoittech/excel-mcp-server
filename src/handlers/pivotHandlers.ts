/**
 * Pivot table operation handlers
 * 
 * Note: Due to limitations in ExcelJS pivot table functionality,
 * the current implementation does not actually create pivot tables but only returns success messages.
 * Consider using an improved version of ExcelJS or a different library in the future.
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { 
  CreatePivotTableArgs, 
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, getWorksheet } from '../utils/fileUtils.js';
import { parseCellOrRange } from '../utils/cellUtils.js';

/**
 * Create pivot table
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleCreatePivotTable(args: CreatePivotTableArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, dataRange, rows, values, columns, aggFunc = 'sum' } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse data range
    const range = parseCellOrRange(dataRange);
    
    // Validate aggregation function
    const validAggFunc = validateAggFunc(aggFunc);
    if (!validAggFunc.valid) {
      throw new McpError(ErrorCode.InvalidParams, validAggFunc.error || 'Invalid aggregation function');
    }
    
    // Note: Due to limitations in the current version of ExcelJS,
    // we don't actually create a pivot table but only return a success message.
    
    // For future implementation:
    // 1. Get data
    // 2. Create pivot table
    // 3. Set rows, columns, values, and aggregation function
    // 4. Add pivot table to worksheet
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Pivot table created successfully (Note: The current implementation does not actually create a pivot table)\n` +
              `- Data range: ${dataRange}\n` +
              `- Rows: ${rows.join(', ')}\n` +
              `- Columns: ${columns?.join(', ') || '(none)'}\n` +
              `- Values: ${values.join(', ')}\n` +
              `- Aggregation function: ${validAggFunc.func}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Pivot table creation error: ${error.message}` : 'Unknown error occurred while creating pivot table'
    );
  }
}

/**
 * Validate aggregation function
 * @param aggFunc - Aggregation function
 * @returns Validation result
 */
function validateAggFunc(aggFunc: string): { valid: boolean; func?: string; error?: string } {
  const validFuncs = [
    'sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'
  ];
  
  const lowerFunc = aggFunc.toLowerCase();
  
  // Exact match
  if (validFuncs.includes(lowerFunc)) {
    return { valid: true, func: lowerFunc };
  }
  
  // Partial match
  for (const func of validFuncs) {
    if (lowerFunc.includes(func)) {
      return { valid: true, func };
    }
  }
  
  // Special mappings
  const funcMap: Record<string, string> = {
    'mean': 'average',
    'avg': 'average',
    'total': 'sum',
    'add': 'sum',
    'maximum': 'max',
    'minimum': 'min',
    'multiply': 'product'
  };
  
  if (funcMap[lowerFunc]) {
    return { valid: true, func: funcMap[lowerFunc] };
  }
  
  return { 
    valid: false, 
    error: `Aggregation function "${aggFunc}" is invalid. Valid functions: ${validFuncs.join(', ')}` 
  };
}
/**
 * Formula operation handlers
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { 
  ApplyFormulaArgs, 
  ValidateFormulaSyntaxArgs, 
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, getWorksheet } from '../utils/fileUtils.js';
import { cellRefToCoordinate } from '../utils/cellUtils.js';

/**
 * Apply formula to a cell
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleApplyFormula(args: ApplyFormulaArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, cell, formula } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Validate formula
    const validationResult = validateFormula(formula);
    if (!validationResult.valid) {
      throw new McpError(ErrorCode.InvalidParams, `Formula error: ${validationResult.error}`);
    }
    
    // Convert cell reference to coordinates
    const { row, col } = cellRefToCoordinate(cell);
    
    // Apply formula
    worksheet.getCell(row, col).value = { formula: formula.startsWith('=') ? formula.substring(1) : formula };
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Formula "${formula}" successfully applied to cell ${cell}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Formula application error: ${error.message}` : 'Unknown error occurred while applying formula'
    );
  }
}

/**
 * Validate formula syntax
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleValidateFormulaSyntax(args: ValidateFormulaSyntaxArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, cell, formula } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Validate formula
    const validationResult = validateFormula(formula);
    
    if (validationResult.valid) {
      return {
        content: [{
          type: 'text',
          text: `Formula "${formula}" is valid`
        }]
      };
    } else {
      return {
        content: [{
          type: 'text',
          text: `Formula error: ${validationResult.error}`
        }]
      };
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Formula validation error: ${error.message}` : 'Unknown error occurred during formula validation'
    );
  }
}

/**
 * Validate formula
 * @param formula - Formula to validate
 * @returns Validation result
 */
function validateFormula(formula: string): { valid: boolean; error?: string } {
  // Check if formula is empty
  if (!formula || formula.trim() === '') {
    return { valid: false, error: 'Formula is empty' };
  }
  
  // Remove leading = from formula (if present)
  const formulaContent = formula.startsWith('=') ? formula.substring(1) : formula;
  
  // Basic syntax check
  try {
    // Check parentheses matching
    const openParenCount = (formulaContent.match(/\(/g) || []).length;
    const closeParenCount = (formulaContent.match(/\)/g) || []).length;
    if (openParenCount !== closeParenCount) {
      return { valid: false, error: 'Mismatched parentheses' };
    }
    
    // Check quotation marks matching
    const quoteCount = (formulaContent.match(/"/g) || []).length;
    if (quoteCount % 2 !== 0) {
      return { valid: false, error: 'Mismatched quotation marks' };
    }
    
    // Check basic operator usage
    if (/[+\-*/]$/.test(formulaContent)) {
      return { valid: false, error: 'Formula ends with an operator' };
    }
    
    if (/[+\-*/]{2,}/.test(formulaContent)) {
      return { valid: false, error: 'Consecutive operators' };
    }
    
    // Check function names
    const functionMatches = formulaContent.match(/[A-Za-z0-9_]+\(/g);
    if (functionMatches) {
      const knownFunctions = [
        'SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'AND', 'OR', 'NOT',
        'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'CONCATENATE', 'LEFT', 'RIGHT',
        'MID', 'LEN', 'FIND', 'SEARCH', 'SUBSTITUTE', 'UPPER', 'LOWER', 'PROPER',
        'TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND',
        'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'ABS', 'SQRT', 'POWER', 'MOD',
        'SUMIF', 'COUNTIF', 'AVERAGEIF', 'SUMIFS', 'COUNTIFS', 'AVERAGEIFS',
        'IFERROR', 'IFNA', 'IFS', 'SWITCH', 'CHOOSE', 'INDIRECT', 'OFFSET',
        'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'TRANSPOSE', 'UNIQUE', 'SORT',
        'FILTER', 'RANDARRAY', 'SEQUENCE', 'XLOOKUP', 'XMATCH', 'LET', 'LAMBDA'
      ];
      
      for (const match of functionMatches) {
        const funcName = match.substring(0, match.length - 1).toUpperCase();
        if (!knownFunctions.includes(funcName)) {
          // Warn about unknown function names but don't treat as error
          // ExcelJS might accept unknown functions
          console.warn(`Warning: Unknown function name "${funcName}" is used`);
        }
      }
    }
    
    // Other basic checks
    // In a real implementation, more detailed syntax checking would be needed
    
    return { valid: true };
  } catch (error) {
    return { 
      valid: false, 
      error: error instanceof Error ? error.message : 'Unknown formula error' 
    };
  }
}
/**
 * Format handlers
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  FormatRangeArgs,
  MergeCellsArgs,
  UnmergeCellsArgs,
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, getWorksheet } from '../utils/fileUtils.js';
import { cellRefToCoordinate, parseCellOrRange } from '../utils/cellUtils.js';
import ExcelJS from 'exceljs';

// ExcelJSの型定義
type BorderStyle = 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double';
type HorizontalAlignment = 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
type VerticalAlignment = 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';

/**
 * Format cell range
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleFormatRange(args: FormatRangeArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { 
      filePath, 
      sheetName, 
      startCell, 
      endCell,
      bold,
      italic,
      underline,
      fontSize,
      fontColor,
      bgColor,
      borderStyle,
      borderColor,
      numberFormat,
      alignment,
      wrapText,
      mergeCells
    } = args;
    
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse range
    let range;
    if (endCell) {
      range = parseCellOrRange(`${startCell}:${endCell}`);
    } else {
      range = parseCellOrRange(startCell);
    }
    
    // Apply formatting to cell range
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = worksheet.getCell(row, col);
        
        // Font settings
        if (!cell.font) cell.font = {};
        if (bold !== undefined) cell.font.bold = bold;
        if (italic !== undefined) cell.font.italic = italic;
        if (underline !== undefined) cell.font.underline = underline ? true : false;
        if (fontSize !== undefined) cell.font.size = fontSize;
        if (fontColor !== undefined) cell.font.color = { argb: parseColor(fontColor) };
        
        // Background color
        if (bgColor !== undefined) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: parseColor(bgColor) }
          } as ExcelJS.Fill;
        }
        
        // Borders
        if (borderStyle !== undefined) {
          const style = parseBorderStyle(borderStyle);
          const color = borderColor ? { argb: parseColor(borderColor) } : { argb: 'FF000000' };
          
          if (!cell.border) cell.border = {};
          cell.border = {
            top: { style, color },
            left: { style, color },
            bottom: { style, color },
            right: { style, color }
          };
        }
        
        // Number format
        if (numberFormat !== undefined) {
          cell.numFmt = numberFormat;
        }
        
        // Alignment
        if (alignment !== undefined || wrapText !== undefined) {
          if (!cell.alignment) cell.alignment = {};
          
          if (alignment !== undefined) {
            const { horizontal, vertical } = parseAlignment(alignment);
            if (horizontal) cell.alignment.horizontal = horizontal;
            if (vertical) cell.alignment.vertical = vertical;
          }
          
          if (wrapText !== undefined) {
            cell.alignment.wrapText = wrapText;
          }
        }
      }
    }
    
    // Merge cells
    if (mergeCells) {
      worksheet.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
    }
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Format applied successfully to range ${startCell}${endCell ? `:${endCell}` : ''}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Formatting error: ${error.message}` : 'Unknown error occurred while formatting'
    );
  }
}

/**
 * Merge cells
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleMergeCells(args: MergeCellsArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse range
    const range = parseCellOrRange(`${startCell}:${endCell}`);
    
    // Merge cells
    worksheet.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Cell range ${startCell}:${endCell} merged successfully`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Cell merge error: ${error.message}` : 'Unknown error occurred while merging cells'
    );
  }
}

/**
 * Unmerge cells
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleUnmergeCells(args: UnmergeCellsArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse range
    const range = parseCellOrRange(`${startCell}:${endCell}`);
    
    // Unmerge cells
    worksheet.unMergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
    
    // Save file
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `Cell range ${startCell}:${endCell} unmerged successfully`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Cell unmerge error: ${error.message}` : 'Unknown error occurred while unmerging cells'
    );
  }
}

/**
 * Convert color string to ARGB format
 * @param color - Color string (e.g., '#FF0000', 'red')
 * @returns Color string in ARGB format
 */
function parseColor(color: string): string {
  // If already in ARGB format
  if (/^[0-9A-Fa-f]{8}$/.test(color)) {
    return color.toUpperCase();
  }
  
  // If in #RRGGBB format
  if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
    return `FF${color.substring(1).toUpperCase()}`;
  }
  
  // If color name
  const colorMap: Record<string, string> = {
    'black': 'FF000000',
    'white': 'FFFFFFFF',
    'red': 'FFFF0000',
    'green': 'FF00FF00',
    'blue': 'FF0000FF',
    'yellow': 'FFFFFF00',
    'cyan': 'FF00FFFF',
    'magenta': 'FFFF00FF',
    'gray': 'FF808080',
    'grey': 'FF808080',
    'lightgray': 'FFD3D3D3',
    'lightgrey': 'FFD3D3D3',
    'darkgray': 'FFA9A9A9',
    'darkgrey': 'FFA9A9A9',
    'orange': 'FFFFA500',
    'purple': 'FF800080',
    'brown': 'FFA52A2A'
  };
  
  const lowerColor = color.toLowerCase();
  if (colorMap[lowerColor]) {
    return colorMap[lowerColor];
  }
  
  // Default (black)
  return 'FF000000';
}

/**
 * Parse border style
 * @param style - Border style string
 * @returns ExcelJS border style
 */
function parseBorderStyle(style: string): BorderStyle | undefined {
  const styleMap: Record<string, BorderStyle | undefined> = {
    'thin': 'thin',
    'medium': 'medium',
    'thick': 'thick',
    'dotted': 'dotted',
    'dashed': 'dashed',
    'double': 'double',
    'none': undefined
  };
  
  const lowerStyle = style.toLowerCase();
  return styleMap[lowerStyle] || 'thin';
}

/**
 * Parse alignment
 * @param alignment - Alignment string
 * @returns Horizontal and vertical alignment
 */
function parseAlignment(alignment: string): { horizontal?: HorizontalAlignment; vertical?: VerticalAlignment } {
  const result: { horizontal?: HorizontalAlignment; vertical?: VerticalAlignment } = {};
  
  const lowerAlignment = alignment.toLowerCase();
  
  // Horizontal alignment
  if (lowerAlignment.includes('left')) {
    result.horizontal = 'left';
  } else if (lowerAlignment.includes('center')) {
    result.horizontal = 'center';
  } else if (lowerAlignment.includes('right')) {
    result.horizontal = 'right';
  }
  
  // Vertical alignment
  if (lowerAlignment.includes('top')) {
    result.vertical = 'top';
  } else if (lowerAlignment.includes('middle')) {
    result.vertical = 'middle';
  } else if (lowerAlignment.includes('bottom')) {
    result.vertical = 'bottom';
  }
  
  return result;
}
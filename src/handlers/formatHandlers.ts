/**
 * 書式設定ハンドラー
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
 * セル範囲の書式を設定
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
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
    
    // 範囲を解析
    let range;
    if (endCell) {
      range = parseCellOrRange(`${startCell}:${endCell}`);
    } else {
      range = parseCellOrRange(startCell);
    }
    
    // セル範囲に書式を適用
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        const cell = worksheet.getCell(row, col);
        
        // フォント設定
        if (!cell.font) cell.font = {};
        if (bold !== undefined) cell.font.bold = bold;
        if (italic !== undefined) cell.font.italic = italic;
        if (underline !== undefined) cell.font.underline = underline ? true : false;
        if (fontSize !== undefined) cell.font.size = fontSize;
        if (fontColor !== undefined) cell.font.color = { argb: parseColor(fontColor) };
        
        // 背景色
        if (bgColor !== undefined) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: parseColor(bgColor) }
          } as ExcelJS.Fill;
        }
        
        // 罫線
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
        
        // 数値書式
        if (numberFormat !== undefined) {
          cell.numFmt = numberFormat;
        }
        
        // 配置
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
    
    // セルの結合
    if (mergeCells) {
      worksheet.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
    }
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `範囲 ${startCell}${endCell ? `:${endCell}` : ''} の書式が正常に設定されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `書式設定エラー: ${error.message}` : '書式設定中に不明なエラーが発生しました'
    );
  }
}

/**
 * セルを結合
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleMergeCells(args: MergeCellsArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 範囲を解析
    const range = parseCellOrRange(`${startCell}:${endCell}`);
    
    // セルを結合
    worksheet.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `セル範囲 ${startCell}:${endCell} が正常に結合されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `セル結合エラー: ${error.message}` : 'セル結合中に不明なエラーが発生しました'
    );
  }
}

/**
 * セルの結合を解除
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleUnmergeCells(args: UnmergeCellsArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, startCell, endCell } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 範囲を解析
    const range = parseCellOrRange(`${startCell}:${endCell}`);
    
    // セルの結合を解除
    worksheet.unMergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `セル範囲 ${startCell}:${endCell} の結合が正常に解除されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `セル結合解除エラー: ${error.message}` : 'セル結合解除中に不明なエラーが発生しました'
    );
  }
}

/**
 * 色文字列をARGB形式に変換
 * @param color - 色文字列（例：'#FF0000', 'red'）
 * @returns ARGB形式の色文字列
 */
function parseColor(color: string): string {
  // 既にARGB形式の場合
  if (/^[0-9A-Fa-f]{8}$/.test(color)) {
    return color.toUpperCase();
  }
  
  // #RRGGBB形式の場合
  if (/^#[0-9A-Fa-f]{6}$/.test(color)) {
    return `FF${color.substring(1).toUpperCase()}`;
  }
  
  // 色名の場合
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
  
  // デフォルト（黒）
  return 'FF000000';
}

/**
 * 罫線スタイルを解析
 * @param style - 罫線スタイル文字列
 * @returns ExcelJSの罫線スタイル
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
 * 配置を解析
 * @param alignment - 配置文字列
 * @returns 水平・垂直配置
 */
function parseAlignment(alignment: string): { horizontal?: HorizontalAlignment; vertical?: VerticalAlignment } {
  const result: { horizontal?: HorizontalAlignment; vertical?: VerticalAlignment } = {};
  
  const lowerAlignment = alignment.toLowerCase();
  
  // 水平配置
  if (lowerAlignment.includes('left')) {
    result.horizontal = 'left';
  } else if (lowerAlignment.includes('center')) {
    result.horizontal = 'center';
  } else if (lowerAlignment.includes('right')) {
    result.horizontal = 'right';
  }
  
  // 垂直配置
  if (lowerAlignment.includes('top')) {
    result.vertical = 'top';
  } else if (lowerAlignment.includes('middle')) {
    result.vertical = 'middle';
  } else if (lowerAlignment.includes('bottom')) {
    result.vertical = 'bottom';
  }
  
  return result;
}
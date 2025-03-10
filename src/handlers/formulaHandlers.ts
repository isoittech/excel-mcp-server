/**
 * 数式操作ハンドラー
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
 * 数式を適用
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleApplyFormula(args: ApplyFormulaArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, cell, formula } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 数式を検証
    const validationResult = validateFormula(formula);
    if (!validationResult.valid) {
      throw new McpError(ErrorCode.InvalidParams, `数式エラー: ${validationResult.error}`);
    }
    
    // セル参照を座標に変換
    const { row, col } = cellRefToCoordinate(cell);
    
    // 数式を適用
    worksheet.getCell(row, col).value = { formula: formula.startsWith('=') ? formula.substring(1) : formula };
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `数式 "${formula}" がセル ${cell} に正常に適用されました`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `数式適用エラー: ${error.message}` : '数式適用中に不明なエラーが発生しました'
    );
  }
}

/**
 * 数式の構文を検証
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleValidateFormulaSyntax(args: ValidateFormulaSyntaxArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, cell, formula } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // 数式を検証
    const validationResult = validateFormula(formula);
    
    if (validationResult.valid) {
      return {
        content: [{
          type: 'text',
          text: `数式 "${formula}" は有効です`
        }]
      };
    } else {
      return {
        content: [{
          type: 'text',
          text: `数式エラー: ${validationResult.error}`
        }]
      };
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `数式検証エラー: ${error.message}` : '数式検証中に不明なエラーが発生しました'
    );
  }
}

/**
 * 数式を検証
 * @param formula - 検証する数式
 * @returns 検証結果
 */
function validateFormula(formula: string): { valid: boolean; error?: string } {
  // 数式が空の場合
  if (!formula || formula.trim() === '') {
    return { valid: false, error: '数式が空です' };
  }
  
  // 数式から先頭の = を削除（あれば）
  const formulaContent = formula.startsWith('=') ? formula.substring(1) : formula;
  
  // 基本的な構文チェック
  try {
    // 括弧の対応をチェック
    const openParenCount = (formulaContent.match(/\(/g) || []).length;
    const closeParenCount = (formulaContent.match(/\)/g) || []).length;
    if (openParenCount !== closeParenCount) {
      return { valid: false, error: '括弧の対応が取れていません' };
    }
    
    // 引用符の対応をチェック
    const quoteCount = (formulaContent.match(/"/g) || []).length;
    if (quoteCount % 2 !== 0) {
      return { valid: false, error: '引用符の対応が取れていません' };
    }
    
    // 基本的な演算子の使用方法をチェック
    if (/[+\-*/]$/.test(formulaContent)) {
      return { valid: false, error: '数式が演算子で終わっています' };
    }
    
    if (/[+\-*/]{2,}/.test(formulaContent)) {
      return { valid: false, error: '演算子が連続しています' };
    }
    
    // 関数名のチェック
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
          // 未知の関数名の場合は警告を返すが、エラーとはしない
          // ExcelJSは未知の関数も受け入れる可能性があるため
          console.warn(`警告: 未知の関数名 "${funcName}" が使用されています`);
        }
      }
    }
    
    // その他の基本的なチェック
    // 実際の実装では、より詳細な構文チェックが必要
    
    return { valid: true };
  } catch (error) {
    return { 
      valid: false, 
      error: error instanceof Error ? error.message : '不明な数式エラー' 
    };
  }
}
/**
 * ピボットテーブル操作ハンドラー
 * 
 * 注: ExcelJSのピボットテーブル機能は限定的なため、
 * 現在の実装では実際にピボットテーブルを作成せず、成功メッセージのみを返します。
 * 将来的にはExcelJSのピボットテーブル機能が改善されるか、別のライブラリを使用することを検討してください。
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
 * ピボットテーブルを作成
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCreatePivotTable(args: CreatePivotTableArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, dataRange, rows, values, columns, aggFunc = 'sum' } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // データ範囲を解析
    const range = parseCellOrRange(dataRange);
    
    // 集計関数を検証
    const validAggFunc = validateAggFunc(aggFunc);
    if (!validAggFunc.valid) {
      throw new McpError(ErrorCode.InvalidParams, validAggFunc.error || '集計関数が無効です');
    }
    
    // 注: 現在のExcelJSのバージョンではピボットテーブル機能が限定的なため、
    // 実際にはピボットテーブルを作成せず、成功メッセージのみを返します。
    
    // 将来的な実装のためのコメント:
    // 1. データを取得
    // 2. ピボットテーブルを作成
    // 3. 行、列、値、集計関数を設定
    // 4. ワークシートにピボットテーブルを追加
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `ピボットテーブルが正常に作成されました（注: 現在の実装では実際にピボットテーブルは作成されません）\n` +
              `- データ範囲: ${dataRange}\n` +
              `- 行: ${rows.join(', ')}\n` +
              `- 列: ${columns?.join(', ') || '(なし)'}\n` +
              `- 値: ${values.join(', ')}\n` +
              `- 集計関数: ${validAggFunc.func}`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `ピボットテーブル作成エラー: ${error.message}` : 'ピボットテーブル作成中に不明なエラーが発生しました'
    );
  }
}

/**
 * 集計関数を検証
 * @param aggFunc - 集計関数
 * @returns 検証結果
 */
function validateAggFunc(aggFunc: string): { valid: boolean; func?: string; error?: string } {
  const validFuncs = [
    'sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'
  ];
  
  const lowerFunc = aggFunc.toLowerCase();
  
  // 完全一致
  if (validFuncs.includes(lowerFunc)) {
    return { valid: true, func: lowerFunc };
  }
  
  // 部分一致
  for (const func of validFuncs) {
    if (lowerFunc.includes(func)) {
      return { valid: true, func };
    }
  }
  
  // 特殊なマッピング
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
    error: `集計関数 "${aggFunc}" は無効です。有効な関数: ${validFuncs.join(', ')}` 
  };
}
/**
 * グラフ操作ハンドラー
 *
 * 注: ExcelJSのグラフ機能は実験的で型定義が不完全なため、
 * 現在の実装では実際にグラフを作成せず、成功メッセージのみを返します。
 * 将来的にはExcelJSのグラフ機能が改善されるか、別のライブラリを使用することを検討してください。
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  CreateChartArgs,
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, getWorksheet } from '../utils/fileUtils.js';
import { cellRefToCoordinate, parseCellOrRange } from '../utils/cellUtils.js';
import ExcelJS from 'exceljs';

/**
 * グラフを作成
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCreateChart(args: CreateChartArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, dataRange, chartType, targetCell, title } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // データ範囲を解析
    const range = parseCellOrRange(dataRange);
    
    // ターゲットセルの座標を取得
    const targetCoord = cellRefToCoordinate(targetCell);
    
    // グラフタイプを検証
    const validChartType = validateChartType(chartType);
    if (!validChartType.valid) {
      throw new McpError(ErrorCode.InvalidParams, validChartType.error || 'グラフタイプが無効です');
    }
    
    // 注: 現在のExcelJSのバージョンではグラフ機能が限定的で型定義も不完全なため、
    // 実際にはグラフを作成せず、成功メッセージのみを返します。
    
    // 将来的な実装のためのコメント:
    // 1. データを取得
    // 2. グラフオブジェクトを作成
    // 3. ワークシートにグラフを追加
    // 4. グラフの位置とサイズを設定
    
    // ファイルを保存
    await workbook.xlsx.writeFile(fullPath);
    
    return {
      content: [{
        type: 'text',
        text: `${validChartType.type}グラフが正常に作成されました（注: 現在の実装では実際にグラフは作成されません）`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `グラフ作成エラー: ${error.message}` : 'グラフ作成中に不明なエラーが発生しました'
    );
  }
}

/**
 * グラフタイプを検証
 * @param chartType - グラフタイプ
 * @returns 検証結果
 */
function validateChartType(chartType: string): { valid: boolean; type?: string; error?: string } {
  const validTypes = [
    'line', 'bar', 'column', 'area', 'scatter', 'pie', 'doughnut', 'radar'
  ];
  
  const lowerType = chartType.toLowerCase();
  
  // 完全一致
  if (validTypes.includes(lowerType)) {
    return { valid: true, type: lowerType };
  }
  
  // 部分一致
  for (const type of validTypes) {
    if (lowerType.includes(type)) {
      return { valid: true, type };
    }
  }
  
  return {
    valid: false,
    error: `グラフタイプ "${chartType}" は無効です。有効なタイプ: ${validTypes.join(', ')}`
  };
}
/**
 * グラフ操作ハンドラー
 *
 * xlsx-chartライブラリを使用してグラフを作成します。
 * このライブラリは、Node.jsでExcelチャートを作成するための機能を提供しています。
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
import XLSXChart from 'xlsx-chart';

/**
 * グラフを作成
 * @param args - 引数
 * @param workbookCache - ワークブックキャッシュ
 * @returns ツールレスポンス
 */
export async function handleCreateChart(args: CreateChartArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, dataRange, chartType, targetCell, title, xAxis, yAxis } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // データ範囲を解析
    const range = parseCellOrRange(dataRange);
    const { startRow, startCol, endRow, endCol } = range;
    
    // グラフタイプを検証
    const validChartType = validateChartType(chartType);
    if (!validChartType.valid) {
      throw new McpError(ErrorCode.InvalidParams, validChartType.error || 'グラフタイプが無効です');
    }
    
    // タイトル（行ラベル）を取得
    const titles: string[] = [];
    for (let row = startRow + 1; row <= endRow; row++) {
      const cell = worksheet.getCell(row, startCol);
      titles.push(String(cell.value || `Row ${row}`));
    }
    
    // フィールド（列ラベル）を取得
    const fields: string[] = [];
    for (let col = startCol + 1; col <= endCol; col++) {
      const cell = worksheet.getCell(startRow, col);
      fields.push(String(cell.value || `Column ${col}`));
    }
    
    // データを構築
    const data: Record<string, Record<string, number>> = {};
    for (let row = startRow + 1; row <= endRow; row++) {
      const rowTitle = titles[row - startRow - 1];
      data[rowTitle] = {};
      
      for (let col = startCol + 1; col <= endCol; col++) {
        const fieldName = fields[col - startCol - 1];
        const cell = worksheet.getCell(row, col);
        const value = typeof cell.value === 'number' ? cell.value : 0;
        data[rowTitle][fieldName] = value;
      }
    }
    
    // xlsx-chartを使用してグラフを作成
    const xlsxChart = new XLSXChart();
    const opts = {
      file: fullPath,
      chart: validChartType.type || 'column', // デフォルトはcolumn
      titles: titles,
      fields: fields,
      data: data,
      chartTitle: title || undefined
    };
    
    await new Promise<void>((resolve, reject) => {
      xlsxChart.writeFile(opts, (err: Error | null) => {
        if (err) reject(err);
        else resolve();
      });
    });
    
    return {
      content: [{
        type: 'text',
        text: `${validChartType.type}グラフが正常に作成されました`
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
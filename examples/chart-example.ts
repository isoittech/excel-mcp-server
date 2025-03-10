/**
 * xlsx-chartライブラリを使用したグラフ作成の実装例
 *
 * 注意: このファイルを実行するには、まず xlsx-chart ライブラリをインストールする必要があります。
 * npm install xlsx-chart --save
 */

import XLSXChart from 'xlsx-chart';
import { resolve, dirname, basename } from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

// ESモジュールで__dirnameを使用するための設定
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * サンプルデータを使用してグラフを作成する
 */
async function createSampleChart() {
  // サンプルデータ
  const data = {
    '2020年': {
      '1月': 100,
      '2月': 120,
      '3月': 140,
      '4月': 160
    },
    '2021年': {
      '1月': 110,
      '2月': 130,
      '3月': 150,
      '4月': 170
    },
    '2022年': {
      '1月': 120,
      '2月': 140,
      '3月': 160,
      '4月': 180
    }
  };

  // タイトルとフィールド
  const titles = Object.keys(data);
  const fields = Object.keys(data[titles[0]]);

  // 出力ファイルパス
  const outputPath = resolve(__dirname, '../excel_files/chart-example.xlsx');

  // xlsx-chartを使用してグラフを作成
  const xlsxChart = new XLSXChart();
  const opts = {
    file: outputPath,
    chart: 'column',
    titles: titles,
    fields: fields,
    data: data,
    chartTitle: '月別売上推移'
  };

  return new Promise<void>((resolve, reject) => {
    xlsxChart.writeFile(opts, (err) => {
      if (err) {
        console.error('グラフ作成エラー:', err);
        reject(err);
      } else {
        console.log(`グラフが正常に作成されました: ${outputPath}`);
        resolve();
      }
    });
  });
}

/**
 * 既存のExcelファイルからデータを読み込んでグラフを作成する
 * @param filePath Excelファイルパス
 * @param sheetName シート名
 * @param dataRange データ範囲（例: 'A1:E4'）
 * @param chartType グラフタイプ
 * @param title グラフタイトル
 */
async function createChartFromExcel(
  filePath: string,
  sheetName: string,
  dataRange: string,
  chartType: string,
  title?: string
): Promise<void> {
  try {
    // ExcelJSでファイルを読み込む
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    
    if (!worksheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }

    // データ範囲を解析
    const [startCell, endCell] = dataRange.split(':');
    const startMatch = startCell.match(/([A-Z]+)(\d+)/);
    const endMatch = endCell.match(/([A-Z]+)(\d+)/);
    
    if (!startMatch || !endMatch) {
      throw new Error(`無効なデータ範囲: ${dataRange}`);
    }
    
    const startCol = columnNameToNumber(startMatch[1]);
    const startRow = parseInt(startMatch[2]);
    const endCol = columnNameToNumber(endMatch[1]);
    const endRow = parseInt(endMatch[2]);

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

    // 出力ファイルパス
    const outputPath = resolve(dirname(filePath), `${basename(filePath, '.xlsx')}-chart.xlsx`);

    // xlsx-chartを使用してグラフを作成
    const xlsxChart = new XLSXChart();
    const opts = {
      file: outputPath,
      chart: chartType,
      titles: titles,
      fields: fields,
      data: data,
      chartTitle: title
    };

    return new Promise<void>((resolvePromise, reject) => {
      xlsxChart.writeFile(opts, (err) => {
        if (err) {
          console.error('グラフ作成エラー:', err);
          reject(err);
        } else {
          console.log(`グラフが正常に作成されました: ${outputPath}`);
          resolvePromise();
        }
      });
    });
  } catch (error) {
    console.error('エラー:', error);
    throw error;
  }
}

/**
 * 列名を数値に変換する（例: 'A' -> 1, 'Z' -> 26, 'AA' -> 27）
 * @param columnName 列名
 * @returns 列番号
 */
function columnNameToNumber(columnName: string): number {
  let result = 0;
  for (let i = 0; i < columnName.length; i++) {
    result = result * 26 + (columnName.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * メイン関数
 */
async function main() {
  try {
    // サンプルデータでグラフを作成
    await createSampleChart();
    
    // 既存のExcelファイルからグラフを作成
    // await createChartFromExcel(
    //   resolve(__dirname, '../excel_files/sample-data.xlsx'),
    //   'Sheet1',
    //   'A1:E4',
    //   'line',
    //   '売上推移'
    // );
    
    console.log('処理が完了しました');
  } catch (error) {
    console.error('エラー:', error);
  }
}

// ESモジュールでの直接実行の検出
if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}

export {
  createSampleChart,
  createChartFromExcel
};
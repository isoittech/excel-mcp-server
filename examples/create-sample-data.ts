/**
 * サンプルデータを含むExcelファイルを作成するスクリプト
 */

import ExcelJS from 'exceljs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

// ESモジュールで__dirnameを使用するための設定
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * サンプルデータを含むExcelファイルを作成する
 */
async function createSampleDataFile() {
  // 出力ファイルパス
  const outputPath = resolve(__dirname, '../../excel_files/sample-data.xlsx');
  
  // ワークブックを作成
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('売上データ');
  
  // ヘッダー行を追加
  worksheet.addRow(['', '1月', '2月', '3月', '4月']);
  
  // データ行を追加
  worksheet.addRow(['2020年', 100, 120, 140, 160]);
  worksheet.addRow(['2021年', 110, 130, 150, 170]);
  worksheet.addRow(['2022年', 120, 140, 160, 180]);
  
  // ヘッダー行のスタイルを設定
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };
  });
  
  // 1列目のスタイルを設定
  for (let i = 2; i <= 4; i++) {
    const cell = worksheet.getCell(`A${i}`);
    cell.font = { bold: true };
  }
  
  // 列幅を調整
  worksheet.getColumn('A').width = 15;
  worksheet.getColumn('B').width = 10;
  worksheet.getColumn('C').width = 10;
  worksheet.getColumn('D').width = 10;
  worksheet.getColumn('E').width = 10;
  
  // ファイルを保存
  await workbook.xlsx.writeFile(outputPath);
  console.log(`サンプルデータファイルが作成されました: ${outputPath}`);
}

/**
 * メイン関数
 */
async function main() {
  try {
    await createSampleDataFile();
    console.log('処理が完了しました');
  } catch (error) {
    console.error('エラー:', error);
  }
}

// スクリプトが直接実行された場合のみ実行
if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}

export {
  createSampleDataFile
};
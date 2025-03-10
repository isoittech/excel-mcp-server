# グラフ機能の実装提案

## 現状

現在のExcel MCPサーバーのNode.js/TypeScript版では、グラフ機能が仮実装されています。これはExcelJSのグラフ機能が実験的で型定義が不完全なためです。現在の実装では、実際にグラフを作成せず、成功メッセージのみを返しています。

## 提案: xlsx-chartライブラリの使用

[objectum/xlsx-chart](https://github.com/objectum/xlsx-chart)ライブラリを使用することで、グラフ機能を実装できる可能性があります。このライブラリは、Node.jsでExcelチャートを作成するための機能を提供しています。

### インストール方法

```bash
npm install xlsx-chart --save
```

### 実装例

以下は、xlsx-chartライブラリを使用したグラフ作成の実装例です。

```typescript
import XLSXChart from 'xlsx-chart';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { CreateChartArgs, ToolResponse } from '../types/index.js';
import { getExcelPath } from '../utils/fileUtils.js';
import { parseCellOrRange } from '../utils/cellUtils.js';
import ExcelJS from 'exceljs';

export async function handleCreateChart(args: CreateChartArgs): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, dataRange, chartType, targetCell, title, xAxis, yAxis } = args;
    const fullPath = getExcelPath(filePath);
    
    // データを準備するためにExcelJSでファイルを読み込む
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fullPath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `シート '${sheetName}' が見つかりません`);
    }
    
    // データ範囲を解析
    const range = parseCellOrRange(dataRange);
    const { startRow, startCol, endRow, endCol } = range;
    
    // タイトル（行ラベル）を取得
    const titles: string[] = [];
    for (let row = startRow + 1; row <= endRow; row++) {
      const cell = worksheet.getCell(row, startCol);
      titles.push(cell.text || `Row ${row}`);
    }
    
    // フィールド（列ラベル）を取得
    const fields: string[] = [];
    for (let col = startCol + 1; col <= endCol; col++) {
      const cell = worksheet.getCell(startRow, col);
      fields.push(cell.text || `Column ${col}`);
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
    
    // グラフタイプを検証
    const validChartTypes = ['line', 'bar', 'column', 'area', 'radar', 'scatter', 'pie'];
    const lowerType = chartType.toLowerCase();
    let chartTypeToUse = lowerType;
    
    if (!validChartTypes.includes(lowerType)) {
      // 部分一致を試みる
      const matchedType = validChartTypes.find(type => lowerType.includes(type));
      if (!matchedType) {
        throw new McpError(
          ErrorCode.InvalidParams,
          `グラフタイプ "${chartType}" は無効です。有効なタイプ: ${validChartTypes.join(', ')}`
        );
      }
      chartTypeToUse = matchedType;
    }
    
    // xlsx-chartを使用してグラフを作成
    const xlsxChart = new XLSXChart();
    const opts = {
      file: fullPath,
      chart: chartTypeToUse,
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
        text: `${chartTypeToUse}グラフが正常に作成されました`
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
```

### 注意点

1. xlsx-chartライブラリは、ExcelJSとは異なるアプローチでExcelチャートを作成します。そのため、既存のExcelJSのワークブックキャッシュ機能との統合には注意が必要です。

2. xlsx-chartライブラリは、特定のデータ構造を要求します。そのため、ExcelJSで読み込んだデータを適切な形式に変換する必要があります。

3. xlsx-chartライブラリは、グラフの位置を指定するための直接的な方法を提供していない可能性があります。そのため、targetCellパラメータの使用方法については、さらなる調査が必要です。

4. xlsx-chartライブラリは、TypeScriptの型定義を提供していない可能性があります。そのため、型定義ファイルを作成する必要があるかもしれません。

## 実装手順

1. xlsx-chartライブラリをインストールする
   ```bash
   npm install xlsx-chart --save
   ```

2. 必要に応じて、xlsx-chartライブラリの型定義ファイルを作成する
   ```typescript
   // src/types/xlsx-chart.d.ts
   declare module 'xlsx-chart' {
     interface XLSXChartOptions {
       file?: string;
       chart: string;
       titles: string[];
       fields: string[];
       data: Record<string, Record<string, number>>;
       chartTitle?: string;
       templatePath?: string;
     }
     
     class XLSXChart {
       constructor();
       writeFile(options: XLSXChartOptions, callback: (err: Error | null) => void): void;
       generate(options: XLSXChartOptions, callback: (err: Error | null, data: Buffer) => void): void;
     }
     
     export = XLSXChart;
   }
   ```

3. chartHandlers.tsファイルを修正して、xlsx-chartライブラリを使用してグラフを作成する

4. テストを実施して、グラフ機能が正常に動作することを確認する

## 代替案

1. ExcelJSのグラフ機能が将来的に改善される可能性があるため、ExcelJSの最新バージョンを定期的に確認する

2. 他のExcelグラフ作成ライブラリを調査する
   - [officegen](https://github.com/Ziv-Barber/officegen)
   - [excel4node](https://github.com/natergj/excel4node)
   - [node-xlsx](https://github.com/mgcrea/node-xlsx)

3. Python版の実装を参考にして、独自のグラフ作成機能を実装する
# EXCEL 操作 MCP サーバー

このサーバーはローカルの EXCEL ファイルを操作するための MCP ツールを提供します。

## 目次

1. [システム要件](#システム要件)
2. [インストール](#インストール)
3. [使用方法](#使用方法)
4. [注意事項](#注意事項)

## システム要件

- Node.js: 18.x 以上
- npm: 9.x 以上
- Excel ファイル形式: .xlsx
- 対応 OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## インストール

### 1. リポジトリのクローン

```bash
git clone git@github.com:virtuarian/excel-server.git
cd excel-server
```

### 2. 依存関係のインストールとビルド

```bash
npm install
npm run build
```

### 3. MCP サーバーの設定

▼ Cline から使用する場合

1. MCP サーバーを開く
2. Edit MCP Settings をクリック
3. 以下の excel エントリを追加:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["/cloned-path/build/index.js"],  // Github をクローンしたフォルダ
      "env": {},
      "disabled": false,
      "alwaysAllow": []
    }
  }
}
```

※ または、json ファイルを手動で開いて編集することもできます

## 使用方法

### EXCEL ファイルの読み込み

```json
{
  "server_name": "excel",
  "tool_name": "read_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "range": "A1:C10"
  }
}
```

### EXCEL ファイルへの書き込み

```json
{
  "server_name": "excel",
  "tool_name": "write_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "data": [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"]
    ]
  }
}
```

### 新規シートの作成

```json
{
  "server_name": "excel",
  "tool_name": "create_sheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "NewSheet"
  }
}
```

### 新規 EXCEL ファイルの作成

```json
{
  "server_name": "excel",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // オプション、デフォルトは "Sheet1"
  }
}
```

### ワークブックのメタデータ取得

```json
{
  "server_name": "excel",
  "tool_name": "get_workbook_metadata",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "includeRanges": false  // オプション、使用範囲情報を含めるかどうか
  }
}
```

### ワークシート名の変更

```json
{
  "server_name": "excel",
  "tool_name": "rename_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "oldName": "Sheet1",
    "newName": "NewName"
  }
}
```

### ワークシートの削除

```json
{
  "server_name": "excel",
  "tool_name": "delete_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1"
  }
}
```

### ワークシートのコピー

```json
{
  "server_name": "excel",
  "tool_name": "copy_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sourceSheet": "Sheet1",
    "targetSheet": "Sheet1Copy"
  }
}
```

### 数式の適用

```json
{
  "server_name": "excel",
  "tool_name": "apply_formula",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### 数式構文の検証

```json
{
  "server_name": "excel",
  "tool_name": "validate_formula_syntax",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### セル範囲の書式設定

```json
{
  "server_name": "excel",
  "tool_name": "format_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "bold": true,
    "italic": false,
    "fontSize": 12,
    "fontColor": "#FF0000",
    "bgColor": "#FFFF00"
  }
}
```

### セルの結合

```json
{
  "server_name": "excel",
  "tool_name": "merge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### セルの結合解除

```json
{
  "server_name": "excel",
  "tool_name": "unmerge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### セル範囲のコピー

```json
{
  "server_name": "excel",
  "tool_name": "copy_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "sourceStart": "A1",
    "sourceEnd": "C3",
    "targetStart": "D1",
    "targetSheet": "Sheet2"  // オプション、省略するとソースシートと同じ
  }
}
```

### セル範囲の削除

```json
{
  "server_name": "excel",
  "tool_name": "delete_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "shiftDirection": "up"  // "up" または "left"、デフォルトは "up"
  }
}
```

### Excel 範囲の検証

```json
{
  "server_name": "excel",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"  // オプション
  }
}
```

### グラフの作成

```json
{
  "server_name": "excel",
  "tool_name": "create_chart",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:C10",
    "chartType": "column",  // "column", "line", "bar", "area", "scatter", "pie"
    "targetCell": "E1",
    "title": "サンプルグラフ",  // オプション
    "xAxis": "X軸ラベル",  // オプション
    "yAxis": "Y軸ラベル"   // オプション
  }
}
```

### ピボットテーブルの作成

```json
{
  "server_name": "excel",
  "tool_name": "create_pivot_table",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:D100",
    "rows": ["Category"],
    "values": ["Sales"],
    "columns": ["Region"],  // オプション
    "aggFunc": "sum"  // "sum", "count", "average", "max", "min" など
  }
}
```

## 注意事項

- ファイルパスは絶対パスで指定する必要があります
- シート名が指定されていない場合、最初のシートが対象となります
- 範囲指定は "A1:C10" のような形式で行います
- create_excel で既存のファイルパスを指定するとエラーになります
- ピボットテーブル機能は現在の ExcelJS の制限により、実際にはピボットテーブルを作成せず、成功メッセージのみを返します

## 作者
- 作者: isoittech
- fork元: virtuarian

## ライセンス
この MCP サーバーは MIT ライセンスの下で提供されています。これは、MIT ライセンスの条件に従って、ソフトウェアを自由に使用、変更、配布できることを意味します。詳細については、プロジェクトリポジトリの LICENSE ファイルを参照してください。

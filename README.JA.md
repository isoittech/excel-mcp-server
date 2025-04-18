# EXCEL操作用 MCP サーバー

このサーバーはローカルの EXCEL ファイルを操作するための MCP ツールを提供します。

## 目次

1. [システム要件](#システム要件)
2. [インストール](#インストール)
3. [使い方](#使い方)
4. [注意事項](#注意事項)

## システム要件

- Node.js: 18.x 以上
- npm: 9.x 以上
- Excelファイル形式: .xlsx
- 対応OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## インストール

### 1. リポジトリのクローン

```bash
git clone git@github.com:isoittech/excel-mcp-server.git
cd excel-mcp-server
```

### 2. MCPサーバーの設定

#### Cline/Roo Code から利用する場合

1. MCP Servers を開く
2. 「Edit MCP Settings」をクリック
3. 下記の excel エントリを追加

#### Docker を使う場合

```json
{
  "mcpServers": {
    "excel": {
      "command": "docker",
      "args": [
        "run",
        "-i",
        "--rm",
        "-v",
        "/your/host/excel_files:/app/excel_files",
        "-e",
        "EXCEL_FILES_PATH=/app/excel_files",
        "isoittech/excel-mcp-server"
      ],
      "env": {},
      "disabled": false,
      "alwaysAllow": []
    }
  }
}
```

## 使い方

### EXCELファイルの読み込み

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

### EXCELファイルへの書き込み

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

### 新しいシートの作成

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

### 新しいEXCELファイルの作成

```json
{
  "server_name": "excel",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // 省略可、デフォルトは "Sheet1"
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
    "includeRanges": false  // 省略可、範囲情報を含めるかどうか
  }
}
```

### シート名の変更

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

### シートの削除

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

### シートのコピー

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

### セルへの数式適用

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
    "targetSheet": "Sheet2"  // 省略可、省略時は同じシート
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

### Excel範囲の検証

```json
{
  "server_name": "excel",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"  // 省略可
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
    "title": "サンプルグラフ",  // 省略可
    "xAxis": "X軸ラベル",      // 省略可
    "yAxis": "Y軸ラベル"       // 省略可
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
    "columns": ["Region"],  // 省略可
    "aggFunc": "sum"  // "sum", "count", "average", "max", "min" など
  }
}
```

## 注意事項

- ファイルパスは絶対パスで指定すること
- シート名を指定しない場合、最初のシートが対象になる
- 範囲指定は "A1:C10" のような形式で記述する
- create_excel で既存ファイルパスを指定するとエラーになる
- 現在の ExcelJS の制限により、ピボットテーブル機能は実際にはピボットテーブルを作成せず、成功メッセージのみ返す

## 作者

- 作者: isoittech
- フォーク元: virtuarian

## ライセンス

この MCP サーバーは MIT ライセンスで提供されています。  
自由に利用・改変・再配布できます。詳細はリポジトリ内の LICENSE ファイルを参照してください。

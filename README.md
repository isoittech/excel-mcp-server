# EXCEL操作MCPサーバー

このサーバーは、ローカルのEXCELファイルを操作するためのMCPツールを提供します。

## 目次

1. [インストール手順](#インストール手順)
2. [使用方法](#使用方法)
3. [注意事項](#注意事項)

## インストール手順

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

### 3. MCPサーバの設定

▼ Clineから利用する場合

1. MCP Servers を開く
2. Edit MCP Settings をクリック
3. 以下のexcelのエントリーを追加する

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["/cloned-path/build/index.js"], // GithubをCloneしたフォルダ
      "env": {},
      "disabled": false,
      "alwaysAllow": []
    }
  }
}
```

※ あるいは手動でjsonファイルを開いて修正してください

## 使用方法

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
    "sheetName": "Sheet1"  // オプション、デフォルトは"Sheet1"
  }
}
```

## 注意事項

- ファイルパスは絶対パスで指定してください
- シート名を指定しない場合、最初のシートが操作対象になります
- 範囲指定は"A1:C10"のような形式で指定できます
- create_excelで既存のファイルパスを指定するとエラーになります



## 使用方法

### 1. EXCELファイルの読み込み

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

### 2. EXCELファイルへの書き込み

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

### 3. 新しいシートの作成

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

### 4. 新しいEXCELファイルの作成

```json
{
  "server_name": "excel",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // オプション、デフォルトは"Sheet1"
  }
}
```

## 注意事項
- ファイルパスは絶対パスで指定してください
- シート名を指定しない場合、最初のシートが操作対象になります
- 範囲指定は"A1:C10"のような形式で指定できます
- create_excelで既存のファイルパスを指定するとエラーになります

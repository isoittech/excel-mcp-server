# EXCEL操作MCPサーバー

このサーバーは、ローカルのEXCELファイルを操作するためのMCPツールを提供します。

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

## 注意事項
- ファイルパスは絶対パスで指定してください
- シート名を指定しない場合、最初のシートが操作対象になります
- 範囲指定は"A1:C10"のような形式で指定できます

# EXCEL Operation MCP Server

This server provides MCP tools for manipulating local EXCEL files.

## Table of Contents

1. [System Requirements](#system-requirements)
2. [Installation](#installation)
3. [Usage](#usage)
4. [Notes](#notes)

## System Requirements

- Node.js: 18.x or higher
- npm: 9.x or higher
- Excel file format: .xlsx
- Supported OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## Installation

### 1. Clone the Repository

```bash
git clone git@github.com:virtuarian/excel-server.git
cd excel-server
```

### 2. Install Dependencies and Build

```bash
npm install
npm run build
```

### 3. Configure MCP Server

▼ When using from Cline

1. Open MCP Servers
2. Click Edit MCP Settings
3. Add the following excel entry:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["/cloned-path/build/index.js"],  // Folder where Github is cloned
      "env": {},
      "disabled": false,
      "alwaysAllow": []
    }
  }
}
```

※ Alternatively, manually open and edit the json file

## Usage

### Read EXCEL File

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

### Write to EXCEL File

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

### Create New Sheet

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

### Create New EXCEL File

```json
{
  "server_name": "excel",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // Optional, default is "Sheet1"
  }
}
```

### Get Workbook Metadata

```json
{
  "server_name": "excel",
  "tool_name": "get_workbook_metadata",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "includeRanges": false  // Optional, whether to include range information
  }
}
```

### Rename Worksheet

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

### Delete Worksheet

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

### Copy Worksheet

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

### Apply Formula

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

### Validate Formula Syntax

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

### Format Cell Range

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

### Merge Cells

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

### Unmerge Cells

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

### Copy Cell Range

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
    "targetSheet": "Sheet2"  // Optional, defaults to source sheet if omitted
  }
}
```

### Delete Cell Range

```json
{
  "server_name": "excel",
  "tool_name": "delete_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "shiftDirection": "up"  // "up" or "left", defaults to "up"
  }
}
```

### Validate Excel Range

```json
{
  "server_name": "excel",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"  // Optional
  }
}
```

### Create Chart

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
    "title": "Sample Chart",  // Optional
    "xAxis": "X Axis Label",  // Optional
    "yAxis": "Y Axis Label"   // Optional
  }
}
```

### Create Pivot Table

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
    "columns": ["Region"],  // Optional
    "aggFunc": "sum"  // "sum", "count", "average", "max", "min", etc.
  }
}
```

## Notes

- File paths must be specified as absolute paths
- If no sheet name is specified, the first sheet will be the target
- Range specification should be in formats like "A1:C10"
- Specifying an existing file path in create_excel will result in an error
- Due to limitations in the current version of ExcelJS, the pivot table functionality does not actually create pivot tables but only returns success messages

## Author
- Author: isoittech
- Forked from: virtuarian

## License
This MCP server is licensed under the MIT License. This means you are free to use, modify, and distribute the software, subject to the terms and conditions of the MIT License. For more details, please see the LICENSE file in the project repository.
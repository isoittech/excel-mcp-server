# EXCEL Operation MCP Server

This server provides MCP tools for manipulating local EXCEL files.

## Table of Contents

1. [Installation](#installation)
2. [Usage](#usage)
3. [Notes](#notes)

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

## Notes

- File paths must be specified as absolute paths
- If no sheet name is specified, the first sheet will be the target
- Range specification should be in formats like "A1:C10"
- Specifying an existing file path in create_excel will result in an error
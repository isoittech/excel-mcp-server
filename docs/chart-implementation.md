# Chart Implementation Proposal

## Current Status

In the current Excel MCP server's Node.js/TypeScript version, chart functionality is provisionally implemented. This is because ExcelJS's chart functionality is experimental and has incomplete type definitions. The current implementation does not actually create charts but only returns a success message.

## Proposal: Using xlsx-chart Library

The [objectum/xlsx-chart](https://github.com/objectum/xlsx-chart) library could be used to implement chart functionality. This library provides functionality for creating Excel charts in Node.js.

### Installation

```bash
npm install xlsx-chart --save
```

### Implementation Example

Below is an implementation example of chart creation using the xlsx-chart library.

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
    
    // Load file with ExcelJS to prepare data
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fullPath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet '${sheetName}' not found`);
    }
    
    // Parse data range
    const range = parseCellOrRange(dataRange);
    const { startRow, startCol, endRow, endCol } = range;
    
    // Get titles (row labels)
    const titles: string[] = [];
    for (let row = startRow + 1; row <= endRow; row++) {
      const cell = worksheet.getCell(row, startCol);
      titles.push(cell.text || `Row ${row}`);
    }
    
    // Get fields (column labels)
    const fields: string[] = [];
    for (let col = startCol + 1; col <= endCol; col++) {
      const cell = worksheet.getCell(startRow, col);
      fields.push(cell.text || `Column ${col}`);
    }
    
    // Build data
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
    
    // Validate chart type
    const validChartTypes = ['line', 'bar', 'column', 'area', 'radar', 'scatter', 'pie'];
    const lowerType = chartType.toLowerCase();
    let chartTypeToUse = lowerType;
    
    if (!validChartTypes.includes(lowerType)) {
      // Try partial match
      const matchedType = validChartTypes.find(type => lowerType.includes(type));
      if (!matchedType) {
        throw new McpError(
          ErrorCode.InvalidParams,
          `Chart type "${chartType}" is invalid. Valid types: ${validChartTypes.join(', ')}`
        );
      }
      chartTypeToUse = matchedType;
    }
    
    // Create chart using xlsx-chart
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
        text: `${chartTypeToUse} chart created successfully`
      }]
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      error instanceof Error ? `Chart creation error: ${error.message}` : 'Unknown error occurred while creating chart'
    );
  }
}
```

### Notes

1. The xlsx-chart library creates Excel charts using a different approach than ExcelJS. Care must be taken when integrating with the existing ExcelJS workbook cache functionality.

2. The xlsx-chart library requires a specific data structure. Therefore, data loaded with ExcelJS needs to be converted to the appropriate format.

3. The xlsx-chart library may not provide a direct method to specify chart position. Therefore, further investigation is needed on how to use the targetCell parameter.

4. The xlsx-chart library may not provide TypeScript type definitions. Therefore, it may be necessary to create type definition files.

## Implementation Steps

1. Install the xlsx-chart library
   ```bash
   npm install xlsx-chart --save
   ```

2. Create type definition file for xlsx-chart library if needed
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

3. Modify chartHandlers.ts file to create charts using the xlsx-chart library

4. Test to ensure chart functionality works correctly

## Alternatives

1. Regularly check for the latest version of ExcelJS as its chart functionality may improve in the future

2. Investigate other Excel chart creation libraries
   - [officegen](https://github.com/Ziv-Barber/officegen)
   - [excel4node](https://github.com/natergj/excel4node)
   - [node-xlsx](https://github.com/mgcrea/node-xlsx)

3. Implement custom chart creation functionality based on the Python version implementation
/**
 * Chart operation handlers
 *
 * Uses xlsx-chart library to create charts.
 * This library provides functionality for creating Excel charts in Node.js.
 */

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  CreateChartArgs,
  ToolResponse,
  WorkbookCache
} from '../types/index.js';
import { loadWorkbook, getExcelPath, getWorksheet } from '../utils/fileUtils.js';
import { cellRefToCoordinate, parseCellOrRange } from '../utils/cellUtils.js';
import ExcelJS from 'exceljs';
import XLSXChart from 'xlsx-chart';

/**
 * Create chart
 * @param args - Arguments
 * @param workbookCache - Workbook cache
 * @returns Tool response
 */
export async function handleCreateChart(args: CreateChartArgs, workbookCache: WorkbookCache): Promise<ToolResponse> {
  try {
    const { filePath, sheetName, dataRange, chartType, targetCell, title, xAxis, yAxis } = args;
    const fullPath = getExcelPath(filePath);
    const workbook = await loadWorkbook(fullPath, workbookCache);
    const worksheet = getWorksheet(workbook, sheetName);
    
    // Parse data range
    const range = parseCellOrRange(dataRange);
    const { startRow, startCol, endRow, endCol } = range;
    
    // Validate chart type
    const validChartType = validateChartType(chartType);
    if (!validChartType.valid) {
      throw new McpError(ErrorCode.InvalidParams, validChartType.error || 'Invalid chart type');
    }
    
    // Get titles (row labels)
    const titles: string[] = [];
    for (let row = startRow + 1; row <= endRow; row++) {
      const cell = worksheet.getCell(row, startCol);
      titles.push(String(cell.value || `Row ${row}`));
    }
    
    // Get fields (column labels)
    const fields: string[] = [];
    for (let col = startCol + 1; col <= endCol; col++) {
      const cell = worksheet.getCell(startRow, col);
      fields.push(String(cell.value || `Column ${col}`));
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
    
    // Create chart using xlsx-chart
    const xlsxChart = new XLSXChart();
    const opts = {
      file: fullPath,
      chart: validChartType.type || 'column', // Default is column
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
        text: `${validChartType.type} chart created successfully`
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

/**
 * Validate chart type
 * @param chartType - Chart type
 * @returns Validation result
 */
function validateChartType(chartType: string): { valid: boolean; type?: string; error?: string } {
  const validTypes = [
    'line', 'bar', 'column', 'area', 'scatter', 'pie', 'doughnut', 'radar'
  ];
  
  const lowerType = chartType.toLowerCase();
  
  // Exact match
  if (validTypes.includes(lowerType)) {
    return { valid: true, type: lowerType };
  }
  
  // Partial match
  for (const type of validTypes) {
    if (lowerType.includes(type)) {
      return { valid: true, type };
    }
  }
  
  return {
    valid: false,
    error: `Chart type "${chartType}" is invalid. Valid types: ${validTypes.join(', ')}`
  };
}
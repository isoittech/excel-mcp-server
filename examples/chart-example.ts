/**
 * Implementation example of chart creation using xlsx-chart library
 *
 * Note: To run this file, you need to install the xlsx-chart library first.
 * npm install xlsx-chart --save
 */

import XLSXChart from 'xlsx-chart';
import { resolve, dirname, basename } from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

// Configuration to use __dirname in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * Create chart using sample data
 */
async function createSampleChart() {
  // Sample data
  const data = {
    '2020': {
      'Jan': 100,
      'Feb': 120,
      'Mar': 140,
      'Apr': 160
    },
    '2021': {
      'Jan': 110,
      'Feb': 130,
      'Mar': 150,
      'Apr': 170
    },
    '2022': {
      'Jan': 120,
      'Feb': 140,
      'Mar': 160,
      'Apr': 180
    }
  };

  // Titles and fields
  const titles = Object.keys(data);
  const fields = Object.keys(data[titles[0]]);

  // Output file path
  const outputPath = resolve(__dirname, '../excel_files/chart-example.xlsx');

  // Create chart using xlsx-chart
  const xlsxChart = new XLSXChart();
  const opts = {
    file: outputPath,
    chart: 'column',
    titles: titles,
    fields: fields,
    data: data,
    chartTitle: '月別売上推移'
  };

  return new Promise<void>((resolve, reject) => {
    xlsxChart.writeFile(opts, (err) => {
      if (err) {
        console.error('Chart creation error:', err);
        reject(err);
      } else {
        console.log(`Chart created successfully: ${outputPath}`);
        resolve();
      }
    });
  });
}

/**
 * Create chart from data loaded from existing Excel file
 * @param filePath Excel file path
 * @param sheetName Sheet name
 * @param dataRange Data range (e.g., 'A1:E4')
 * @param chartType Chart type
 * @param title Chart title
 */
async function createChartFromExcel(
  filePath: string,
  sheetName: string,
  dataRange: string,
  chartType: string,
  title?: string
): Promise<void> {
  try {
    // Load file with ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);
    
    if (!worksheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }

    // Parse data range
    const [startCell, endCell] = dataRange.split(':');
    const startMatch = startCell.match(/([A-Z]+)(\d+)/);
    const endMatch = endCell.match(/([A-Z]+)(\d+)/);
    
    if (!startMatch || !endMatch) {
      throw new Error(`Invalid data range: ${dataRange}`);
    }
    
    const startCol = columnNameToNumber(startMatch[1]);
    const startRow = parseInt(startMatch[2]);
    const endCol = columnNameToNumber(endMatch[1]);
    const endRow = parseInt(endMatch[2]);

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

    // Output file path
    const outputPath = resolve(dirname(filePath), `${basename(filePath, '.xlsx')}-chart.xlsx`);

    // Create chart using xlsx-chart
    const xlsxChart = new XLSXChart();
    const opts = {
      file: outputPath,
      chart: chartType,
      titles: titles,
      fields: fields,
      data: data,
      chartTitle: title
    };

    return new Promise<void>((resolvePromise, reject) => {
      xlsxChart.writeFile(opts, (err) => {
        if (err) {
          console.error('Chart creation error:', err);
          reject(err);
        } else {
          console.log(`Chart created successfully: ${outputPath}`);
          resolvePromise();
        }
      });
    });
  } catch (error) {
    console.error('Error:', error);
    throw error;
  }
}

/**
 * Convert column name to number (e.g., 'A' -> 1, 'Z' -> 26, 'AA' -> 27)
 * @param columnName Column name
 * @returns Column number
 */
function columnNameToNumber(columnName: string): number {
  let result = 0;
  for (let i = 0; i < columnName.length; i++) {
    result = result * 26 + (columnName.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Main function
 */
async function main() {
  try {
    // Create chart with sample data
    await createSampleChart();
    
    // Create chart from existing Excel file
    // await createChartFromExcel(
    //   resolve(__dirname, '../excel_files/sample-data.xlsx'),
    //   'Sheet1',
    //   'A1:E4',
    //   'line',
    //   'Sales Trend'
    // );
    
    console.log('Processing completed');
  } catch (error) {
    console.error('Error:', error);
  }
}

// Detect direct execution in ES modules
if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}

export {
  createSampleChart,
  createChartFromExcel
};
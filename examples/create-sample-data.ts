/**
 * Script to create Excel file with sample data
 */

import ExcelJS from 'exceljs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

// Configuration to use __dirname in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * Create Excel file with sample data
 */
async function createSampleDataFile() {
  // Output file path
  const outputPath = resolve(__dirname, '../../excel_files/sample-data.xlsx');
  
  // Create workbook
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sales Data');
  
  // Add header row
  worksheet.addRow(['', 'Jan', 'Feb', 'Mar', 'Apr']);
  
  // Add data rows
  worksheet.addRow(['2020', 100, 120, 140, 160]);
  worksheet.addRow(['2021', 110, 130, 150, 170]);
  worksheet.addRow(['2022', 120, 140, 160, 180]);
  
  // Set header row style
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }
    };
  });
  
  // Set first column style
  for (let i = 2; i <= 4; i++) {
    const cell = worksheet.getCell(`A${i}`);
    cell.font = { bold: true };
  }
  
  // Adjust column widths
  worksheet.getColumn('A').width = 15;
  worksheet.getColumn('B').width = 10;
  worksheet.getColumn('C').width = 10;
  worksheet.getColumn('D').width = 10;
  worksheet.getColumn('E').width = 10;
  
  // Save file
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Sample data file created: ${outputPath}`);
}

/**
 * Main function
 */
async function main() {
  try {
    await createSampleDataFile();
    console.log('Processing completed');
  } catch (error) {
    console.error('Error:', error);
  }
}

// Execute only when script is run directly
if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}

export {
  createSampleDataFile
};
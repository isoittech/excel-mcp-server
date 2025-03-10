#!/usr/bin/env node
/**
 * Excel MCP Server
 * Main entry point
 */

import { ExcelServer } from './excelServer.js';
import fs from 'fs';
import path from 'path';

// Get Excel files path from environment variable
const EXCEL_FILES_PATH = process.env.EXCEL_FILES_PATH || './excel_files';

// Create directory if it doesn't exist
if (!fs.existsSync(EXCEL_FILES_PATH)) {
  fs.mkdirSync(EXCEL_FILES_PATH, { recursive: true });
}

// Start the server
console.error(`Starting Excel MCP server (files directory: ${EXCEL_FILES_PATH})`);
const server = new ExcelServer();
server.run().catch((error) => {
  console.error('Server failed:', error);
  process.exit(1);
});

// Signal handling
process.on('SIGINT', () => {
  console.error('Server stopped by user');
  process.exit(0);
});
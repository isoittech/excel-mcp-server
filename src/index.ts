#!/usr/bin/env node
/**
 * Excel MCP Server
 * メインエントリーポイント
 */

import { ExcelServer } from './excelServer.js';
import fs from 'fs';
import path from 'path';

// 環境変数からExcelファイルのパスを取得
const EXCEL_FILES_PATH = process.env.EXCEL_FILES_PATH || './excel_files';

// ディレクトリが存在しない場合は作成
if (!fs.existsSync(EXCEL_FILES_PATH)) {
  fs.mkdirSync(EXCEL_FILES_PATH, { recursive: true });
}

// サーバーを起動
console.error(`Starting Excel MCP server (files directory: ${EXCEL_FILES_PATH})`);
const server = new ExcelServer();
server.run().catch((error) => {
  console.error('Server failed:', error);
  process.exit(1);
});

// シグナルハンドリング
process.on('SIGINT', () => {
  console.error('Server stopped by user');
  process.exit(0);
});
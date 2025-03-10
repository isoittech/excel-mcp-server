/**
 * セル操作ユーティリティ関数
 */

import { CellCoordinate, CellRange } from '../types/index.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * セル参照（例：A1）を行と列の数値に変換
 * @param cellRef - セル参照（例：A1）
 * @returns 行と列の数値
 */
export function cellRefToCoordinate(cellRef: string): CellCoordinate {
  const colMatch = cellRef.match(/[A-Za-z]+/);
  const rowMatch = cellRef.match(/\d+/);
  
  if (!colMatch || !rowMatch) {
    throw new McpError(ErrorCode.InvalidParams, `無効なセル参照形式: ${cellRef}`);
  }
  
  const col = columnLetterToNumber(colMatch[0]);
  const row = parseInt(rowMatch[0], 10);
  
  return { row, col };
}

/**
 * 列の文字（例：A, B, AA）を数値に変換
 * @param letters - 列の文字（例：A, B, AA）
 * @returns 数値（1から始まる）
 */
export function columnLetterToNumber(letters: string): number {
  let column = 0;
  letters = letters.toUpperCase();
  for (let i = 0; i < letters.length; i++) {
    column = column * 26 + (letters.charCodeAt(i) - 64);
  }
  return column;
}

/**
 * 数値を列の文字（例：A, B, AA）に変換
 * @param num - 数値（1から始まる）
 * @returns 列の文字
 */
export function numberToColumnLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    const modulo = (num - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    num = Math.floor((num - modulo) / 26);
  }
  return letter;
}

/**
 * 座標をセル参照（例：A1）に変換
 * @param row - 行番号（1から始まる）
 * @param col - 列番号（1から始まる）
 * @returns セル参照
 */
export function coordinateToCellRef(row: number, col: number): string {
  return `${numberToColumnLetter(col)}${row}`;
}

/**
 * 範囲文字列（例：A1:C10）を解析して座標に変換
 * @param range - 範囲文字列（例：A1:C10）
 * @returns 範囲の座標
 */
export function parseRange(range: string): CellRange {
  const rangeParts = range.split(':');
  if (rangeParts.length !== 2) {
    throw new McpError(ErrorCode.InvalidParams, '無効な範囲形式。"A1:C10"のような形式が必要です');
  }

  const [startCell, endCell] = rangeParts;
  const startCoord = cellRefToCoordinate(startCell);
  const endCoord = cellRefToCoordinate(endCell);

  return {
    startCol: startCoord.col,
    startRow: startCoord.row,
    endCol: endCoord.col,
    endRow: endCoord.row
  };
}

/**
 * 単一セルの場合も範囲として解析（例：A1 → A1:A1）
 * @param cellOrRange - セル参照または範囲文字列
 * @returns 範囲の座標
 */
export function parseCellOrRange(cellOrRange: string): CellRange {
  if (cellOrRange.includes(':')) {
    return parseRange(cellOrRange);
  } else {
    const coord = cellRefToCoordinate(cellOrRange);
    return {
      startCol: coord.col,
      startRow: coord.row,
      endCol: coord.col,
      endRow: coord.row
    };
  }
}

/**
 * 範囲の座標を範囲文字列に変換
 * @param range - 範囲の座標
 * @returns 範囲文字列（例：A1:C10）
 */
export function formatRange(range: CellRange): string {
  const startCell = coordinateToCellRef(range.startRow, range.startCol);
  const endCell = coordinateToCellRef(range.endRow, range.endCol);
  return `${startCell}:${endCell}`;
}

/**
 * 範囲が有効かどうかを検証
 * @param range - 範囲の座標
 * @returns 有効な場合はtrue、無効な場合はfalse
 */
export function isValidRange(range: CellRange): boolean {
  return (
    range.startRow > 0 &&
    range.startCol > 0 &&
    range.endRow >= range.startRow &&
    range.endCol >= range.startCol
  );
}
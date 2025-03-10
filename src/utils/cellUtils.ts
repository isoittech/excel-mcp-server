/**
 * Cell operation utility functions
 */

import { CellCoordinate, CellRange } from '../types/index.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

/**
 * Convert cell reference (e.g., A1) to row and column numbers
 * @param cellRef - Cell reference (e.g., A1)
 * @returns Row and column numbers
 */
export function cellRefToCoordinate(cellRef: string): CellCoordinate {
  const colMatch = cellRef.match(/[A-Za-z]+/);
  const rowMatch = cellRef.match(/\d+/);
  
  if (!colMatch || !rowMatch) {
    throw new McpError(ErrorCode.InvalidParams, `Invalid cell reference format: ${cellRef}`);
  }
  
  const col = columnLetterToNumber(colMatch[0]);
  const row = parseInt(rowMatch[0], 10);
  
  return { row, col };
}

/**
 * Convert column letters (e.g., A, B, AA) to number
 * @param letters - Column letters (e.g., A, B, AA)
 * @returns Number (starting from 1)
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
 * Convert number to column letters (e.g., A, B, AA)
 * @param num - Number (starting from 1)
 * @returns Column letters
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
 * Convert coordinates to cell reference (e.g., A1)
 * @param row - Row number (starting from 1)
 * @param col - Column number (starting from 1)
 * @returns Cell reference
 */
export function coordinateToCellRef(row: number, col: number): string {
  return `${numberToColumnLetter(col)}${row}`;
}

/**
 * Parse range string (e.g., A1:C10) and convert to coordinates
 * @param range - Range string (e.g., A1:C10)
 * @returns Range coordinates
 */
export function parseRange(range: string): CellRange {
  const rangeParts = range.split(':');
  if (rangeParts.length !== 2) {
    throw new McpError(ErrorCode.InvalidParams, 'Invalid range format. Format like "A1:C10" is required');
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
 * Parse cell or range (e.g., A1 → A1:A1)
 * @param cellOrRange - Cell reference or range string
 * @returns Range coordinates
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
 * Convert range coordinates to range string
 * @param range - Range coordinates
 * @returns Range string (e.g., A1:C10)
 */
export function formatRange(range: CellRange): string {
  const startCell = coordinateToCellRef(range.startRow, range.startCol);
  const endCell = coordinateToCellRef(range.endRow, range.endCol);
  return `${startCell}:${endCell}`;
}

/**
 * Validate if range is valid
 * @param range - Range coordinates
 * @returns true if valid, false if invalid
 */
export function isValidRange(range: CellRange): boolean {
  return (
    range.startRow > 0 &&
    range.startCol > 0 &&
    range.endRow >= range.startRow &&
    range.endCol >= range.startCol
  );
}
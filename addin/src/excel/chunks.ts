/**
 * Chunk planning for budgeted range I/O (INV-5). Pure — unit tested in Node.
 *
 * Splits an R x C rectangle into row bands of at most `maxCells` cells each.
 * If a single row is wider than `maxCells` (possible: a sheet row can span
 * 16384 columns), rows are additionally split into column bands.
 */

export interface Chunk {
  rowOffset: number;
  rowCount: number;
  colOffset: number;
  colCount: number;
}

export const MAX_CELLS_PER_SYNC = 10_000;
export const DEFAULT_MAX_CELLS_PER_READ = 100_000;

export function planChunks(
  rowCount: number,
  columnCount: number,
  maxCells: number = MAX_CELLS_PER_SYNC
): Chunk[] {
  if (rowCount < 1 || columnCount < 1) {
    throw new Error(`planChunks: invalid dimensions ${rowCount}x${columnCount}`);
  }
  if (maxCells < 1) {
    throw new Error(`planChunks: invalid maxCells ${maxCells}`);
  }

  const chunks: Chunk[] = [];
  const colBand = Math.min(columnCount, maxCells);
  const rowBand = Math.max(1, Math.floor(maxCells / colBand));

  for (let colOffset = 0; colOffset < columnCount; colOffset += colBand) {
    const colsHere = Math.min(colBand, columnCount - colOffset);
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += rowBand) {
      chunks.push({
        rowOffset,
        rowCount: Math.min(rowBand, rowCount - rowOffset),
        colOffset,
        colCount: colsHere,
      });
    }
  }
  return chunks;
}

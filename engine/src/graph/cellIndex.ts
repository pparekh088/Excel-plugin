/**
 * Sparse spatial index over a sheet's populated cells.
 *
 * Area resolution must never iterate a whole rectangle: `SUM(A:A)` covers
 * 1,048,576 addresses but perhaps 40 populated cells. Every lookup here is
 * proportional to the populated cells actually inside the area, chosen by
 * scanning whichever axis is cheaper.
 */

import { Sheet } from "../model/workbook";

function lowerBound(sorted: number[], target: number): number {
  let low = 0;
  let high = sorted.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (sorted[mid]! < target) low = mid + 1;
    else high = mid;
  }
  return low;
}

export class CellIndex {
  /** column -> ascending row indices of populated cells. */
  private readonly byColumn = new Map<number, number[]>();
  /** row -> ascending column indices of populated cells. */
  private readonly byRow = new Map<number, number[]>();

  constructor(sheet: Sheet) {
    const columns = new Map<number, number[]>();
    const rows = new Map<number, number[]>();
    for (const cell of sheet.cells.values()) {
      let colList = columns.get(cell.col);
      if (!colList) columns.set(cell.col, (colList = []));
      colList.push(cell.row);

      let rowList = rows.get(cell.row);
      if (!rowList) rows.set(cell.row, (rowList = []));
      rowList.push(cell.col);
    }
    for (const [col, list] of columns) {
      list.sort((a, b) => a - b);
      this.byColumn.set(col, list);
    }
    for (const [row, list] of rows) {
      list.sort((a, b) => a - b);
      this.byRow.set(row, list);
    }
  }

  get columns(): Iterable<number> {
    return this.byColumn.keys();
  }

  get rows(): Iterable<number> {
    return this.byRow.keys();
  }

  /**
   * Populated cells inside the rectangle. `null` bounds mean the whole axis
   * (full-column or full-row references).
   */
  *cellsIn(
    startRow: number | null,
    endRow: number | null,
    startCol: number | null,
    endCol: number | null
  ): Generator<[number, number]> {
    const rowLow = startRow ?? -Infinity;
    const rowHigh = endRow ?? Infinity;
    const colLow = startCol ?? -Infinity;
    const colHigh = endCol ?? Infinity;

    // Choose the cheaper axis to scan: bounded axes narrow the candidate set.
    const colSpan = colHigh - colLow;
    const rowSpan = rowHigh - rowLow;

    if (colSpan <= rowSpan) {
      const from = startCol ?? Math.min(...this.byColumn.keys(), 0);
      const to = endCol ?? Math.max(...this.byColumn.keys(), 0);
      for (let col = from; col <= to; col++) {
        const list = this.byColumn.get(col);
        if (!list) continue;
        for (let i = lowerBound(list, rowLow === -Infinity ? 0 : rowLow); i < list.length; i++) {
          const row = list[i]!;
          if (row > rowHigh) break;
          yield [row, col];
        }
      }
      return;
    }

    const from = startRow ?? Math.min(...this.byRow.keys(), 0);
    const to = endRow ?? Math.max(...this.byRow.keys(), 0);
    for (let row = from; row <= to; row++) {
      const list = this.byRow.get(row);
      if (!list) continue;
      for (let i = lowerBound(list, colLow === -Infinity ? 0 : colLow); i < list.length; i++) {
        const col = list[i]!;
        if (col > colHigh) break;
        yield [row, col];
      }
    }
  }
}

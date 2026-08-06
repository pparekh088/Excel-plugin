/**
 * Range-run collapsing (handoff §5.3).
 *
 * Contiguous cells sharing an R1C1-normalized formula are one logical node.
 * This is what makes a 200k-formula workbook tractable (a 10k-row model
 * column becomes 1 node, not 10k), and it is simultaneously the mechanism
 * behind AUD-001: a cell that breaks its neighbours' run IS the inconsistency.
 *
 * Rectangles are found greedily from the top-left: extend right while the
 * whole column strip matches, then extend down while whole rows match. That
 * reproduces exactly how people fill formulas (down a column, across a row,
 * or over a block) and never merges two cells that are not actually adjacent.
 */

export interface RunCell {
  row: number;
  col: number;
  signature: string;
}

export interface Run {
  signature: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  cellCount: number;
}

/**
 * Numeric cell key. Excel is 16384 columns wide, so (row, col) packs into one
 * number losslessly for every legal address. Strings here cost real memory at
 * scale — a 500k-formula workbook allocated hundreds of megabytes of
 * `"row,col"` keys before this changed.
 */
const COLUMN_STRIDE = 16_384;
const key = (row: number, col: number): number => row * COLUMN_STRIDE + col;

export function findRuns(cells: readonly RunCell[]): Run[] {
  const bySignature = new Map<string, RunCell[]>();
  for (const cell of cells) {
    let list = bySignature.get(cell.signature);
    if (!list) bySignature.set(cell.signature, (list = []));
    list.push(cell);
  }

  const runs: Run[] = [];
  for (const [signature, group] of bySignature) {
    const remaining = new Set(group.map((c) => key(c.row, c.col)));
    // Row-major order so we always start from a rectangle's top-left corner.
    const ordered = [...group].sort((a, b) => a.row - b.row || a.col - b.col);

    for (const cell of ordered) {
      if (!remaining.has(key(cell.row, cell.col))) continue;

      // Extend right along the starting row.
      let endCol = cell.col;
      while (remaining.has(key(cell.row, endCol + 1))) endCol++;

      // Extend down while every cell of the candidate row is present.
      let endRow = cell.row;
      for (;;) {
        const nextRow = endRow + 1;
        let complete = true;
        for (let col = cell.col; col <= endCol; col++) {
          if (!remaining.has(key(nextRow, col))) {
            complete = false;
            break;
          }
        }
        if (!complete) break;
        endRow = nextRow;
      }

      for (let row = cell.row; row <= endRow; row++) {
        for (let col = cell.col; col <= endCol; col++) remaining.delete(key(row, col));
      }

      runs.push({
        signature,
        startRow: cell.row,
        startCol: cell.col,
        endRow,
        endCol,
        cellCount: (endRow - cell.row + 1) * (endCol - cell.col + 1),
      });
    }
  }

  return runs.sort((a, b) => a.startRow - b.startRow || a.startCol - b.startCol);
}

/**
 * Runs that are adjacent along an axis and differ only by being separate —
 * used by AUD-001 to find the "odd cell out": a 1-cell run wedged between or
 * beside a much larger run in the same row/column band.
 */
export function neighbouringRuns(runs: readonly Run[], target: Run): Run[] {
  return runs.filter((run) => {
    if (run === target) return false;
    const verticallyAligned =
      run.startCol === target.startCol && run.endCol === target.endCol;
    const horizontallyAligned =
      run.startRow === target.startRow && run.endRow === target.endRow;
    if (verticallyAligned) {
      return run.endRow + 1 === target.startRow || target.endRow + 1 === run.startRow;
    }
    if (horizontallyAligned) {
      return run.endCol + 1 === target.startCol || target.endCol + 1 === run.startCol;
    }
    return false;
  });
}

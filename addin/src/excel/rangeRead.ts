/**
 * range.read executor — the only code that touches the workbook in Phase 0.
 *
 * INV-5: chunked reads (<= 10k cells per sync), hard cap on total cells,
 * tracked objects released as soon as each chunk lands. Reads `formulas`
 * (en-US), never `formulasLocal` — normalization to en-US is a system-wide
 * invariant (see PLATFORM_QUIRKS.md).
 */

import {
  DEFAULT_MAX_CELLS_PER_READ,
  MAX_CELLS_PER_SYNC,
  planChunks,
} from "./chunks";
import type { RangeReadParams, RangeReadResult } from "../tools/schemas";

export class ToolExecutionError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly detail?: Record<string, unknown>
  ) {
    super(message);
    this.name = "ToolExecutionError";
  }
}

type CellScalar = string | number | boolean | null;

function emptyGrid<T>(rows: number, cols: number, fill: T): T[][] {
  return Array.from({ length: rows }, () => new Array<T>(cols).fill(fill));
}

function blit<T>(target: T[][], source: T[][], rowOffset: number, colOffset: number): void {
  for (let r = 0; r < source.length; r++) {
    const sourceRow = source[r];
    const targetRow = target[rowOffset + r];
    if (!sourceRow || !targetRow) continue;
    for (let c = 0; c < sourceRow.length; c++) {
      targetRow[colOffset + c] = sourceRow[c] as T;
    }
  }
}

export async function executeRangeRead(params: RangeReadParams): Promise<RangeReadResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(params.sheet);
    const range = sheet.getRange(params.a1);
    range.load(["rowCount", "columnCount", "rowIndex", "columnIndex", "address"]);
    await context.sync();

    const { rowCount, columnCount, rowIndex, columnIndex, address } = range;
    const cellCount = rowCount * columnCount;
    const maxCells = params.maxCells ?? DEFAULT_MAX_CELLS_PER_READ;
    if (cellCount > maxCells) {
      throw new ToolExecutionError(
        "RANGE_TOO_LARGE",
        `Range ${address} has ${cellCount} cells; the budget for this read is ${maxCells}.`,
        { cellCount, maxCells, address }
      );
    }

    const include = new Set(params.include);
    const values = include.has("values") ? emptyGrid<CellScalar>(rowCount, columnCount, "") : null;
    const formulas = include.has("formulas")
      ? emptyGrid<CellScalar>(rowCount, columnCount, "")
      : null;
    const numberFormats = include.has("numberFormats")
      ? emptyGrid<string>(rowCount, columnCount, "General")
      : null;

    const loadProps: string[] = [];
    if (values) loadProps.push("values");
    if (formulas) loadProps.push("formulas");
    if (numberFormats) loadProps.push("numberFormat");

    const chunks = planChunks(rowCount, columnCount, MAX_CELLS_PER_SYNC);
    for (const chunk of chunks) {
      const sub = sheet.getRangeByIndexes(
        rowIndex + chunk.rowOffset,
        columnIndex + chunk.colOffset,
        chunk.rowCount,
        chunk.colCount
      );
      sub.load(loadProps);
      await context.sync();

      if (values) blit(values, sub.values as CellScalar[][], chunk.rowOffset, chunk.colOffset);
      if (formulas) {
        blit(formulas, sub.formulas as CellScalar[][], chunk.rowOffset, chunk.colOffset);
      }
      if (numberFormats) {
        blit(numberFormats, sub.numberFormat as string[][], chunk.rowOffset, chunk.colOffset);
      }
      // Release per-chunk proxy objects promptly (INV-5).
      sub.untrack();
    }

    return {
      sheet: params.sheet,
      a1: address,
      rowCount,
      columnCount,
      cellCount,
      chunkCount: chunks.length,
      ...(values ? { values } : {}),
      ...(formulas ? { formulas } : {}),
      ...(numberFormats ? { numberFormats } : {}),
    };
  });
}

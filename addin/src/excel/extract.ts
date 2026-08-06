/**
 * Bulk workbook extraction — builds the engine's Workbook model from a live
 * Excel session (INV-5: chunked, budgeted, tracked objects released).
 *
 * Reads `formulas` (always en-US, never `formulasLocal` — see Q-003) plus
 * `values`, and normalizes the Office.js quirk that the formulas grid carries
 * raw VALUES for non-formula cells (Q-004).
 *
 * Progress is reported per sheet so the task pane can show something moving
 * on a 200k-formula workbook instead of freezing.
 */

import { Workbook, type CellValue } from "ledger-engine";
import { MAX_CELLS_PER_SYNC, planChunks } from "./chunks";

export interface ExtractProgress {
  phase: "sheets" | "cells" | "metadata" | "done";
  sheetName?: string;
  sheetIndex: number;
  sheetCount: number;
  cellsRead: number;
}

export interface ExtractOptions {
  onProgress?: (progress: ExtractProgress) => void;
  /** Refuse to read a sheet larger than this many cells. */
  maxCellsPerSheet?: number;
  /** Include number formats (slower; needed for AUD-010). */
  includeNumberFormats?: boolean;
  /** Abort signal for cancel-from-UI. */
  signal?: AbortSignal;
}

const DEFAULT_MAX_CELLS_PER_SHEET = 2_000_000;

function isFormula(value: unknown): value is string {
  return typeof value === "string" && (value.startsWith("=") || value.startsWith("{="));
}

/** Office.js reports empty cells as "" — collapse to null for the model. */
function normalizeValue(value: unknown): CellValue {
  if (value === "" || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  return null;
}

export async function extractWorkbook(options: ExtractOptions = {}): Promise<Workbook> {
  const maxCells = options.maxCellsPerSheet ?? DEFAULT_MAX_CELLS_PER_SHEET;
  const workbook = new Workbook("Workbook");
  let cellsRead = 0;

  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/visibility,items/position");
    await context.sync();

    const sheetInfos = worksheets.items.map((sheet) => ({
      name: sheet.name,
      visible: sheet.visibility === Excel.SheetVisibility.visible,
    }));

    options.onProgress?.({
      phase: "sheets",
      sheetIndex: 0,
      sheetCount: sheetInfos.length,
      cellsRead,
    });

    for (const [index, info] of sheetInfos.entries()) {
      if (options.signal?.aborted) throw new Error("Extraction cancelled");

      const sheet = workbook.addSheet(info.name);
      sheet.visible = info.visible;

      // Used range first: never load a whole sheet blindly (INV-5).
      const worksheet = context.workbook.worksheets.getItem(info.name);
      const used = worksheet.getUsedRangeOrNullObject(true);
      used.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
      const protection = worksheet.protection;
      protection.load("protected");
      await context.sync();

      sheet.protectedSheet = protection.protected;

      if (used.isNullObject) {
        options.onProgress?.({
          phase: "cells",
          sheetName: info.name,
          sheetIndex: index + 1,
          sheetCount: sheetInfos.length,
          cellsRead,
        });
        continue;
      }

      const { rowIndex, columnIndex, rowCount, columnCount } = used;
      if (rowCount * columnCount > maxCells) {
        throw new Error(
          `Sheet "${info.name}" has ${rowCount * columnCount} cells in its used range, ` +
            `over the ${maxCells} budget. Narrow the scope and try again.`
        );
      }

      for (const chunk of planChunks(rowCount, columnCount, MAX_CELLS_PER_SYNC)) {
        if (options.signal?.aborted) throw new Error("Extraction cancelled");

        const range = worksheet.getRangeByIndexes(
          rowIndex + chunk.rowOffset,
          columnIndex + chunk.colOffset,
          chunk.rowCount,
          chunk.colCount
        );
        const props = ["values", "formulas"];
        if (options.includeNumberFormats) props.push("numberFormat");
        range.load(props);
        await context.sync();

        const values = range.values as unknown[][];
        const formulas = range.formulas as unknown[][];
        const numberFormats = options.includeNumberFormats
          ? (range.numberFormat as unknown[][])
          : undefined;

        for (let r = 0; r < chunk.rowCount; r++) {
          const valueRow = values[r];
          const formulaRow = formulas[r];
          if (!valueRow || !formulaRow) continue;
          for (let c = 0; c < chunk.colCount; c++) {
            const rawValue = normalizeValue(valueRow[c]);
            const rawFormula = formulaRow[c];
            const hasFormula = isFormula(rawFormula);
            // Skip genuinely empty cells: the used range is inflated by
            // formatting, and storing blanks would bloat the model.
            if (!hasFormula && rawValue === null) continue;

            const numberFormat = numberFormats?.[r]?.[c];
            sheet.set({
              row: rowIndex + chunk.rowOffset + r,
              col: columnIndex + chunk.colOffset + c,
              value: rawValue,
              ...(hasFormula ? { formula: rawFormula } : {}),
              ...(typeof numberFormat === "string" ? { numberFormat } : {}),
            });
            cellsRead++;
          }
        }
        range.untrack();
      }

      options.onProgress?.({
        phase: "cells",
        sheetName: info.name,
        sheetIndex: index + 1,
        sheetCount: sheetInfos.length,
        cellsRead,
      });
    }

    // ---- metadata: names, tables, charts, pivots -----------------------
    options.onProgress?.({
      phase: "metadata",
      sheetIndex: sheetInfos.length,
      sheetCount: sheetInfos.length,
      cellsRead,
    });

    await extractMetadata(context, workbook, sheetInfos.map((info) => info.name));
  });

  options.onProgress?.({
    phase: "done",
    sheetIndex: workbook.sheets.length,
    sheetCount: workbook.sheets.length,
    cellsRead,
  });
  return workbook;
}

async function extractMetadata(
  context: Excel.RequestContext,
  workbook: Workbook,
  sheetNames: string[]
): Promise<void> {
  const names = context.workbook.names;
  names.load("items/name,items/formula,items/scope,items/comment");
  const tables = context.workbook.tables;
  tables.load("items/name,items/showTotals,items/worksheet/name");
  await context.sync();

  for (const name of names.items) {
    workbook.names.push({
      name: name.name,
      scope: null,
      refersTo: typeof name.formula === "string" ? name.formula : String(name.formula ?? ""),
      ...(name.comment ? { comment: name.comment } : {}),
    });
  }

  // Table geometry needs a second round trip per table.
  const tableRanges = tables.items.map((table) => {
    const range = table.getRange();
    range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    const header = table.getHeaderRowRange();
    header.load(["rowIndex", "values"]);
    return { table, range, header };
  });
  await context.sync();

  for (const { table, range, header } of tableRanges) {
    const headerValues = (header.values as unknown[][])[0] ?? [];
    workbook.tables.push({
      name: table.name,
      sheet: table.worksheet.name,
      headerRow: header.rowIndex,
      startRow: range.rowIndex,
      endRow: range.rowIndex + range.rowCount - 1,
      startCol: range.columnIndex,
      endCol: range.columnIndex + range.columnCount - 1,
      hasTotals: table.showTotals,
      columns: headerValues.map((value, offset) => ({
        name: String(value ?? ""),
        col: range.columnIndex + offset,
      })),
    });
  }

  // Charts and pivots, per sheet. PivotTable source ranges are not exposed
  // for OLAP/PowerPivot sources (handoff §4 known gap) — recorded as such.
  for (const sheetName of sheetNames) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const charts = worksheet.charts;
    charts.load("items/name");
    const pivots = worksheet.pivotTables;
    pivots.load("items/name");
    await context.sync();

    for (const chart of charts.items) {
      workbook.charts.push({ name: chart.name, sheet: sheetName, sourceRanges: [] });
    }
    for (const pivot of pivots.items) {
      workbook.pivots.push({
        name: pivot.name,
        sheet: sheetName,
        sourceRange: "",
        // The API cannot inspect OLAP pivots; we surface them rather than
        // pretending we understand their sources.
        olap: true,
      });
    }
  }
}

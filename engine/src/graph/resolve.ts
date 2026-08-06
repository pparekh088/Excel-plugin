/**
 * Reference resolution: extracted AST references -> concrete sheet rectangles.
 *
 * Handles defined names (including names that refer to other names), 3D sheet
 * spans (one rectangle per sheet in the span), and structured table refs
 * (column spans, #Headers/#Totals/#Data selectors, [@this row]).
 *
 * Anything that cannot be resolved is reported as an `unresolved` entry
 * rather than silently dropped — the WIL surfaces the count so we never claim
 * coverage we do not have (INV-4).
 */

import { parseFormula } from "../parser/parser";
import { AreaRef, ExtractedRefs, NameUse, StructuredUse } from "../parser/refs";
import { extractRefs } from "../parser/refs";
import { SheetSpec } from "../parser/ast";
import { TableDef, Workbook, parseA1Range } from "../model/workbook";

export interface ResolvedArea {
  sheet: string;
  /** null = the whole axis (full column/row reference). */
  startRow: number | null;
  endRow: number | null;
  startCol: number | null;
  endCol: number | null;
}

export interface ResolutionOutcome {
  areas: ResolvedArea[];
  unresolved: string[];
}

export interface ResolveContext {
  workbook: Workbook;
  /** Sheet the formula lives on — resolves unqualified refs and name scope. */
  hostSheet: string;
  /** Row of the host cell (the run's anchor), for [@ThisRow] structured refs. */
  hostRow: number;
  /**
   * Extent of the collapsed run this formula represents, as extra rows/cols
   * beyond the anchor. Relative references are expanded by these amounts so a
   * run node carries the union of every member cell's precedents — without
   * this, a filled row would only report the leftmost cell's dependencies.
   */
  spanRows?: number;
  spanCols?: number;
  /** Last row of the run, so [@ThisRow] can cover the whole table body. */
  hostRowEnd?: number;
}

const MAX_NAME_DEPTH = 8;

function sheetNames(workbook: Workbook, spec: SheetSpec, hostSheet: string): string[] {
  if (spec.external !== null) return []; // external workbooks are not resolvable
  if (spec.start === null) return [hostSheet];
  if (spec.end !== null) return workbook.sheetsInSpan(spec.start, spec.end).map((s) => s.name);
  const sheet = workbook.sheet(spec.start);
  return sheet ? [sheet.name] : [];
}

function resolveAreaRef(area: AreaRef, context: ResolveContext): ResolutionOutcome {
  const names = sheetNames(context.workbook, area.sheet, context.hostSheet);
  if (names.length === 0) {
    const label = area.sheet.external !== null
      ? `[${area.sheet.external}]${area.sheet.start ?? ""}`
      : (area.sheet.start ?? "");
    return { areas: [], unresolved: [`${label}!<range>`] };
  }
  // A run of N filled cells reads the union of what each member reads:
  // relative bounds sweep with the fill, absolute ($) bounds stay put.
  const spanRows = context.spanRows ?? 0;
  const spanCols = context.spanCols ?? 0;
  const endRow =
    area.endRow === null || area.endRowAbs ? area.endRow : area.endRow + spanRows;
  const endCol =
    area.endCol === null || area.endColAbs ? area.endCol : area.endCol + spanCols;

  return {
    areas: names.map((sheet) => ({
      sheet,
      startRow: area.startRow,
      endRow,
      startCol: area.startCol,
      endCol,
    })),
    unresolved: [],
  };
}

function resolveName(use: NameUse, context: ResolveContext, depth: number): ResolutionOutcome {
  if (depth > MAX_NAME_DEPTH) {
    return { areas: [], unresolved: [`${use.name} (name nesting too deep)`] };
  }
  const scope = use.sheet.start ?? context.hostSheet;
  const defined = context.workbook.definedName(use.name, scope);
  if (!defined) return { areas: [], unresolved: [use.name] };

  // refersTo is a formula ("=Sheet1!$A$1:$B$5", "=OFFSET(...)", "=42").
  const parsed = parseFormula(
    defined.refersTo.startsWith("=") ? defined.refersTo : `=${defined.refersTo}`
  );
  const refs = extractRefs(parsed.ast);
  const nested: ResolveContext = {
    ...context,
    hostSheet: defined.scope ?? context.hostSheet,
  };
  return resolveExtracted(refs, nested, depth + 1);
}

function tableArea(
  table: TableDef,
  use: StructuredUse,
  context: ResolveContext
): ResolvedArea | null {
  let startCol = table.startCol;
  let endCol = table.endCol;
  if (use.columns.length > 0) {
    const indices = use.columns.map((columnName) => {
      const column = table.columns.find(
        (c) => c.name.toUpperCase() === columnName.toUpperCase()
      );
      return column ? column.col : null;
    });
    if (indices.some((index) => index === null)) return null;
    const numeric = indices as number[];
    startCol = Math.min(...numeric);
    endCol = Math.max(...numeric);
  }

  const dataStart = table.headerRow + 1;
  const dataEnd = table.hasTotals ? table.endRow - 1 : table.endRow;

  let startRow = dataStart;
  let endRow = dataEnd;

  if (use.thisRow) {
    // [@Col] resolves to the host row — or, for a run filled down the table,
    // every row the run covers.
    const lastRow = context.hostRowEnd ?? context.hostRow;
    if (context.hostRow < dataStart || context.hostRow > dataEnd) return null;
    startRow = context.hostRow;
    endRow = Math.min(lastRow, dataEnd);
  } else if (use.items.includes("#All")) {
    startRow = table.startRow;
    endRow = table.endRow;
  } else if (use.items.includes("#Headers")) {
    startRow = table.headerRow;
    endRow = table.headerRow;
  } else if (use.items.includes("#Totals")) {
    if (!table.hasTotals) return null;
    startRow = table.endRow;
    endRow = table.endRow;
  }

  return { sheet: table.sheet, startRow, endRow, startCol, endCol };
}

function resolveStructured(use: StructuredUse, context: ResolveContext): ResolutionOutcome {
  // A bare [@Col] refers to the table containing the host cell.
  const table = use.table
    ? context.workbook.table(use.table)
    : context.workbook.tables.find(
        (candidate) =>
          candidate.sheet.toUpperCase() === context.hostSheet.toUpperCase() &&
          context.hostRow >= candidate.startRow &&
          context.hostRow <= candidate.endRow
      );
  if (!table) {
    return { areas: [], unresolved: [`${use.table || "<host table>"}[...]`] };
  }
  const area = tableArea(table, use, context);
  if (!area) {
    return {
      areas: [],
      unresolved: [`${table.name}[${use.columns.join(",") || use.items.join(",")}]`],
    };
  }
  return { areas: [area], unresolved: [] };
}

export function resolveExtracted(
  refs: ExtractedRefs,
  context: ResolveContext,
  depth = 0
): ResolutionOutcome {
  const areas: ResolvedArea[] = [];
  const unresolved: string[] = [];

  for (const area of refs.areas) {
    const outcome = resolveAreaRef(area, context);
    areas.push(...outcome.areas);
    unresolved.push(...outcome.unresolved);
  }
  for (const use of refs.names) {
    const outcome = resolveName(use, context, depth);
    areas.push(...outcome.areas);
    unresolved.push(...outcome.unresolved);
  }
  for (const use of refs.structured) {
    const outcome = resolveStructured(use, context);
    areas.push(...outcome.areas);
    unresolved.push(...outcome.unresolved);
  }
  return { areas, unresolved };
}

/** Resolve a defined name's refersTo directly to areas (used by the WIL). */
export function resolveRefersTo(
  workbook: Workbook,
  refersTo: string,
  hostSheet: string
): ResolvedArea[] {
  const direct = parseA1Range(refersTo.replace(/^=/, ""));
  if (direct) {
    return [
      {
        sheet: direct.sheet ?? hostSheet,
        startRow: direct.startRow,
        endRow: direct.endRow,
        startCol: direct.startCol,
        endCol: direct.endCol,
      },
    ];
  }
  const parsed = parseFormula(refersTo.startsWith("=") ? refersTo : `=${refersTo}`);
  return resolveExtracted(extractRefs(parsed.ast), { workbook, hostSheet, hostRow: 0 }).areas;
}

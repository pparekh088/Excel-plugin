/**
 * Reference extraction from parsed formulas — the graph builder's input.
 *
 * Honesty rules (INV-4 / handoff §5.2):
 * - INDIRECT/OFFSET and computed range endpoints (INDEX(..):B5) are OPAQUE:
 *   their literal ref arguments are still extracted as edges, but the formula
 *   is flagged `opaque` so the graph never claims full coverage over it.
 * - LET/LAMBDA locals shadow workbook names and are NOT emitted as name refs.
 * - #REF! inside a formula is counted as a broken reference.
 */

import { Node, SheetSpec, StructuredRefNode } from "./ast";

export type AreaKind = "cell" | "range" | "colRange" | "rowRange";

export interface AreaRef {
  kind: AreaKind;
  sheet: SheetSpec;
  /** 0-based inclusive bounds; rows/cols may be null for full col/row ranges. */
  startRow: number | null;
  endRow: number | null;
  startCol: number | null;
  endCol: number | null;
  /**
   * Absoluteness of each bound ($). Required to expand a reference across a
   * collapsed run: relative bounds shift with the fill, absolute ones pin.
   */
  startRowAbs: boolean;
  endRowAbs: boolean;
  startColAbs: boolean;
  endColAbs: boolean;
}

export interface NameUse {
  name: string;
  sheet: SheetSpec;
}

export interface StructuredUse {
  table: string;
  columns: string[];
  items: string[];
  thisRow: boolean;
}

export interface ExtractedRefs {
  areas: AreaRef[];
  names: NameUse[];
  structured: StructuredUse[];
  /** Functions used, canonical upper-case names, deduplicated. */
  functions: string[];
  /** Formula contains INDIRECT/OFFSET or a computed range endpoint. */
  opaque: boolean;
  /** Formula uses a volatile function. */
  volatile: boolean;
  /** #REF! literals inside the formula. */
  brokenRefs: number;
  /** References into other workbooks ([Book.xlsx]Sheet!A1). */
  external: boolean;
  /** 3D sheet spans (Sheet1:Sheet3!A1). */
  threeD: boolean;
  /** Spill references (A1#). */
  spills: number;
}

export const VOLATILE_FUNCTIONS = new Set([
  "NOW",
  "TODAY",
  "RAND",
  "RANDBETWEEN",
  "RANDARRAY",
  "OFFSET",
  "INDIRECT",
  "INFO",
  "CELL",
]);

export const OPAQUE_FUNCTIONS = new Set(["INDIRECT", "OFFSET"]);

/** Order bounds ascending, carrying each bound's $ flag with it. */
function normalizeArea(area: AreaRef): AreaRef {
  let { startRow, endRow, startCol, endCol } = area;
  let { startRowAbs, endRowAbs, startColAbs, endColAbs } = area;
  if (startRow !== null && endRow !== null && startRow > endRow) {
    [startRow, endRow] = [endRow, startRow];
    [startRowAbs, endRowAbs] = [endRowAbs, startRowAbs];
  }
  if (startCol !== null && endCol !== null && startCol > endCol) {
    [startCol, endCol] = [endCol, startCol];
    [startColAbs, endColAbs] = [endColAbs, startColAbs];
  }
  return {
    ...area,
    startRow,
    endRow,
    startCol,
    endCol,
    startRowAbs,
    endRowAbs,
    startColAbs,
    endColAbs,
  };
}

export function extractRefs(root: Node): ExtractedRefs {
  const out: ExtractedRefs = {
    areas: [],
    names: [],
    structured: [],
    functions: [],
    opaque: false,
    volatile: false,
    brokenRefs: 0,
    external: false,
    threeD: false,
    spills: 0,
  };
  const functions = new Set<string>();

  const noteSheet = (sheet: SheetSpec) => {
    if (sheet.external !== null) out.external = true;
    if (sheet.end !== null) out.threeD = true;
  };

  const visit = (node: Node, locals: ReadonlySet<string>): void => {
    switch (node.kind) {
      case "cell":
        noteSheet(node.sheet);
        out.areas.push(
          normalizeArea({
            kind: "cell",
            sheet: node.sheet,
            startRow: node.addr.row,
            endRow: node.addr.row,
            startCol: node.addr.col,
            endCol: node.addr.col,
            startRowAbs: node.addr.rowAbs,
            endRowAbs: node.addr.rowAbs,
            startColAbs: node.addr.colAbs,
            endColAbs: node.addr.colAbs,
          })
        );
        return;
      case "range":
        noteSheet(node.sheet);
        out.areas.push(
          normalizeArea({
            kind: "range",
            sheet: node.sheet,
            startRow: node.start.row,
            endRow: node.end.row,
            startCol: node.start.col,
            endCol: node.end.col,
            startRowAbs: node.start.rowAbs,
            endRowAbs: node.end.rowAbs,
            startColAbs: node.start.colAbs,
            endColAbs: node.end.colAbs,
          })
        );
        return;
      case "colRange":
        noteSheet(node.sheet);
        out.areas.push(
          normalizeArea({
            kind: "colRange",
            sheet: node.sheet,
            startRow: null,
            endRow: null,
            startCol: node.startCol,
            endCol: node.endCol,
            startRowAbs: true,
            endRowAbs: true,
            startColAbs: node.startAbs,
            endColAbs: node.endAbs,
          })
        );
        return;
      case "rowRange":
        noteSheet(node.sheet);
        out.areas.push(
          normalizeArea({
            kind: "rowRange",
            sheet: node.sheet,
            startRow: node.startRow,
            endRow: node.endRow,
            startCol: null,
            endCol: null,
            startRowAbs: node.startAbs,
            endRowAbs: node.endAbs,
            startColAbs: true,
            endColAbs: true,
          })
        );
        return;
      case "name": {
        noteSheet(node.sheet);
        if (node.sheet.start === null && locals.has(node.name.toUpperCase())) return;
        out.names.push({ name: node.name, sheet: node.sheet });
        return;
      }
      case "structured":
        out.structured.push(structuredUse(node));
        return;
      case "error":
        if (node.value === "#REF!") out.brokenRefs++;
        return;
      case "spill":
        out.spills++;
        visit(node.operand, locals);
        return;
      case "unary":
      case "percent":
      case "implicitIntersection":
        visit(node.operand, locals);
        return;
      case "group":
        visit(node.expr, locals);
        return;
      case "binary": {
        if (node.op === ":") {
          // Computed range endpoint (INDEX(..):B5 etc.) — opaque coverage.
          out.opaque = true;
        }
        visit(node.left, locals);
        visit(node.right, locals);
        return;
      }
      case "array":
        for (const row of node.rows) for (const item of row) visit(item, locals);
        return;
      case "func":
        visitFunc(node.name, node.args, locals);
        return;
      case "callExpr":
        visit(node.callee, locals);
        for (const arg of node.args) visit(arg, locals);
        return;
      default:
        return;
    }
  };

  const visitFunc = (name: string, args: Node[], locals: ReadonlySet<string>): void => {
    functions.add(name);
    if (VOLATILE_FUNCTIONS.has(name)) out.volatile = true;
    if (OPAQUE_FUNCTIONS.has(name)) out.opaque = true;

    if (name === "LET") {
      // LET(name1, value1, [name2, value2, ...], body)
      const scoped = new Set(locals);
      let index = 0;
      while (index + 1 < args.length) {
        const nameArg = args[index]!;
        const valueArg = args[index + 1]!;
        visit(valueArg, scoped); // value may use previously bound locals
        if (nameArg.kind === "name") scoped.add(nameArg.name.toUpperCase());
        index += 2;
      }
      if (index < args.length) visit(args[index]!, scoped); // body
      return;
    }
    if (name === "LAMBDA") {
      // LAMBDA(p1, ..., pn, body)
      const scoped = new Set(locals);
      for (let i = 0; i < args.length - 1; i++) {
        const param = args[i]!;
        if (param.kind === "name") scoped.add(param.name.toUpperCase());
      }
      if (args.length > 0) visit(args[args.length - 1]!, scoped);
      return;
    }
    for (const arg of args) visit(arg, locals);
  };

  visit(root, new Set());
  out.functions = [...functions].sort();
  return out;
}

function structuredUse(node: StructuredRefNode): StructuredUse {
  return {
    table: node.table,
    columns: node.columns,
    items: node.items,
    thisRow: node.thisRow,
  };
}

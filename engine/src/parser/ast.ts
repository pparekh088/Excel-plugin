/**
 * Formula AST (INV-4: we parse every formula ourselves; the graph is built
 * from these trees, never from getPrecedents()).
 *
 * Design rules:
 * - Parsing NEVER throws. Malformed input yields a `bad` node plus
 *   diagnostics, so bulk extraction always completes and coverage gaps are
 *   reported honestly (WIL stats count unparsed formulas).
 * - Grammar is en-US A1 notation — exactly what Office.js `Range.formulas`
 *   returns. R1C1 exists only as our normalized *output* for run detection.
 */

// ---------------------------------------------------------------- sheets

export interface SheetSpec {
  /** null = unprefixed (host sheet). For 3D spans, `end` is set. */
  start: string | null;
  end: string | null;
  /** Workbook qualifier: "Book1.xlsx" from [Book1.xlsx]Sheet1!A1, or "1" from [1]Sheet1!. */
  external: string | null;
}

export const NO_SHEET: SheetSpec = { start: null, end: null, external: null };

// ---------------------------------------------------------------- addresses

export interface CellAddr {
  /** 0-based. */
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

// ---------------------------------------------------------------- nodes

export type Node =
  | NumberLit
  | StringLit
  | BoolLit
  | ErrorLit
  | CellRefNode
  | RangeRefNode
  | ColRangeNode
  | RowRangeNode
  | NameRefNode
  | StructuredRefNode
  | ArrayLit
  | UnaryNode
  | PercentNode
  | BinaryNode
  | FuncCallNode
  | CallExprNode
  | GroupNode
  | MissingArg
  | ImplicitIntersectionNode
  | SpillRefNode
  | BadNode;

export interface NumberLit {
  kind: "number";
  value: number;
}

export interface StringLit {
  kind: "string";
  value: string;
}

export interface BoolLit {
  kind: "bool";
  value: boolean;
}

export interface ErrorLit {
  kind: "error";
  /** Canonical error text, e.g. "#REF!", "#DIV/0!", "#N/A". */
  value: string;
}

export interface CellRefNode {
  kind: "cell";
  sheet: SheetSpec;
  addr: CellAddr;
}

/** Rectangular range folded from `cell:cell` with compatible sheets. */
export interface RangeRefNode {
  kind: "range";
  sheet: SheetSpec;
  start: CellAddr;
  end: CellAddr;
}

export interface ColRangeNode {
  kind: "colRange";
  sheet: SheetSpec;
  startCol: number;
  endCol: number;
  startAbs: boolean;
  endAbs: boolean;
}

export interface RowRangeNode {
  kind: "rowRange";
  sheet: SheetSpec;
  startRow: number;
  endRow: number;
  startAbs: boolean;
  endAbs: boolean;
}

export interface NameRefNode {
  kind: "name";
  sheet: SheetSpec;
  name: string;
}

export type StructuredItem =
  | "#All"
  | "#Data"
  | "#Headers"
  | "#Totals"
  | "#This Row";

export interface StructuredRefNode {
  kind: "structured";
  /** Empty string for bare [@Col] / [Col] inside the host table. */
  table: string;
  items: StructuredItem[];
  /** Column span: [start] or [start, end]. Empty = whole table/items only. */
  columns: string[];
  thisRow: boolean;
  raw: string;
}

export interface ArrayLit {
  kind: "array";
  /** rows x cols of literal scalars. */
  rows: Node[][];
}

export type UnaryOp = "-" | "+";

export interface UnaryNode {
  kind: "unary";
  op: UnaryOp;
  operand: Node;
}

export interface PercentNode {
  kind: "percent";
  operand: Node;
}

export type BinaryOp =
  | "+"
  | "-"
  | "*"
  | "/"
  | "^"
  | "&"
  | "="
  | "<>"
  | "<"
  | "<="
  | ">"
  | ">="
  | ":" // range operator over non-simple operands (e.g. INDEX(..):B5)
  | "," // reference union
  | " "; // reference intersection

export interface BinaryNode {
  kind: "binary";
  op: BinaryOp;
  left: Node;
  right: Node;
}

export interface FuncCallNode {
  kind: "func";
  /** Canonical upper-case name, _xlfn. prefix stripped (kept in `rawName`). */
  name: string;
  rawName: string;
  args: Node[];
}

/** Calling an expression: LAMBDA(x,x*2)(A1), (myLambda)(1). */
export interface CallExprNode {
  kind: "callExpr";
  callee: Node;
  args: Node[];
}

export interface GroupNode {
  kind: "group";
  expr: Node;
}

/** Empty argument slot: IF(A1,,B1). */
export interface MissingArg {
  kind: "missing";
}

export interface ImplicitIntersectionNode {
  kind: "implicitIntersection";
  operand: Node;
}

/** Spilled-range operator: A1#, Name#. */
export interface SpillRefNode {
  kind: "spill";
  operand: Node;
}

/** Unparseable region — raw text preserved, never thrown away. */
export interface BadNode {
  kind: "bad";
  raw: string;
  message: string;
}

// ------------------------------------------------------------- diagnostics

export interface Diagnostic {
  message: string;
  /** Character offset into the original formula (after any leading '='). */
  position: number;
}

export interface ParseResult {
  /** Original text as given (with '=' if present). */
  source: string;
  ast: Node;
  diagnostics: Diagnostic[];
  /** True when the formula parsed cleanly (no bad nodes, no diagnostics). */
  ok: boolean;
}

// ---------------------------------------------------------------- helpers

export function isRefValued(node: Node): boolean {
  switch (node.kind) {
    case "cell":
    case "range":
    case "colRange":
    case "rowRange":
    case "name":
    case "structured":
    case "spill":
      return true;
    case "implicitIntersection":
      return isRefValued(node.operand);
    case "group":
      return isRefValued(node.expr);
    case "binary":
      return (
        (node.op === ":" || node.op === "," || node.op === " ") &&
        isRefValued(node.left) &&
        isRefValued(node.right)
      );
    default:
      return false;
  }
}

export function walk(node: Node, visit: (n: Node) => void): void {
  visit(node);
  switch (node.kind) {
    case "unary":
    case "percent":
      walk(node.operand, visit);
      break;
    case "implicitIntersection":
    case "spill":
      walk(node.operand, visit);
      break;
    case "binary":
      walk(node.left, visit);
      walk(node.right, visit);
      break;
    case "func":
      for (const arg of node.args) walk(arg, visit);
      break;
    case "callExpr":
      walk(node.callee, visit);
      for (const arg of node.args) walk(arg, visit);
      break;
    case "group":
      walk(node.expr, visit);
      break;
    case "array":
      for (const row of node.rows) for (const item of row) walk(item, visit);
      break;
    default:
      break;
  }
}

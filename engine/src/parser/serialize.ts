/**
 * AST → canonical en-US A1 text. Canonical form: upper-case function names,
 * no incidental whitespace (single space only for intersection), minimal
 * sheet quoting. parse(serialize(ast)) must be structurally identical —
 * round-trip tests enforce it.
 */

import { CellAddr, Node, SheetSpec } from "./ast";
import { colIndexToLetters } from "./tokenizer";

const PLAIN_SHEET_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;
// Sheet names that lex like cell refs (Q1, A1, XFD10) must be quoted or the
// serialized text re-parses as a cell/range instead of a sheet prefix.
const CELL_LIKE_RE = /^[A-Za-z]{1,3}[0-9]+$/;

function plainSheetName(name: string): boolean {
  return PLAIN_SHEET_RE.test(name) && !CELL_LIKE_RE.test(name);
}

export function sheetPrefix(sheet: SheetSpec): string {
  if (sheet.start === null && sheet.external === null) return "";
  const body =
    (sheet.external !== null ? `[${sheet.external}]` : "") +
    (sheet.start ?? "") +
    (sheet.end !== null ? `:${sheet.end}` : "");
  const needsQuotes =
    sheet.external !== null ||
    !plainSheetName(sheet.start ?? "") ||
    (sheet.end !== null && !plainSheetName(sheet.end));
  return needsQuotes ? `'${body.replace(/'/g, "''")}'!` : `${body}!`;
}

export function cellText(addr: CellAddr): string {
  return (
    (addr.colAbs ? "$" : "") +
    colIndexToLetters(addr.col) +
    (addr.rowAbs ? "$" : "") +
    String(addr.row + 1)
  );
}

export function serialize(node: Node): string {
  switch (node.kind) {
    case "number": {
      return formatNumber(node.value);
    }
    case "string":
      return `"${node.value.replace(/"/g, '""')}"`;
    case "bool":
      return node.value ? "TRUE" : "FALSE";
    case "error":
      return node.value;
    case "cell":
      return sheetPrefix(node.sheet) + cellText(node.addr);
    case "range":
      return sheetPrefix(node.sheet) + cellText(node.start) + ":" + cellText(node.end);
    case "colRange":
      return (
        sheetPrefix(node.sheet) +
        (node.startAbs ? "$" : "") +
        colIndexToLetters(node.startCol) +
        ":" +
        (node.endAbs ? "$" : "") +
        colIndexToLetters(node.endCol)
      );
    case "rowRange":
      return (
        sheetPrefix(node.sheet) +
        (node.startAbs ? "$" : "") +
        String(node.startRow + 1) +
        ":" +
        (node.endAbs ? "$" : "") +
        String(node.endRow + 1)
      );
    case "name":
      return sheetPrefix(node.sheet) + node.name;
    case "structured": {
      return serializeStructured(node.table, node.items, node.columns, node.thisRow);
    }
    case "array":
      return (
        "{" +
        node.rows.map((row) => row.map(serialize).join(",")).join(";") +
        "}"
      );
    case "unary":
      return node.op + serialize(node.operand);
    case "percent":
      return serialize(node.operand) + "%";
    case "binary": {
      const op = node.op === " " ? " " : node.op;
      return serialize(node.left) + op + serialize(node.right);
    }
    case "func":
      return node.name + "(" + node.args.map(serialize).join(",") + ")";
    case "callExpr":
      return serialize(node.callee) + "(" + node.args.map(serialize).join(",") + ")";
    case "group":
      return "(" + serialize(node.expr) + ")";
    case "missing":
      return "";
    case "implicitIntersection":
      return "@" + serialize(node.operand);
    case "spill":
      return serialize(node.operand) + "#";
    case "bad":
      return node.raw;
  }
}

export function serializeStructured(
  table: string,
  items: string[],
  columns: string[],
  thisRow: boolean
): string {
  const escape = (name: string) => name.replace(/([[\]#'@])/g, "'$1");
  const parts: string[] = [];
  for (const item of items) {
    if (item === "#This Row" && thisRow) continue;
    parts.push(`[${item}]`);
  }
  if (columns.length === 1) {
    parts.push(`[${escape(columns[0]!)}]`);
  } else if (columns.length >= 2) {
    parts.push(`[${escape(columns[0]!)}]:[${escape(columns[1]!)}]`);
  }
  const atPrefix = thisRow ? "@" : "";
  if (parts.length === 0) {
    return `${table}[${atPrefix}]`;
  }
  // Single column, no item specifiers: simple form, escaping in place
  // (Table[Col A], Table[@Amount], Table['#Col]).
  if (columns.length === 1 && items.length === 0) {
    return `${table}[${atPrefix}${escape(columns[0]!)}]`;
  }
  // Single item specifier, nothing else: Table[#All].
  if (items.length === 1 && columns.length === 0 && !thisRow) {
    return `${table}[${items[0]}]`;
  }
  return `${table}[${atPrefix}${parts.join(",")}]`;
}

export function formatNumber(value: number): string {
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
  return String(value);
}

/** Serialize with the leading '='. */
export function toFormulaText(node: Node): string {
  return "=" + serialize(node);
}

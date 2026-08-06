/**
 * R1C1 normalization — the run-detection signature (INV-4 / handoff §5.3).
 *
 * A formula is rendered relative to its host cell: relative refs become
 * offsets (R[1]C[-2]), absolute refs become fixed indices (R5C3). Two cells
 * whose formulas normalize to the same string are "the same formula" in the
 * fill/copy sense — the basis for range-run collapsing and AUD-001
 * (formula-inconsistency) detection.
 *
 * Canonicalization: function names upper-cased, defined names upper-cased
 * (Excel names are case-insensitive), string literals preserved, sheet
 * prefixes normalized.
 */

import { CellAddr, Node } from "./ast";
import { sheetPrefix } from "./serialize";

function r1c1Cell(addr: CellAddr, hostRow: number, hostCol: number): string {
  const rowPart = addr.rowAbs
    ? `R${addr.row + 1}`
    : addr.row === hostRow
      ? "R"
      : `R[${addr.row - hostRow}]`;
  const colPart = addr.colAbs
    ? `C${addr.col + 1}`
    : addr.col === hostCol
      ? "C"
      : `C[${addr.col - hostCol}]`;
  return rowPart + colPart;
}

export function normalizeR1C1(node: Node, hostRow: number, hostCol: number): string {
  const walk = (n: Node): string => {
    switch (n.kind) {
      case "number":
        return String(n.value);
      case "string":
        return `"${n.value.replace(/"/g, '""')}"`;
      case "bool":
        return n.value ? "TRUE" : "FALSE";
      case "error":
        return n.value;
      case "cell":
        return sheetPrefix(n.sheet) + r1c1Cell(n.addr, hostRow, hostCol);
      case "range":
        return (
          sheetPrefix(n.sheet) +
          r1c1Cell(n.start, hostRow, hostCol) +
          ":" +
          r1c1Cell(n.end, hostRow, hostCol)
        );
      case "colRange": {
        const part = (col: number, abs: boolean) =>
          abs ? `C${col + 1}` : col === hostCol ? "C" : `C[${col - hostCol}]`;
        return (
          sheetPrefix(n.sheet) + part(n.startCol, n.startAbs) + ":" + part(n.endCol, n.endAbs)
        );
      }
      case "rowRange": {
        const part = (row: number, abs: boolean) =>
          abs ? `R${row + 1}` : row === hostRow ? "R" : `R[${row - hostRow}]`;
        return (
          sheetPrefix(n.sheet) + part(n.startRow, n.startAbs) + ":" + part(n.endRow, n.endAbs)
        );
      }
      case "name":
        return sheetPrefix(n.sheet) + n.name.toUpperCase();
      case "structured":
        return n.raw.toUpperCase();
      case "array":
        return "{" + n.rows.map((row) => row.map(walk).join(",")).join(";") + "}";
      case "unary":
        return n.op + walk(n.operand);
      case "percent":
        return walk(n.operand) + "%";
      case "binary":
        return walk(n.left) + n.op + walk(n.right);
      case "func":
        return n.name + "(" + n.args.map(walk).join(",") + ")";
      case "callExpr":
        return walk(n.callee) + "(" + n.args.map(walk).join(",") + ")";
      case "group":
        return "(" + walk(n.expr) + ")";
      case "missing":
        return "";
      case "implicitIntersection":
        return "@" + walk(n.operand);
      case "spill":
        return walk(n.operand) + "#";
      case "bad":
        return `<bad:${n.raw}>`;
    }
  };
  return walk(node);
}

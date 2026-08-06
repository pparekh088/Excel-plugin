/**
 * Test helper: build workbooks from a compact literal notation.
 *
 *   sheet("Model", {
 *     A1: "Revenue", B1: 100, C1: "=B1*1.05",
 *   })
 */

import { Cell, CellValue, Workbook, parseA1Range } from "../../src/model/workbook";

export type CellSpec = Record<string, CellValue>;

export function addSheet(workbook: Workbook, name: string, cells: CellSpec): Workbook {
  const sheet = workbook.addSheet(name);
  for (const [address, raw] of Object.entries(cells)) {
    const parsed = parseA1Range(address);
    if (!parsed) throw new Error(`bad address in fixture: ${address}`);
    const isFormula = typeof raw === "string" && raw.startsWith("=");
    const cell: Cell = {
      row: parsed.startRow,
      col: parsed.startCol,
      value: isFormula ? 0 : raw,
      ...(isFormula ? { formula: raw } : {}),
    };
    sheet.set(cell);
  }
  return workbook;
}

export function workbookOf(sheets: Record<string, CellSpec>): Workbook {
  const workbook = new Workbook("Test");
  for (const [name, cells] of Object.entries(sheets)) addSheet(workbook, name, cells);
  return workbook;
}

/** Fill a formula down a column, mimicking a real fill operation. */
export function fillDown(
  template: (row: number) => string,
  startRow: number,
  endRow: number,
  col: string
): CellSpec {
  const out: CellSpec = {};
  for (let row = startRow; row <= endRow; row++) {
    out[`${col}${row}`] = template(row);
  }
  return out;
}

/**
 * Workbook model — the in-memory representation the whole engine operates on.
 *
 * It is populated three ways, all producing the same shape:
 *  1. Office.js bulk extraction in the add-in (real workbooks).
 *  2. The corpus generators (eval fixtures).
 *  3. Deserialization of a stored fingerprint.
 *
 * Cells are stored sparsely per sheet keyed by "row,col" — real workbooks are
 * mostly empty, and the used range is a lie (formatting inflates it).
 */

import { colIndexToLetters } from "../parser/tokenizer";

export type CellValue = string | number | boolean | null;

export interface Cell {
  row: number;
  col: number;
  /** Computed value as Excel reports it. */
  value: CellValue;
  /**
   * Formula text including '=' when the cell holds a formula; otherwise
   * undefined. (Office.js returns the raw value in the formulas grid for
   * non-formula cells — extraction normalizes that away here.)
   */
  formula?: string;
  numberFormat?: string;
}

export interface DefinedName {
  name: string;
  /** null = workbook scope. */
  scope: string | null;
  refersTo: string;
  comment?: string;
}

export interface TableColumn {
  name: string;
  /** 0-based column index within the sheet. */
  col: number;
}

export interface TableDef {
  name: string;
  sheet: string;
  /** Header row index; data starts at headerRow + 1. */
  headerRow: number;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  columns: TableColumn[];
  hasTotals: boolean;
}

export interface ChartDef {
  name: string;
  sheet: string;
  /** A1 ranges the chart plots (may be empty when unknown). */
  sourceRanges: string[];
}

export interface PivotDef {
  name: string;
  sheet: string;
  sourceRange: string;
  /** true when the pivot is OLAP/PowerPivot backed (API cannot inspect it). */
  olap: boolean;
}

/**
 * Excel is 16384 columns wide, so (row, col) packs losslessly into one number.
 * Numeric keys rather than `"row,col"` strings: a 750k-cell workbook allocated
 * roughly a third of a gigabyte in key strings alone before this changed.
 */
const COLUMN_STRIDE = 16_384;

export class Sheet {
  readonly cells = new Map<number, Cell>();
  /** Sheets hidden from the user still participate in calculation. */
  visible = true;
  protectedSheet = false;
  /** Merged areas as [startRow, startCol, endRow, endCol]. */
  merged: Array<[number, number, number, number]> = [];

  constructor(
    /** Mutable only so rollback can undo a rename (see Workbook.renameSheet). */
    public name: string,
    readonly index: number
  ) {}

  static key(row: number, col: number): number {
    return row * COLUMN_STRIDE + col;
  }

  get(row: number, col: number): Cell | undefined {
    return this.cells.get(Sheet.key(row, col));
  }

  set(cell: Cell): void {
    this.cells.set(Sheet.key(cell.row, cell.col), cell);
  }

  delete(row: number, col: number): void {
    this.cells.delete(Sheet.key(row, col));
  }

  /** Inclusive bounds of populated cells, or null when the sheet is empty. */
  bounds(): { startRow: number; endRow: number; startCol: number; endCol: number } | null {
    let startRow = Infinity;
    let endRow = -Infinity;
    let startCol = Infinity;
    let endCol = -Infinity;
    for (const cell of this.cells.values()) {
      if (cell.row < startRow) startRow = cell.row;
      if (cell.row > endRow) endRow = cell.row;
      if (cell.col < startCol) startCol = cell.col;
      if (cell.col > endCol) endCol = cell.col;
    }
    if (endRow === -Infinity) return null;
    return { startRow, endRow, startCol, endCol };
  }
}

export class Workbook {
  readonly sheets: Sheet[] = [];
  readonly names: DefinedName[] = [];
  readonly tables: TableDef[] = [];
  readonly charts: ChartDef[] = [];
  readonly pivots: PivotDef[] = [];

  constructor(readonly name = "Workbook") {}

  addSheet(name: string): Sheet {
    const existing = this.sheet(name);
    if (existing) return existing;
    const sheet = new Sheet(name, this.sheets.length);
    this.sheets.push(sheet);
    return sheet;
  }

  /** Sheet names are case-insensitive in Excel. */
  sheet(name: string): Sheet | undefined {
    const target = name.toUpperCase();
    return this.sheets.find((s) => s.name.toUpperCase() === target);
  }

  sheetIndex(name: string): number {
    const target = name.toUpperCase();
    return this.sheets.findIndex((s) => s.name.toUpperCase() === target);
  }

  /** Sheets between two names inclusive — the span a 3D reference covers. */
  sheetsInSpan(start: string, end: string): Sheet[] {
    const from = this.sheetIndex(start);
    const to = this.sheetIndex(end);
    if (from < 0 || to < 0) return [];
    const [low, high] = from <= to ? [from, to] : [to, from];
    return this.sheets.slice(low, high + 1);
  }

  table(name: string): TableDef | undefined {
    const target = name.toUpperCase();
    return this.tables.find((t) => t.name.toUpperCase() === target);
  }

  /** Defined name lookup: sheet-scoped first, then workbook scope. */
  definedName(name: string, scope: string | null = null): DefinedName | undefined {
    const target = name.toUpperCase();
    if (scope !== null) {
      const scoped = this.names.find(
        (n) =>
          n.name.toUpperCase() === target &&
          n.scope !== null &&
          n.scope.toUpperCase() === scope.toUpperCase()
      );
      if (scoped) return scoped;
    }
    return this.names.find((n) => n.name.toUpperCase() === target && n.scope === null);
  }

  /**
   * Remove a sheet and everything scoped to it. Used by rollback to undo a
   * sheet this agent created — never as an agent-facing operation, because
   * deleting a sheet is not something a change set is allowed to propose.
   */
  removeSheet(name: string): boolean {
    const index = this.sheetIndex(name);
    if (index < 0) return false;
    const target = this.sheets[index]!.name.toUpperCase();
    this.sheets.splice(index, 1);
    this.dropScopedTo(target);
    return true;
  }

  /**
   * Rename a sheet in the model. This does NOT rewrite the formulas that
   * reference it — Excel does that itself on a real rename, and the simulator
   * has no need to, since it never applies renames as a forward operation
   * (see engine.ts). Kept narrow deliberately: a half-implemented rename that
   * silently orphaned references would be worse than none.
   */
  renameSheet(from: string, to: string): boolean {
    const sheet = this.sheet(from);
    if (!sheet || this.sheet(to) !== undefined) return false;
    const previous = sheet.name.toUpperCase();
    sheet.name = to;
    for (const name of this.names) {
      if (name.scope !== null && name.scope.toUpperCase() === previous) name.scope = to;
    }
    for (const table of this.tables) {
      if (table.sheet.toUpperCase() === previous) table.sheet = to;
    }
    return true;
  }

  removeName(name: string, scope: string | null = null): boolean {
    const target = name.toUpperCase();
    const index = this.names.findIndex(
      (candidate) =>
        candidate.name.toUpperCase() === target &&
        (scope === null
          ? candidate.scope === null
          : candidate.scope !== null && candidate.scope.toUpperCase() === scope.toUpperCase())
    );
    if (index < 0) return false;
    this.names.splice(index, 1);
    return true;
  }

  removeTable(name: string): boolean {
    const target = name.toUpperCase();
    const index = this.tables.findIndex((table) => table.name.toUpperCase() === target);
    if (index < 0) return false;
    this.tables.splice(index, 1);
    return true;
  }

  /** Drop names, tables, charts and pivots belonging to a removed sheet. */
  private dropScopedTo(sheetUpper: string): void {
    const drop = <T>(list: T[], belongs: (item: T) => boolean): void => {
      for (let index = list.length - 1; index >= 0; index--) {
        if (belongs(list[index]!)) list.splice(index, 1);
      }
    };
    drop(this.names, (name) => name.scope !== null && name.scope.toUpperCase() === sheetUpper);
    drop(this.tables, (table) => table.sheet.toUpperCase() === sheetUpper);
    drop(this.charts, (chart) => chart.sheet.toUpperCase() === sheetUpper);
    drop(this.pivots, (pivot) => pivot.sheet.toUpperCase() === sheetUpper);
  }

  get cellCount(): number {
    let total = 0;
    for (const sheet of this.sheets) total += sheet.cells.size;
    return total;
  }

  get formulaCount(): number {
    let total = 0;
    for (const sheet of this.sheets) {
      for (const cell of sheet.cells.values()) if (cell.formula !== undefined) total++;
    }
    return total;
  }

  *allCells(): Generator<{ sheet: Sheet; cell: Cell }> {
    for (const sheet of this.sheets) {
      for (const cell of sheet.cells.values()) yield { sheet, cell };
    }
  }
}

// -------------------------------------------------------------- addressing

export function a1(row: number, col: number): string {
  return `${colIndexToLetters(col)}${row + 1}`;
}

export function a1Range(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): string {
  if (startRow === endRow && startCol === endCol) return a1(startRow, startCol);
  return `${a1(startRow, startCol)}:${a1(endRow, endCol)}`;
}

/** "Sheet1!B7" — quoting the sheet name when it needs it. */
export function fullAddress(sheet: string, row: number, col: number): string {
  return `${quoteSheet(sheet)}!${a1(row, col)}`;
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}[0-9]+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

/** Parse "Sheet1!A1:B2" / "A1" into components; null when unparseable. */
export function parseA1Range(text: string): {
  sheet: string | null;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} | null {
  let body = text.trim();
  let sheet: string | null = null;
  const bang = findUnquotedBang(body);
  if (bang >= 0) {
    let raw = body.slice(0, bang);
    if (raw.startsWith("'") && raw.endsWith("'")) raw = raw.slice(1, -1).replace(/''/g, "'");
    sheet = raw;
    body = body.slice(bang + 1);
  }
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)(?::(\$?)([A-Za-z]{1,3})(\$?)([0-9]+))?$/.exec(
    body
  );
  if (!match) return null;
  const startCol = lettersToIndex(match[2]!);
  const startRow = Number(match[4]!) - 1;
  const endCol = match[6] ? lettersToIndex(match[6]) : startCol;
  const endRow = match[8] ? Number(match[8]) - 1 : startRow;
  return {
    sheet,
    startRow: Math.min(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endRow: Math.max(startRow, endRow),
    endCol: Math.max(startCol, endCol),
  };
}

function findUnquotedBang(text: string): number {
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "'") inQuotes = !inQuotes;
    else if (text[i] === "!" && !inQuotes) return i;
  }
  return -1;
}

function lettersToIndex(letters: string): number {
  let value = 0;
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value - 1;
}

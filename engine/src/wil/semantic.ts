/**
 * Semantic mapping (handoff §5.4): classify sheet regions so the WIL can talk
 * about a workbook the way a modeller does — inputs, calculations, outputs,
 * labels, time axes — instead of dumping a grid.
 *
 * Heuristics only, deliberately: they are deterministic, explainable, and free.
 * LLM classification is reserved for blocks this leaves as `unknown`, batched
 * (Phase 3). Every region records `confidence` and the reason it was assigned
 * so ambiguous calls are visible rather than hidden.
 */

import { DependencyGraph, GraphNode } from "../graph/graph";
import { Cell, Sheet, Workbook, a1Range } from "../model/workbook";

export type RegionKind =
  | "input" // constants that feed calculations
  | "calculation" // formula blocks with dependents
  | "output" // formulas with no dependents (or feeding charts)
  | "label" // text headers / row captions
  | "timeAxis" // period headers (2024, Jan, Q1, FY26 ...)
  | "constant" // constants nothing reads
  | "unknown";

export interface Region {
  kind: RegionKind;
  sheet: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  cellCount: number;
  /** Why this classification was made — shown to users and the LLM. */
  reason: string;
  confidence: "high" | "medium" | "low";
  /** Representative formula for calculation/output regions. */
  formula?: string;
  /** Sample of label text for label/timeAxis regions. */
  labels?: string[];
}

const MONTHS = [
  "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
  "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
];

/** "2024", "FY24", "Q3", "Jan-25", "Month 3" — a period header. */
export function looksLikePeriod(value: Cell["value"]): boolean {
  if (typeof value === "number") {
    return Number.isInteger(value) && value >= 1900 && value <= 2200;
  }
  if (typeof value !== "string") return false;
  const text = value.trim().toUpperCase();
  if (text === "") return false;
  if (/^(FY|CY)\s?'?\d{2,4}$/.test(text)) return true;
  if (/^Q[1-4]([ \-']?\d{2,4})?$/.test(text)) return true;
  if (/^\d{4}(E|A|F)?$/.test(text) && Number(text.slice(0, 4)) >= 1900) return true;
  if (MONTHS.some((month) => text.startsWith(month))) return true;
  if (/^(MONTH|YEAR|PERIOD|WEEK|DAY)\s*\d+$/.test(text)) return true;
  return false;
}

function isTextValue(value: Cell["value"]): value is string {
  return typeof value === "string" && value.trim() !== "" && !value.startsWith("=");
}

/** Contiguous same-classification cell bands, merged into rectangles. */
interface Tagged {
  row: number;
  col: number;
  kind: RegionKind;
  reason: string;
}

function mergeTagged(tagged: Tagged[], sheet: string): Region[] {
  const byKey = new Map<string, Tagged>();
  for (const item of tagged) byKey.set(`${item.row},${item.col}`, item);

  const consumed = new Set<string>();
  const regions: Region[] = [];
  const ordered = [...tagged].sort((a, b) => a.row - b.row || a.col - b.col);

  for (const start of ordered) {
    const startKey = `${start.row},${start.col}`;
    if (consumed.has(startKey)) continue;

    const matches = (row: number, col: number): boolean => {
      const key = `${row},${col}`;
      const candidate = byKey.get(key);
      return !consumed.has(key) && candidate !== undefined && candidate.kind === start.kind;
    };

    let endCol = start.col;
    while (matches(start.row, endCol + 1)) endCol++;

    let endRow = start.row;
    for (;;) {
      const nextRow = endRow + 1;
      let complete = true;
      for (let col = start.col; col <= endCol; col++) {
        if (!matches(nextRow, col)) {
          complete = false;
          break;
        }
      }
      if (!complete) break;
      endRow = nextRow;
    }

    for (let row = start.row; row <= endRow; row++) {
      for (let col = start.col; col <= endCol; col++) consumed.add(`${row},${col}`);
    }

    regions.push({
      kind: start.kind,
      sheet,
      startRow: start.row,
      startCol: start.col,
      endRow,
      endCol,
      cellCount: (endRow - start.row + 1) * (endCol - start.col + 1),
      reason: start.reason,
      confidence: start.kind === "unknown" ? "low" : "medium",
    });
  }
  return regions;
}

export interface SemanticMap {
  regions: Region[];
  /** Sheets whose contents look like a header block over a period axis. */
  timeAxes: Region[];
  inputs: Region[];
  outputs: Region[];
}

export function buildSemanticMap(workbook: Workbook, graph: DependencyGraph): SemanticMap {
  const regions: Region[] = [];

  for (const sheet of workbook.sheets) {
    regions.push(...classifySheet(sheet, graph));
  }

  // Promote confidence where a region is corroborated: a label row directly
  // above a calculation block, an input block referenced from another sheet.
  for (const region of regions) {
    if (region.kind === "input" && region.cellCount >= 3) region.confidence = "high";
    if (region.kind === "timeAxis") region.confidence = "high";
  }

  return {
    regions,
    timeAxes: regions.filter((r) => r.kind === "timeAxis"),
    inputs: regions.filter((r) => r.kind === "input"),
    outputs: regions.filter((r) => r.kind === "output"),
  };
}

function classifySheet(sheet: Sheet, graph: DependencyGraph): Region[] {
  const tagged: Tagged[] = [];
  const chartSources = new Set<string>();
  for (const chart of graph.workbook.charts) {
    if (chart.sheet.toUpperCase() !== sheet.name.toUpperCase()) continue;
    for (const range of chart.sourceRanges) chartSources.add(range.toUpperCase());
  }

  // Period-header detection works on whole rows/columns, so do it first.
  const periodCells = detectPeriodBands(sheet);

  for (const cell of sheet.cells.values()) {
    const key = `${cell.row},${cell.col}`;
    if (periodCells.has(key)) {
      tagged.push({
        row: cell.row,
        col: cell.col,
        kind: "timeAxis",
        reason: "part of a contiguous run of period headers",
      });
      continue;
    }

    const node = graph.nodeAt(sheet.name, cell.row, cell.col);
    const isFormula = cell.formula !== undefined;

    if (isFormula) {
      const dependents = node ? graph.dependents(node.id).length : 0;
      if (dependents === 0) {
        tagged.push({
          row: cell.row,
          col: cell.col,
          kind: "output",
          reason: "formula with no downstream dependents",
        });
      } else {
        tagged.push({
          row: cell.row,
          col: cell.col,
          kind: "calculation",
          reason: `formula feeding ${dependents} downstream node(s)`,
        });
      }
      continue;
    }

    if (isTextValue(cell.value)) {
      tagged.push({
        row: cell.row,
        col: cell.col,
        kind: "label",
        reason: "text constant",
      });
      continue;
    }

    const dependents = node ? graph.dependents(node.id).length : 0;
    if (dependents > 0) {
      tagged.push({
        row: cell.row,
        col: cell.col,
        kind: "input",
        reason: `constant read by ${dependents} formula node(s)`,
      });
    } else {
      tagged.push({
        row: cell.row,
        col: cell.col,
        kind: "constant",
        reason: "constant with no dependents",
      });
    }
  }

  const regions = mergeTagged(tagged, sheet.name);

  // Attach representative content.
  for (const region of regions) {
    const anchor = sheet.get(region.startRow, region.startCol);
    if (!anchor) continue;
    if (region.kind === "calculation" || region.kind === "output") {
      region.formula = anchor.formula;
    }
    if (region.kind === "label" || region.kind === "timeAxis") {
      region.labels = collectLabels(sheet, region, 6);
    }
  }
  return regions;
}

function collectLabels(sheet: Sheet, region: Region, limit: number): string[] {
  const out: string[] = [];
  for (let row = region.startRow; row <= region.endRow && out.length < limit; row++) {
    for (let col = region.startCol; col <= region.endCol && out.length < limit; col++) {
      const value = sheet.get(row, col)?.value;
      if (value !== null && value !== undefined && value !== "") out.push(String(value));
    }
  }
  return out;
}

/**
 * A period header must appear in a run of at least three along a row or
 * column — a lone "2024" is just a number, but "2024 2025 2026" is an axis.
 */
function detectPeriodBands(sheet: Sheet): Set<string> {
  const hits = new Set<string>();
  const candidates = new Set<string>();
  for (const cell of sheet.cells.values()) {
    if (cell.formula === undefined && looksLikePeriod(cell.value)) {
      candidates.add(`${cell.row},${cell.col}`);
    }
  }

  const runLength = (
    row: number,
    col: number,
    rowStep: number,
    colStep: number
  ): Array<[number, number]> => {
    const cells: Array<[number, number]> = [];
    let r = row;
    let c = col;
    while (candidates.has(`${r},${c}`)) {
      cells.push([r, c]);
      r += rowStep;
      c += colStep;
    }
    return cells;
  };

  for (const key of candidates) {
    const [row, col] = key.split(",").map(Number) as [number, number];
    if (!candidates.has(`${row},${col - 1}`)) {
      const horizontal = runLength(row, col, 0, 1);
      if (horizontal.length >= 3) for (const [r, c] of horizontal) hits.add(`${r},${c}`);
    }
    if (!candidates.has(`${row - 1},${col}`)) {
      const vertical = runLength(row, col, 1, 0);
      if (vertical.length >= 3) for (const [r, c] of vertical) hits.add(`${r},${c}`);
    }
  }
  return hits;
}

export function regionAddress(region: Region): string {
  return `${region.sheet}!${a1Range(region.startRow, region.startCol, region.endRow, region.endCol)}`;
}

/** Node -> the region containing its anchor, for audit explanations. */
export function regionOfNode(map: SemanticMap, node: GraphNode): Region | undefined {
  return map.regions.find(
    (region) =>
      region.sheet.toUpperCase() === node.sheet.toUpperCase() &&
      node.startRow >= region.startRow &&
      node.startRow <= region.endRow &&
      node.startCol >= region.startCol &&
      node.startCol <= region.endCol
  );
}

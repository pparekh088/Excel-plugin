/**
 * Eval harness (handoff §9).
 *
 * Task classes:
 *  - audit-detection: does the audit engine find the injected defects?
 *    (precision/recall per rule)
 *  - edit / build: does a run of the tool runtime reach the expected
 *    workbook state? Graded programmatically on cell formulas and values.
 *
 * Scored per handoff §9: task success, CELLS DESTROYED (must be 0), plan
 * efficiency (tool calls), cost, latency. Cells-destroyed is computed by
 * diffing a pre-run snapshot against the post-run workbook: any cell whose
 * formula or value was lost without being part of the expected change is
 * destruction, and it is a hard failure regardless of task success.
 */

import { CellValue, Workbook } from "../model/workbook";
import { Simulator } from "../sim/simulator";

export type TaskClass = "audit" | "edit" | "build" | "analyze" | "explain";

export interface CellSnapshot {
  address: string;
  value: CellValue;
  formula?: string;
  numberFormat?: string;
}

/** Full snapshot of a workbook's cells — the basis for destruction accounting. */
export function snapshotWorkbook(workbook: Workbook): Map<string, CellSnapshot> {
  const out = new Map<string, CellSnapshot>();
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) {
      const address = `${sheet.name}!${cell.row},${cell.col}`;
      out.set(address, {
        address,
        value: cell.value,
        ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
        ...(cell.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
      });
    }
  }
  return out;
}

export interface DestructionReport {
  /** Cells that had a formula before and no longer do (or vanished). */
  destroyed: string[];
  /** Cells changed but permitted by the task's expectations. */
  intentional: string[];
}

/**
 * A cell counts as DESTROYED when it previously held a formula and afterwards
 * holds a different formula or none — unless the task declared that address as
 * an expected change. Value-only edits to constants are not destruction.
 */
export function accountDestruction(
  before: Map<string, CellSnapshot>,
  after: Workbook,
  allowed: Set<string>
): DestructionReport {
  const destroyed: string[] = [];
  const intentional: string[] = [];

  for (const [address, snapshot] of before) {
    const [sheetName, coords] = address.split("!") as [string, string];
    const [row, col] = coords.split(",").map(Number) as [number, number];
    const cell = after.sheet(sheetName)?.get(row, col);
    const hadFormula = snapshot.formula !== undefined;
    const hasFormula = cell?.formula !== undefined;

    const lost =
      (hadFormula && !hasFormula) ||
      (hadFormula && hasFormula && cell!.formula !== snapshot.formula) ||
      (!cell && snapshot.value !== null && snapshot.value !== "");

    if (!lost) continue;
    if (allowed.has(address) || allowed.has(normalizeAddress(sheetName, row, col))) {
      intentional.push(address);
    } else {
      destroyed.push(address);
    }
  }
  return { destroyed, intentional };
}

function normalizeAddress(sheet: string, row: number, col: number): string {
  return `${sheet}!${row},${col}`;
}

// ---------------------------------------------------------------- grading

export interface ExpectedCell {
  sheet: string;
  /** A1 address. */
  a1: string;
  /** Exact formula expected (canonical en-US), if the task pins one. */
  formula?: string;
  /** Expected value, compared with tolerance for numbers. */
  value?: CellValue;
  tolerance?: number;
}

export interface GradeResult {
  passed: boolean;
  checks: Array<{ address: string; ok: boolean; detail: string }>;
}

export function gradeCells(workbook: Workbook, expectations: ExpectedCell[]): GradeResult {
  const checks: GradeResult["checks"] = [];
  for (const expectation of expectations) {
    const parsed = /^([A-Za-z]{1,3})([0-9]+)$/.exec(expectation.a1);
    if (!parsed) {
      checks.push({ address: expectation.a1, ok: false, detail: "unparseable address" });
      continue;
    }
    let col = 0;
    for (const ch of parsed[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
    col -= 1;
    const row = Number(parsed[2]) - 1;
    const cell = workbook.sheet(expectation.sheet)?.get(row, col);
    const address = `${expectation.sheet}!${expectation.a1}`;

    if (expectation.formula !== undefined) {
      const actual = cell?.formula ?? "(none)";
      const ok = actual.replace(/\s/g, "") === expectation.formula.replace(/\s/g, "");
      checks.push({
        address,
        ok,
        detail: ok ? "formula matches" : `expected ${expectation.formula}, got ${actual}`,
      });
      if (!ok) continue;
    }
    if (expectation.value !== undefined) {
      const actual = cell?.value ?? null;
      const tolerance = expectation.tolerance ?? 1e-9;
      const ok =
        typeof expectation.value === "number" && typeof actual === "number"
          ? Math.abs(actual - expectation.value) <= tolerance
          : actual === expectation.value;
      checks.push({
        address,
        ok,
        detail: ok ? "value matches" : `expected ${expectation.value}, got ${actual}`,
      });
    }
  }
  return { passed: checks.every((check) => check.ok), checks };
}

// ------------------------------------------------------------ audit scoring

export interface DetectionScore {
  truePositives: number;
  falsePositives: number;
  falseNegatives: number;
  precision: number;
  recall: number;
  f1: number;
}

export function scoreDetection(
  found: Array<{ rule: string; address: string }>,
  expected: Array<{ rule: string; address: string }>
): DetectionScore {
  const key = (item: { rule: string; address: string }) =>
    `${item.rule}@${item.address.toUpperCase()}`;
  const foundKeys = new Set(found.map(key));
  const expectedKeys = new Set(expected.map(key));

  let truePositives = 0;
  for (const item of expectedKeys) if (foundKeys.has(item)) truePositives++;
  const falsePositives = foundKeys.size - truePositives;
  const falseNegatives = expectedKeys.size - truePositives;

  const precision =
    foundKeys.size === 0 ? (expectedKeys.size === 0 ? 1 : 0) : truePositives / foundKeys.size;
  const recall =
    expectedKeys.size === 0 ? 1 : truePositives / expectedKeys.size;
  const f1 =
    precision + recall === 0 ? 0 : (2 * precision * recall) / (precision + recall);

  return { truePositives, falsePositives, falseNegatives, precision, recall, f1 };
}

// ----------------------------------------------------------------- runner

export interface TaskRun {
  taskId: string;
  taskClass: TaskClass;
  model?: string;
  success: boolean;
  cellsDestroyed: number;
  destroyedAddresses: string[];
  toolCalls: number;
  latencyMs: number;
  costUsd: number;
  detail: string;
}

export interface EvalSummary {
  runs: TaskRun[];
  successRate: number;
  totalCellsDestroyed: number;
  meanToolCalls: number;
  meanLatencyMs: number;
  totalCostUsd: number;
}

export function summarize(runs: TaskRun[]): EvalSummary {
  const successes = runs.filter((run) => run.success).length;
  return {
    runs,
    successRate: runs.length === 0 ? 0 : successes / runs.length,
    totalCellsDestroyed: runs.reduce((sum, run) => sum + run.cellsDestroyed, 0),
    meanToolCalls:
      runs.length === 0 ? 0 : runs.reduce((sum, run) => sum + run.toolCalls, 0) / runs.length,
    meanLatencyMs:
      runs.length === 0 ? 0 : runs.reduce((sum, run) => sum + run.latencyMs, 0) / runs.length,
    totalCostUsd: runs.reduce((sum, run) => sum + run.costUsd, 0),
  };
}

/** Recalculate and report whether the simulator could evaluate everything. */
export function recalcAndCheck(workbook: Workbook): {
  ok: boolean;
  unsupported: string[];
  circular: string[];
} {
  const simulator = Simulator.of(workbook);
  const result = simulator.recalculate();
  return {
    ok: result.unsupportedFunctions.length === 0 && result.circular.length === 0,
    unsupported: result.unsupportedFunctions,
    circular: result.circular,
  };
}

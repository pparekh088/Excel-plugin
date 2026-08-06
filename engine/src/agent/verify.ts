/**
 * Verifier (handoff §7): after a change set is applied and recalculated,
 * check that what happened is what was planned, and that nothing broke.
 *
 * Deterministic checks only — the LLM critic is a separate, optional pass in
 * the agent runtime. Verification runs on the blast radius rather than the
 * whole workbook so it stays fast on large models.
 */

import { runAudit } from "../audit/engine";
import { BalanceAssertion } from "../audit/types";
import { DependencyGraph, isErrorValue } from "../graph/graph";
import { Workbook, fullAddress } from "../model/workbook";
import { CellEdit, ChangeSet } from "../changeset/types";
import { isCellEdit } from "../changeset/types";

export interface VerificationIssue {
  kind: "new-error" | "plan-mismatch" | "assertion-failed" | "new-cycle" | "destroyed-formula";
  address?: string;
  detail: string;
}

export interface VerificationResult {
  ok: boolean;
  issues: VerificationIssue[];
  /** Cells in the blast radius that now hold errors but did not before. */
  newErrorCells: string[];
  /** Edits whose result does not match what the plan said it would be. */
  planMismatches: string[];
  checkedCells: number;
}

export interface VerifyOptions {
  /** Error cells that existed BEFORE the change set — not our fault. */
  preExistingErrors?: Set<string>;
  assertions?: BalanceAssertion[];
  /** Cycle count before the change set, to detect newly-introduced cycles. */
  cyclesBefore?: number;
}

/** Snapshot of error cells, used as the "before" baseline. */
export function errorCellSet(workbook: Workbook): Set<string> {
  const out = new Set<string>();
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) {
      if (isErrorValue(cell.value)) out.add(fullAddress(sheet.name, cell.row, cell.col));
    }
  }
  return out;
}

export function verify(
  workbook: Workbook,
  changeSet: ChangeSet,
  options: VerifyOptions = {}
): VerificationResult {
  const issues: VerificationIssue[] = [];
  const graph = DependencyGraph.build(workbook);
  const preExisting = options.preExistingErrors ?? new Set<string>();

  // ---- 1. plan vs actual reconciliation -------------------------------
  // Only the LAST edit targeting a cell describes its intended end state: a
  // repair round legitimately overwrites what the first attempt wrote, and
  // checking the superseded edit would report a permanent false mismatch.
  const finalEdits = new Map<string, CellEdit>();
  for (const edit of changeSet.edits) {
    if (!isCellEdit(edit)) continue;
    finalEdits.set(`${edit.sheet.toUpperCase()}!${edit.row},${edit.col}`, edit);
  }

  const planMismatches: string[] = [];
  for (const edit of finalEdits.values()) {
    const cell = workbook.sheet(edit.sheet)?.get(edit.row, edit.col);
    const address = fullAddress(edit.sheet, edit.row, edit.col);

    if (edit.kind === "clear") {
      if (cell !== undefined && (cell.value !== null || cell.formula !== undefined)) {
        planMismatches.push(address);
        issues.push({
          kind: "plan-mismatch",
          address,
          detail: `${address} was supposed to be cleared but still holds content.`,
        });
      }
      continue;
    }
    if (edit.kind === "setFormula") {
      if (cell?.formula !== edit.formula) {
        planMismatches.push(address);
        issues.push({
          kind: "plan-mismatch",
          address,
          detail:
            `${address} should hold ${edit.formula} but holds ` +
            `${cell?.formula ?? String(cell?.value ?? "(nothing)")}.`,
        });
      }
      continue;
    }
    if (edit.kind === "setValue") {
      if (cell === undefined || !Object.is(cell.value, edit.value)) {
        planMismatches.push(address);
        issues.push({
          kind: "plan-mismatch",
          address,
          detail: `${address} should hold ${String(edit.value)} but holds ${String(cell?.value)}.`,
        });
      }
    }
  }

  // ---- 2. error scan over the blast radius ----------------------------
  const radius = new Set<string>();
  for (const edit of changeSet.edits) {
    if (!isCellEdit(edit)) continue;
    radius.add(fullAddress(edit.sheet, edit.row, edit.col));
    for (const id of graph.impactOfArea(edit.sheet, edit.row, edit.col, edit.row, edit.col)) {
      const node = graph.node(id);
      if (!node) continue;
      for (const ref of graph.nodeCells(node)) {
        radius.add(fullAddress(ref.sheet, ref.row, ref.col));
      }
    }
  }

  const newErrorCells: string[] = [];
  for (const address of radius) {
    const bang = address.lastIndexOf("!");
    const sheetName = address.slice(0, bang).replace(/^'|'$/g, "");
    const cellRef = address.slice(bang + 1);
    const match = /^([A-Z]+)([0-9]+)$/.exec(cellRef);
    if (!match) continue;
    let col = 0;
    for (const ch of match[1]!) col = col * 26 + (ch.charCodeAt(0) - 64);
    col -= 1;
    const row = Number(match[2]) - 1;
    const cell = workbook.sheet(sheetName)?.get(row, col);
    if (cell && isErrorValue(cell.value) && !preExisting.has(address)) {
      newErrorCells.push(address);
      issues.push({
        kind: "new-error",
        address,
        detail: `${address} now evaluates to ${cell.value}, and did not before this change.`,
      });
    }
  }

  // ---- 3. newly introduced cycles -------------------------------------
  if (options.cyclesBefore !== undefined && graph.stats.cycleCount > options.cyclesBefore) {
    issues.push({
      kind: "new-cycle",
      detail:
        `This change introduced ${graph.stats.cycleCount - options.cyclesBefore} new circular ` +
        `reference(s).`,
    });
  }

  // ---- 4. balance assertions ------------------------------------------
  if (options.assertions && options.assertions.length > 0) {
    const report = runAudit(workbook, {
      rules: ["AUD-008"],
      assertions: options.assertions,
      skipAssertionDetection: true,
      graph,
    });
    for (const finding of report.findings) {
      issues.push({
        kind: "assertion-failed",
        address: finding.address,
        detail: finding.explanation,
      });
    }
  }

  // ---- 5. destroyed formulas ------------------------------------------
  // A formula the plan never mentioned must not have disappeared.
  const intended = new Set(
    changeSet.edits
      .filter(isCellEdit)
      .map((edit) => fullAddress(edit.sheet, edit.row, edit.col))
  );
  for (const snap of changeSet.snapshots) {
    if (snap.formula === undefined) continue;
    const address = fullAddress(snap.sheet, snap.row, snap.col);
    if (intended.has(address)) continue;
    const cell = workbook.sheet(snap.sheet)?.get(snap.row, snap.col);
    if (cell?.formula === undefined) {
      issues.push({
        kind: "destroyed-formula",
        address,
        detail: `${address} held a formula that is gone, and no edit targeted it.`,
      });
    }
  }

  return {
    ok: issues.length === 0,
    issues,
    newErrorCells,
    planMismatches,
    checkedCells: radius.size,
  };
}

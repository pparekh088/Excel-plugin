/**
 * The repair approval boundary (INV-2).
 *
 * When verification fails, the agent asks the model for a corrective plan and
 * applies it. That is the one place a write reaches the workbook without the
 * user having seen it, and it is the place an approval boundary matters most:
 * the failure that triggered a repair is by definition a moment when the agent
 * has already got something wrong.
 *
 * What the user approved is a specific set of cells at a specific risk tier.
 * A repair that rewrites one of THOSE cells, at no more risk than they already
 * accepted, is finishing the job they said yes to. A repair that reaches a
 * different cell, adds a sheet, or escalates the risk is a new proposal
 * wearing a repair's clothes, and it goes back to the user.
 *
 * Deterministic and pure: the decision is computed from the two change sets,
 * never asked of the model. A model that could talk its way past this boundary
 * would make the boundary decorative.
 */

import { PriorState, overallRisk, priorStateKey } from "../changeset/engine";
import { ChangeSet, Edit, RiskTier, isCellEdit } from "../changeset/types";
import { Workbook, fullAddress } from "../model/workbook";

const TIER_ORDER: Record<RiskTier, number> = { low: 0, medium: 1, high: 2 };

export type RepairVerdict = "in-scope" | "needs-approval";

export interface RepairScopeDecision {
  verdict: RepairVerdict;
  /** Why, in the words the user should see. */
  reasons: string[];
  /** Cells the repair touches that the approved change set never did. */
  outOfScope: string[];
  /** Risk of the repair edits on their own. */
  repairRisk: RiskTier;
  /** Risk the user already accepted. */
  approvedRisk: RiskTier;
}

function cellKey(sheet: string, row: number, col: number): string {
  return `${sheet.toUpperCase()}!${row},${col}`;
}

/**
 * The cells a change set has the user's consent to write. Derived from the
 * snapshots rather than the edits so a merged change set (original + earlier
 * repairs) reports the union, which is what "already approved" means after a
 * first repair round.
 */
export function approvedCells(changeSet: ChangeSet): Set<string> {
  const cells = new Set<string>();
  for (const snap of changeSet.snapshots) {
    cells.add(cellKey(snap.sheet, snap.row, snap.col));
  }
  return cells;
}

/**
 * Decide whether a repair may proceed on the original approval.
 *
 * In scope requires BOTH:
 *   1. every edit targets a cell the user already approved us writing, and
 *   2. the repair's own risk does not exceed the tier they approved.
 *
 * Structural edits are never in scope. Creating a sheet or defining a name is
 * not a correction to a cell the user looked at, and the preview they approved
 * said nothing about it.
 */
export function classifyRepair(
  workbook: Workbook,
  approved: ChangeSet,
  repairEdits: Edit[]
): RepairScopeDecision {
  const scope = approvedCells(approved);
  const reasons: string[] = [];
  const outOfScope: string[] = [];

  const structural = repairEdits.filter((edit) => !isCellEdit(edit));
  if (structural.length > 0) {
    reasons.push(
      `it makes ${structural.length} structural change(s) (` +
        `${[...new Set(structural.map((edit) => edit.kind))].join(", ")}), which were not part ` +
        `of what you approved`
    );
  }

  for (const edit of repairEdits) {
    if (!isCellEdit(edit)) continue;
    if (!scope.has(cellKey(edit.sheet, edit.row, edit.col))) {
      outOfScope.push(fullAddress(edit.sheet, edit.row, edit.col));
    }
  }
  if (outOfScope.length > 0) {
    reasons.push(
      `it writes to ${outOfScope.length} cell(s) outside the change you approved ` +
        `(${outOfScope.slice(0, 6).join(", ")}${outOfScope.length > 6 ? ", ..." : ""})`
    );
  }

  // Judged against the state the user was shown, not against our own write.
  // Otherwise correcting the formula we just put in a cell reads as
  // "overwriting existing logic" and every repair escalates.
  const prior: PriorState = new Map();
  for (const snap of approved.snapshots) {
    prior.set(priorStateKey(snap.sheet, snap.row, snap.col), snap);
  }
  const repairRisk = overallRisk(repairEdits, workbook, prior);
  const approvedRisk = approved.risk;
  if (TIER_ORDER[repairRisk] > TIER_ORDER[approvedRisk]) {
    reasons.push(
      `it is ${repairRisk} risk, above the ${approvedRisk} risk you approved`
    );
  }

  return {
    verdict: reasons.length === 0 ? "in-scope" : "needs-approval",
    reasons,
    outOfScope,
    repairRisk,
    approvedRisk,
  };
}

/** The text shown when a repair has to go back to the user. */
export function describeRepairEscalation(
  decision: RepairScopeDecision,
  summary: string
): string {
  return [
    `The fix I want to apply goes beyond what you approved, so I am asking first.`,
    ``,
    `What I want to do: ${summary}`,
    ``,
    `Why this needs your approval:`,
    ...decision.reasons.map((reason) => `  - ${reason}`),
    ``,
    `Approving applies this fix. Rejecting leaves the workbook as it is now —`,
    `still failing verification — and I will offer to roll the whole change back.`,
  ].join("\n");
}

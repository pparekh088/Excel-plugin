/**
 * Change-set engine types (INV-2, INV-3, INV-8).
 *
 * Every mutation the agent makes goes through a change set:
 *   propose -> preview (diff + impact) -> approve -> drift check -> apply -> log
 *
 * A change set is never partially applied. Before applying, the prior state of
 * every touched cell is captured (values, formulas, number formats) so
 * rollback can restore it — and rollback reports honestly what it cannot
 * restore, because pivot caches, chart internals and conditional-format stacks
 * are not fully recoverable through the API.
 */

import { CellValue } from "../model/workbook";

export type RiskTier = "low" | "medium" | "high";

export type EditKind =
  | "setValue"
  | "setFormula"
  | "clear"
  | "setNumberFormat"
  | "createSheet"
  | "renameSheet"
  | "createTable"
  | "defineName";

export interface CellEdit {
  kind: "setValue" | "setFormula" | "clear" | "setNumberFormat";
  sheet: string;
  row: number;
  col: number;
  /** New formula (setFormula) — always en-US. */
  formula?: string;
  /** New value (setValue). */
  value?: CellValue;
  /** New number format (setNumberFormat). */
  numberFormat?: string;
}

export interface StructuralEdit {
  kind: "createSheet" | "renameSheet" | "createTable" | "defineName";
  sheet?: string;
  name?: string;
  newName?: string;
  position?: number;
  range?: string;
  refersTo?: string;
  hasHeaders?: boolean;
}

export type Edit = CellEdit | StructuralEdit;

export function isCellEdit(edit: Edit): edit is CellEdit {
  return (
    edit.kind === "setValue" ||
    edit.kind === "setFormula" ||
    edit.kind === "clear" ||
    edit.kind === "setNumberFormat"
  );
}

/** Prior state of one cell — the unit of INV-3 snapshotting. */
export interface CellSnapshot {
  sheet: string;
  row: number;
  col: number;
  value: CellValue;
  formula?: string;
  numberFormat?: string;
  /** The cell did not exist before the change set. */
  absent: boolean;
}

export interface DiffEntry {
  address: string;
  sheet: string;
  row: number;
  col: number;
  before: { value: CellValue; formula?: string; numberFormat?: string; absent: boolean };
  after: { value?: CellValue; formula?: string; numberFormat?: string; cleared: boolean };
  /** Overwriting an existing formula is the highest-risk edit we make. */
  overwritesFormula: boolean;
}

export interface ImpactSummary {
  /** Downstream cells that will recalculate. */
  affectedCells: number;
  /** Distinct downstream run nodes. */
  affectedNodes: number;
  /** Named outputs downstream of the change, for the preview headline. */
  affectedOutputs: string[];
  /** Charts whose source data is touched. */
  affectedCharts: string[];
  /** Pivots whose source data is touched — often not refreshable via API. */
  affectedPivots: string[];
  /**
   * Downstream paths we cannot see because they run through INDIRECT/OFFSET
   * or external links. Non-zero means the impact number is a LOWER BOUND.
   */
  opaqueDownstream: number;
}

export type ChangeSetStatus =
  | "proposed"
  | "previewed"
  | "applied"
  | "rolled-back"
  | "aborted"
  | "failed";

export interface ChangeSet {
  id: string;
  createdAt: string;
  /** What the user asked for, in their words. */
  intent: string;
  /** Plain-language summary of what this change set does. */
  summary: string;
  edits: Edit[];
  risk: RiskTier;
  status: ChangeSetStatus;
  /** State BEFORE we wrote — what rollback restores to. */
  snapshots: CellSnapshot[];
  /**
   * State immediately AFTER we wrote and recalculated. Rollback compares the
   * live cell against this: if it still matches, our write is the most recent
   * one and undoing it is safe. If it differs, a human has edited since, and
   * their work must not be silently reverted.
   */
  appliedState?: CellSnapshot[];
  diff: DiffEntry[];
  impact: ImpactSummary;
  /** Set when the change set was aborted or failed. */
  failureReason?: string;
  appliedAt?: string;
  rolledBackAt?: string;
}

export interface RollbackConflict {
  address: string;
  reason: string;
  /** What we wrote when the change set was applied. */
  applied: { value: CellValue; formula?: string };
  /** What is in the cell now — somebody else's work. */
  current: { value: CellValue; formula?: string };
}

/** What rollback could and could not restore — INV-3 honesty requirement. */
export interface RollbackReport {
  changeSetId: string;
  restoredCells: number;
  /** Things we know we cannot restore, stated plainly. */
  unrestorable: string[];
  /**
   * Cells a human edited AFTER we applied. Rollback leaves these alone: the
   * user's newer work outranks our undo, and silently reverting it would be
   * the single most damaging thing this system could do.
   */
  conflicts: RollbackConflict[];
  ok: boolean;
}

export interface RollbackOptions {
  /**
   * Restore cells that changed after apply anyway. Only ever set from an
   * explicit user decision made with the conflict list in front of them.
   */
  force?: boolean;
}

export interface DriftEntry {
  address: string;
  expected: { value: CellValue; formula?: string };
  actual: { value: CellValue; formula?: string };
}

export interface DriftCheck {
  clean: boolean;
  entries: DriftEntry[];
}

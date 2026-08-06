/**
 * WorkbookHost — the seam between the agent loop and a real spreadsheet.
 *
 * The agent loop is deterministic TypeScript and must stay that way, but every
 * step of it eventually has to touch a workbook: read it, check whether anyone
 * else has edited it, write, recalculate, read back, undo. Doing that directly
 * against the in-memory model means CI proves things about the SIMULATOR, and
 * the simulator is not Excel. Doing it directly against Office.js means the
 * loop cannot run in CI at all.
 *
 * So the loop is written once against this interface, with two implementations:
 *
 *   SimulatorWorkbookHost  in-memory model + our evaluator. Fast, deterministic,
 *                          runs everywhere. What CI exercises.
 *   OfficeJsWorkbookHost   the real thing. Batched reads, calculation
 *                          suspension, Excel's own recalculation. What ships.
 *
 * `kind` is on the interface, and results carry it, because a verification that
 * passed against the simulator is a WEAKER claim than one that passed against
 * Excel. Anywhere we report a result, we can say which host produced it — and
 * the honest reading of a green CI run is "green against our model of Excel".
 *
 * What the simulator provably does NOT reproduce is listed in
 * PLATFORM_QUIRKS.md; the ones that bite are value coercion on write, implicit
 * intersection, table calculated columns, and locale-dependent formula text.
 */

import { ChangeSet, DriftEntry, RollbackOptions, RollbackReport } from "../changeset/types";
import { Workbook } from "../model/workbook";

export type HostKind = "simulator" | "office-js";

export interface ApplyOutcome {
  ok: boolean;
  cellsWritten: number;
  /** Set when the apply was refused or reverted, with the reason. */
  failure?: string;
  /** True when a partial write had to be undone. */
  rolledBack?: boolean;
}

/**
 * What a host can tell us about itself. Used to decide what to attempt and,
 * more importantly, what to CLAIM: a report from a host that cannot observe
 * table calculated columns should not assert that none were affected.
 */
export interface HostCapabilities {
  /** Recalculation is performed by the host, not modelled by us. */
  realRecalculation: boolean;
  /** Writes go through Excel, so coercion and autocorrect apply. */
  coercesOnWrite: boolean;
  /** Structural edits (sheets, names, tables) actually take effect. */
  structuralEdits: boolean;
  /** Sheet protection, merged cells and calculated columns are observable. */
  observesHazards: boolean;
}

export interface WorkbookHost {
  readonly kind: HostKind;
  readonly capabilities: HostCapabilities;

  /**
   * The current state of the workbook as the engine's model.
   *
   * For the simulator this is the model itself. For Office.js this is a
   * chunked extraction, and it is expensive — the loop calls it once at the
   * start and then works from the returned snapshot.
   */
  read(): Promise<Workbook>;

  /** INV-8: has anyone edited the cells this change set is about to write? */
  checkDrift(changeSet: ChangeSet): Promise<DriftEntry[]>;

  /**
   * Write the change set and recalculate. Records `appliedState` on the change
   * set so rollback can later tell our write from a human edit (D-026).
   */
  apply(changeSet: ChangeSet): Promise<ApplyOutcome>;

  /**
   * Re-read after applying, for the verifier.
   *
   * This is the point of the whole abstraction. Against the simulator it
   * returns our own model and verification proves the model self-consistent.
   * Against Office.js it returns what EXCEL holds after Excel recalculated —
   * which is the only thing that can catch a coercion, an implicit
   * intersection, or a calculated column silently rewriting our formula.
   */
  refresh(changeSet: ChangeSet): Promise<Workbook>;

  /** Undo, honouring the post-apply conflict rules (D-026, D-027). */
  rollback(changeSet: ChangeSet, options?: RollbackOptions): Promise<RollbackReport>;
}

/** One line naming what a result was verified against, for reports and logs. */
export function describeHost(host: WorkbookHost): string {
  return host.kind === "office-js"
    ? "verified against the live Excel workbook"
    : "verified against the headless simulator, which is our model of Excel, not Excel";
}

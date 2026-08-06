/**
 * OfficeJsWorkbookHost — the production WorkbookHost (P1-5).
 *
 * The agent loop is written against WorkbookHost so it can run in CI against
 * the simulator and in Excel against this. The important difference is not
 * that this one writes for real; it is that `refresh` RE-READS the workbook
 * after Excel has recalculated, so verification checks what Excel actually
 * holds rather than what we believe we wrote.
 *
 * That distinction is the whole reason the abstraction exists. Excel coerces
 * on write, applies implicit intersection, and lets a table's calculated
 * column rewrite a formula out from under us (Q-010). None of those are
 * modelled by the simulator, so none of them can be caught by a verification
 * that only consults our own model — and they are exactly the failures a user
 * would call "the agent broke my workbook".
 */

import type {
  ApplyOutcome,
  ChangeSet,
  DriftEntry,
  HostCapabilities,
  RollbackOptions,
  RollbackReport,
  Workbook,
  WorkbookHost,
} from "ledger-engine";
import { extractWorkbook, type ExtractOptions } from "./extract";
import { applyChangeSet, checkLiveDrift, rollbackChangeSet } from "./writer";

export interface OfficeJsHostOptions {
  /** Passed to extraction — sheet filters, cell budgets. */
  extract?: ExtractOptions;
  /** Cells per sync when writing. */
  batchSize?: number;
  onProgress?: (progress: { written: number; total: number }) => void;
}

export class OfficeJsWorkbookHost implements WorkbookHost {
  readonly kind = "office-js" as const;

  readonly capabilities: HostCapabilities = {
    realRecalculation: true,
    coercesOnWrite: true,
    structuralEdits: true,
    observesHazards: true,
  };

  constructor(private readonly options: OfficeJsHostOptions = {}) {}

  async read(): Promise<Workbook> {
    return extractWorkbook(this.options.extract);
  }

  async checkDrift(changeSet: ChangeSet): Promise<DriftEntry[]> {
    const live = await checkLiveDrift(changeSet);
    return live.map((entry) => ({
      address: entry.address,
      expected: {
        value: entry.expectedValue as DriftEntry["expected"]["value"],
        ...(entry.expectedFormula !== undefined ? { formula: entry.expectedFormula } : {}),
      },
      actual: {
        value: entry.actualValue as DriftEntry["actual"]["value"],
        ...(entry.actualFormula !== undefined ? { formula: entry.actualFormula } : {}),
      },
    }));
  }

  async apply(changeSet: ChangeSet): Promise<ApplyOutcome> {
    const result = await applyChangeSet(changeSet, {
      ...(this.options.batchSize !== undefined ? { batchSize: this.options.batchSize } : {}),
      ...(this.options.onProgress !== undefined ? { onProgress: this.options.onProgress } : {}),
    });
    return {
      ok: result.ok,
      cellsWritten: result.cellsWritten,
      ...(result.failure !== undefined ? { failure: result.failure } : {}),
      ...(result.rolledBack !== undefined ? { rolledBack: result.rolledBack } : {}),
    };
  }

  /**
   * Re-extract after the write. This is a full extraction, not a targeted
   * re-read of the touched cells, because verification is about the BLAST
   * RADIUS: a formula we wrote can put #REF! into a cell three sheets away,
   * and only a fresh read of the workbook shows that.
   *
   * It is the expensive step of a run. The chunked extractor keeps it within
   * the INV-5 budget, and it happens once per apply plus once per repair.
   */
  async refresh(_changeSet: ChangeSet): Promise<Workbook> {
    return extractWorkbook(this.options.extract);
  }

  async rollback(changeSet: ChangeSet, options: RollbackOptions = {}): Promise<RollbackReport> {
    const report = await rollbackChangeSet(changeSet, options);
    return {
      changeSetId: changeSet.id,
      restoredCells: report.restoredCells,
      reversedStructural: report.reversedStructural,
      unrestorable: report.unrestorable,
      conflicts: report.conflicts,
      ok: report.ok,
    };
  }
}

/**
 * The simulator implementation of WorkbookHost.
 *
 * Everything here operates on the in-memory model with our own evaluator. It
 * is what CI runs, and it is deliberately honest about being a model: reads
 * and writes are exact, nothing is coerced, and `capabilities` says so, so no
 * report built on it can claim to have observed something Excel would have
 * done differently.
 */

import { applyToWorkbook, checkDrift, rollback } from "../changeset/engine";
import {
  ChangeSet,
  DriftEntry,
  RollbackOptions,
  RollbackReport,
  isCellEdit,
} from "../changeset/types";
import { Workbook } from "../model/workbook";
import { Simulator } from "../sim/simulator";
import { ApplyOutcome, HostCapabilities, WorkbookHost } from "./types";

export class SimulatorWorkbookHost implements WorkbookHost {
  readonly kind = "simulator" as const;

  readonly capabilities: HostCapabilities = {
    // We recalculate with our own evaluator — faithful for the functions it
    // implements, absent for the ones it does not.
    realRecalculation: false,
    coercesOnWrite: false,
    // createSheet and defineName take effect; renameSheet and createTable are
    // host-side only (see applyToWorkbook).
    structuralEdits: false,
    observesHazards: false,
  };

  constructor(private readonly workbook: Workbook) {}

  async read(): Promise<Workbook> {
    return this.workbook;
  }

  async checkDrift(changeSet: ChangeSet): Promise<DriftEntry[]> {
    return checkDrift(this.workbook, changeSet).entries;
  }

  async apply(changeSet: ChangeSet): Promise<ApplyOutcome> {
    applyToWorkbook(this.workbook, changeSet);
    Simulator.of(this.workbook).recalculate();
    return {
      ok: true,
      cellsWritten: changeSet.edits.filter(isCellEdit).length,
    };
  }

  async refresh(): Promise<Workbook> {
    // Same object the writes went into: verification against this proves the
    // model self-consistent, not that Excel agrees. describeHost() says so.
    return this.workbook;
  }

  async rollback(changeSet: ChangeSet, options: RollbackOptions = {}): Promise<RollbackReport> {
    const report = rollback(this.workbook, changeSet, options);
    Simulator.of(this.workbook).recalculate();
    return report;
  }
}

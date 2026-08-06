/**
 * Repairs against a live-like host (R2 item 1).
 *
 * On the simulator, host.read(), host.refresh() and the runtime's workbook are
 * the same object, so planning a repair against "the workbook" accidentally
 * works. On Office.js they are different extractions, and the original code
 * planned repairs against the PRE-APPLY extraction: the repair's snapshots
 * claimed the original cell state, the live writer's own drift check then saw
 * "original state expected, first AI write present" and refused — and the
 * runtime merged the refused repair into the master change set anyway.
 *
 * The host here behaves like the real one in the two ways that matter: reads
 * hand out clones (a re-extraction, not a reference), and apply refuses when a
 * snapshot does not match what the workbook holds now.
 */

import { describe, expect, it } from "vitest";
import { MockLlmProvider } from "../../src/agent/llm";
import { INTENT_MARKER, createDefaultExecutor, runAgent } from "../../src/agent/runtime";
import { applyToWorkbook } from "../../src/changeset/engine";
import { ApplyOutcome, HostCapabilities, WorkbookHost } from "../../src/host/types";
import {
  ChangeSet,
  DriftEntry,
  RollbackOptions,
  RollbackReport,
  isCellEdit,
} from "../../src/changeset/types";
import { rollback } from "../../src/changeset/engine";
import { Simulator } from "../../src/sim/simulator";
import { Workbook } from "../../src/model/workbook";
import { workbookOf } from "../helpers/build";

function cloneWorkbook(source: Workbook): Workbook {
  const copy = new Workbook(source.name);
  for (const sheet of source.sheets) {
    const target = copy.addSheet(sheet.name);
    for (const cell of sheet.cells.values()) {
      target.set({ ...cell });
    }
  }
  copy.names.push(...source.names.map((name) => ({ ...name })));
  return copy;
}

/**
 * A host that mimics the live writer's contract: reads are re-extractions
 * (clones), and apply performs its own drift check against the change set's
 * snapshots before writing — exactly what checkLiveDrift does.
 */
class LiveLikeHost implements WorkbookHost {
  readonly kind = "office-js" as const;
  readonly capabilities: HostCapabilities = {
    realRecalculation: true,
    coercesOnWrite: true,
    structuralEdits: true,
    observesHazards: true,
  };
  applyCalls = 0;
  refusedApplies = 0;

  constructor(private readonly workbook: Workbook) {}

  async read(): Promise<Workbook> {
    return cloneWorkbook(this.workbook);
  }

  async checkDrift(changeSet: ChangeSet): Promise<DriftEntry[]> {
    return this.driftAgainstSnapshots(changeSet);
  }

  private driftAgainstSnapshots(changeSet: ChangeSet): DriftEntry[] {
    const entries: DriftEntry[] = [];
    for (const snap of changeSet.snapshots) {
      const cell = this.workbook.sheet(snap.sheet)?.get(snap.row, snap.col);
      const sameFormula = cell?.formula === snap.formula;
      const valueMatters = snap.formula === undefined;
      const sameValue = !valueMatters || Object.is(cell?.value ?? null, snap.value);
      if (!sameFormula || !sameValue) {
        entries.push({
          address: `${snap.sheet}!${snap.row},${snap.col}`,
          expected: { value: snap.value, ...(snap.formula ? { formula: snap.formula } : {}) },
          actual: {
            value: cell?.value ?? null,
            ...(cell?.formula ? { formula: cell.formula } : {}),
          },
        });
      }
    }
    return entries;
  }

  async apply(changeSet: ChangeSet): Promise<ApplyOutcome> {
    this.applyCalls++;
    // The live writer refuses to write when the workbook no longer matches
    // the snapshots the change set was proposed against.
    const drift = this.driftAgainstSnapshots(changeSet);
    if (drift.length > 0) {
      this.refusedApplies++;
      return {
        ok: false,
        cellsWritten: 0,
        failure: `Aborted: ${drift.length} cell(s) changed since this plan was made.`,
      };
    }
    applyToWorkbook(this.workbook, changeSet);
    Simulator.of(this.workbook).recalculate();
    return { ok: true, cellsWritten: changeSet.edits.filter(isCellEdit).length };
  }

  async refresh(): Promise<Workbook> {
    return cloneWorkbook(this.workbook);
  }

  async rollback(changeSet: ChangeSet, options: RollbackOptions = {}): Promise<RollbackReport> {
    const report = rollback(this.workbook, changeSet, options);
    Simulator.of(this.workbook).recalculate();
    return report;
  }

  current(): Workbook {
    return this.workbook;
  }
}

function planJson(steps: unknown[], summary = "do the thing"): string {
  return JSON.stringify({ summary, steps });
}

describe("repairs against a live-like host", () => {
  it("a same-cell repair passes the live drift check and converges", async () => {
    const workbook = workbookOf({
      S: { A1: 10, A2: 20, A3: 30, B1: "=A1*2", B2: "=A2*2", B3: "=A3*2" },
    });
    Simulator.of(workbook).recalculate();
    const host = new LiveLikeHost(workbook);

    const provider = new MockLlmProvider()
      .script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/Z9" } }])
      )
      .script(
        "verification failed",
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } }])
      );

    const result = await runAgent(host, {
      intent: "add a ratio",
      provider,
      approver: async () => "approve",
      executor: createDefaultExecutor(),
    });

    // The repair snapshots against the post-first-apply state, so the host's
    // drift check accepts it. Against the pre-apply state it would refuse.
    expect(host.refusedApplies).toBe(0);
    expect(result.outcome).toBe("applied");
    expect(result.repairAttempts).toBe(1);
    expect(host.current().sheet("S")!.get(0, 2)!.formula).toBe("=B1/2");
  });

  it("the merged change set's repair snapshot records the post-apply state", async () => {
    const workbook = workbookOf({ S: { A1: 10, B1: "=A1*2" } });
    Simulator.of(workbook).recalculate();
    const host = new LiveLikeHost(workbook);

    const provider = new MockLlmProvider()
      .script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/Z9" } }])
      )
      .script(
        "verification failed",
        // The repair touches a NEW cell (D1) as well as fixing C1 — so it
        // needs approval, and its snapshot of D1 is D1's true before-state.
        planJson([
          { tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } },
          { tool: "formula.set", params: { sheet: "S", a1: "D1", formula: "=C1+1" } },
        ])
      );

    const result = await runAgent(host, {
      intent: "add a ratio",
      provider,
      approver: async () => "approve",
      executor: createDefaultExecutor(),
    });

    expect(result.outcome).toBe("applied");
    const snapshots = result.changeSet!.snapshots;
    // C1's snapshot is the ORIGINAL before-state (absent), captured by the
    // first change set; D1's is absent too, captured by the repair.
    const c1 = snapshots.find((snap) => snap.row === 0 && snap.col === 2)!;
    const d1 = snapshots.find((snap) => snap.row === 0 && snap.col === 3)!;
    expect(c1.absent).toBe(true);
    expect(d1.absent).toBe(true);

    // Rollback of the merged set through the host removes both cells.
    const report = await host.rollback(result.changeSet!);
    expect(report.restoredCells).toBe(2);
    expect(host.current().sheet("S")!.get(0, 2)).toBeUndefined();
    expect(host.current().sheet("S")!.get(0, 3)).toBeUndefined();
  });

  it("a repair the host refuses is NOT merged into the change set", async () => {
    const workbook = workbookOf({ S: { A1: 10, B1: "=A1*2" } });
    Simulator.of(workbook).recalculate();
    const host = new LiveLikeHost(workbook);

    // Between the first apply and the repair, a human edits the repair's
    // target cell, so the host refuses the repair's write.
    const originalApply = host.apply.bind(host);
    let applied = 0;
    host.apply = async (changeSet: ChangeSet) => {
      applied++;
      if (applied === 2) {
        host.current().sheet("S")!.set({ row: 0, col: 2, value: 777 });
      }
      return originalApply(changeSet);
    };

    const provider = new MockLlmProvider()
      .script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/Z9" } }])
      )
      .script(
        "verification failed",
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } }])
      );

    const result = await runAgent(host, {
      intent: "add a ratio",
      provider,
      approver: async (request) => (request.text.includes("Roll back") ? "reject" : "approve"),
      executor: createDefaultExecutor(),
      maxRepairs: 1,
    });

    // The refused repair's edits must not appear in the audit record.
    expect(result.changeSet!.edits).toHaveLength(1);
    expect(result.transcript.join(" ")).toContain("could not be applied");
    expect(result.outcome).not.toBe("applied");
  });

  it("asks for approval with the repair's own change set, not the original", async () => {
    const workbook = workbookOf({ S: { A1: 10, B1: "=A1*2" } });
    Simulator.of(workbook).recalculate();
    const host = new LiveLikeHost(workbook);

    const provider = new MockLlmProvider()
      .script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/Z9" } }])
      )
      .script(
        "verification failed",
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "F6", formula: "=1" } }])
      );

    const approvals: ChangeSet[] = [];
    await runAgent(host, {
      intent: "add a ratio",
      provider,
      approver: async (request) => {
        if (request.changeSet) approvals.push(request.changeSet);
        return request.text.includes("Roll back") ? "reject" : "approve";
      },
      executor: createDefaultExecutor(),
      maxRepairs: 1,
    });

    // The escalated repair approval must carry the REPAIR's change set: one
    // edit targeting F6, with its own diff — not the original change set.
    const repairApproval = approvals.find((cs) => cs.intent.startsWith("repair:"));
    expect(repairApproval).toBeDefined();
    expect(repairApproval!.edits).toHaveLength(1);
    expect(repairApproval!.diff[0]!.address).toBe("S!F6");
  });
});

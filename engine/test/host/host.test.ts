/**
 * The WorkbookHost contract (P1-5).
 *
 * The agent loop runs against an interface so CI can exercise it against the
 * simulator and production can run the same code against Excel. That is only
 * worth anything if the contract is pinned: these tests state what a host must
 * do, and the simulator implementation has to satisfy it.
 *
 * The Office.js implementation is checked against the same expectations in
 * addin/test/host.test.ts, using the Office.js fake. Neither is a substitute
 * for the live host — sideload item 13 is.
 */

import { describe, expect, it } from "vitest";
import { MockLlmProvider } from "../../src/agent/llm";
import { INTENT_MARKER, createDefaultExecutor, runAgent } from "../../src/agent/runtime";
import { proposeChangeSet } from "../../src/changeset/engine";
import { SimulatorWorkbookHost } from "../../src/host/simulator";
import { describeHost } from "../../src/host/types";
import { Simulator } from "../../src/sim/simulator";
import { workbookOf } from "../helpers/build";

const book = () => workbookOf({ S: { A1: 10, B1: "=A1*2" } });

function planJson(steps: unknown[], summary = "do the thing"): string {
  return JSON.stringify({ summary, steps });
}

describe("SimulatorWorkbookHost satisfies the contract", () => {
  it("reads the workbook it was given", async () => {
    const workbook = book();
    expect(await new SimulatorWorkbookHost(workbook).read()).toBe(workbook);
  });

  it("reports no drift on an untouched workbook", async () => {
    const workbook = book();
    const host = new SimulatorWorkbookHost(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    expect(await host.checkDrift(changeSet)).toHaveLength(0);
  });

  it("reports drift after a co-author edit", async () => {
    const workbook = book();
    const host = new SimulatorWorkbookHost(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 99, formula: "=A1*9" });
    expect(await host.checkDrift(changeSet)).toHaveLength(1);
  });

  it("applies, recalculates, and records post-apply state", async () => {
    const workbook = book();
    Simulator.of(workbook).recalculate();
    const host = new SimulatorWorkbookHost(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });

    const outcome = await host.apply(changeSet);
    expect(outcome.ok).toBe(true);
    expect(outcome.cellsWritten).toBe(1);
    // Recalculated: 10 * 3.
    expect(workbook.sheet("S")!.get(0, 1)!.value).toBe(30);
    // And the evidence rollback needs (D-026) is there.
    expect(changeSet.appliedState).toHaveLength(1);
  });

  it("refreshes to the state the writes went into", async () => {
    const workbook = book();
    const host = new SimulatorWorkbookHost(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 5, col: 5, value: 1 }],
    });
    await host.apply(changeSet);
    const refreshed = await host.refresh(changeSet);
    expect(refreshed.sheet("S")!.get(5, 5)!.value).toBe(1);
  });

  it("rolls back through the host, honouring conflicts", async () => {
    const workbook = book();
    Simulator.of(workbook).recalculate();
    const host = new SimulatorWorkbookHost(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 20 }],
    });
    await host.apply(changeSet);
    workbook.sheet("S")!.set({ row: 0, col: 0, value: 30 });

    const report = await host.rollback(changeSet);
    expect(report.conflicts).toHaveLength(1);
    expect(workbook.sheet("S")!.get(0, 0)!.value).toBe(30);
  });

  it("does not overstate what it can observe", () => {
    const host = new SimulatorWorkbookHost(book());
    // Every one of these is false, and saying so is the point: a report built
    // on this host must not claim Excel-level fidelity.
    expect(host.capabilities.realRecalculation).toBe(false);
    expect(host.capabilities.coercesOnWrite).toBe(false);
    expect(host.capabilities.structuralEdits).toBe(false);
    expect(host.capabilities.observesHazards).toBe(false);
  });
});

describe("describeHost states what a pass is worth", () => {
  it("is explicit that the simulator is not Excel", () => {
    const text = describeHost(new SimulatorWorkbookHost(book()));
    expect(text).toContain("not Excel");
  });

  it("says so plainly for a real host", () => {
    const fake = { kind: "office-js" as const };
    expect(describeHost(fake as never)).toContain("live Excel");
  });
});

describe("runAgent through the host seam", () => {
  it("accepts a bare Workbook and runs on the simulator", async () => {
    const workbook = book();
    Simulator.of(workbook).recalculate();
    const result = await runAgent(workbook, {
      intent: "double it",
      provider: new MockLlmProvider().script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=A1*4" } }])
      ),
      approver: async () => "approve",
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("applied");
    expect(result.host).toBe("simulator");
  });

  it("accepts an explicit host and reports which one ran", async () => {
    const workbook = book();
    Simulator.of(workbook).recalculate();
    const result = await runAgent(new SimulatorWorkbookHost(workbook), {
      intent: "double it",
      provider: new MockLlmProvider().script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=A1*4" } }])
      ),
      approver: async () => "approve",
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("applied");
    expect(result.host).toBe("simulator");
    expect(result.transcript.join(" ")).toContain("not Excel");
  });

  it("verifies against what the host reports, not against the plan", async () => {
    // A host whose refresh() hands back a workbook where our write did NOT
    // land is exactly the production failure mode (coercion, calculated
    // column). Verification must catch it rather than trust the change set.
    const workbook = book();
    Simulator.of(workbook).recalculate();
    const host = new SimulatorWorkbookHost(workbook);
    const lying = {
      ...host,
      kind: host.kind,
      capabilities: host.capabilities,
      read: () => host.read(),
      checkDrift: (cs: Parameters<typeof host.checkDrift>[0]) => host.checkDrift(cs),
      apply: async (cs: Parameters<typeof host.apply>[0]) => {
        const outcome = await host.apply(cs);
        // Excel "rewrites" our formula after the write.
        workbook.sheet("S")!.set({ row: 0, col: 2, value: 0, formula: "=1/0" });
        return outcome;
      },
      refresh: (cs: Parameters<typeof host.refresh>[0]) => host.refresh(cs),
      rollback: (cs: Parameters<typeof host.rollback>[0]) => host.rollback(cs),
    };

    const result = await runAgent(lying, {
      intent: "add a ratio",
      provider: new MockLlmProvider().script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=A1*4" } }])
      ),
      approver: async (request) => (request.text.includes("Roll back") ? "reject" : "approve"),
      executor: createDefaultExecutor(),
      maxRepairs: 0,
    });

    // The plan said =A1*4; the host reports =1/0. Verification must not pass.
    expect(result.verification!.ok).toBe(false);
  });

  it("stops and reports when the host refuses the write", async () => {
    const workbook = book();
    Simulator.of(workbook).recalculate();
    const host = new SimulatorWorkbookHost(workbook);
    const refusing = {
      ...host,
      kind: host.kind,
      capabilities: host.capabilities,
      read: () => host.read(),
      checkDrift: (cs: Parameters<typeof host.checkDrift>[0]) => host.checkDrift(cs),
      apply: async () => ({
        ok: false,
        cellsWritten: 0,
        failure: "Sheet is protected.",
        rolledBack: false,
      }),
      refresh: (cs: Parameters<typeof host.refresh>[0]) => host.refresh(cs),
      rollback: (cs: Parameters<typeof host.rollback>[0]) => host.rollback(cs),
    };

    const result = await runAgent(refusing, {
      intent: "write something",
      provider: new MockLlmProvider().script(
        INTENT_MARKER,
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=A1*4" } }])
      ),
      approver: async () => "approve",
      executor: createDefaultExecutor(),
    });

    expect(result.outcome).toBe("write-failed");
    expect(result.changeSet!.status).toBe("failed");
    expect(result.transcript.join(" ")).toContain("Sheet is protected");
  });
});

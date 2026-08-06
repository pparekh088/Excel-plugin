/**
 * The repair approval boundary (P0-4).
 *
 * The repair loop was the one place a write reached the workbook without the
 * user seeing it: verification fails, the model proposes a fix, the fix is
 * applied. Since a repair only happens when the agent has ALREADY got
 * something wrong, "trust it to fix itself, anywhere, at any risk" is the
 * wrong default.
 *
 * The boundary: a fix confined to cells the user approved us writing, at no
 * more risk than they accepted, is finishing the job they said yes to.
 * Anything wider goes back to them.
 */

import { describe, expect, it } from "vitest";
import { MockLlmProvider } from "../../src/agent/llm";
import { classifyRepair, approvedCells } from "../../src/agent/repairScope";
import { createDefaultExecutor, runAgent } from "../../src/agent/runtime";
import { proposeChangeSet } from "../../src/changeset/engine";
import { Edit } from "../../src/changeset/types";
import { Simulator } from "../../src/sim/simulator";
import { workbookOf } from "../helpers/build";

const workbook = () =>
  workbookOf({ S: { A1: 10, A2: 20, A3: 30, B1: "=A1*2", B2: "=A2*2", B3: "=A3*2" } });

function planJson(steps: unknown[], summary = "do the thing"): string {
  return JSON.stringify({ summary, steps });
}

describe("classifyRepair", () => {
  it("allows a repair that rewrites a cell the user approved", () => {
    const book = workbook();
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 2, formula: "=B1/0" }],
    });
    const decision = classifyRepair(book, approved, [
      { kind: "setFormula", sheet: "S", row: 0, col: 2, formula: "=B1/2" },
    ]);
    expect(decision.verdict).toBe("in-scope");
    expect(decision.reasons).toHaveLength(0);
  });

  it("escalates a repair that reaches a cell outside the approval", () => {
    const book = workbook();
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 2, formula: "=B1/0" }],
    });
    const decision = classifyRepair(book, approved, [
      { kind: "setFormula", sheet: "S", row: 0, col: 2, formula: "=B1/2" },
      { kind: "setValue", sheet: "S", row: 5, col: 5, value: 1 },
    ]);
    expect(decision.verdict).toBe("needs-approval");
    expect(decision.outOfScope).toEqual(["S!F6"]);
    expect(decision.reasons.join(" ")).toContain("outside the change you approved");
  });

  it("escalates a repair that touches a different sheet", () => {
    const book = workbookOf({ S: { A1: 1 }, Other: { A1: 1 } });
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 }],
    });
    const decision = classifyRepair(book, approved, [
      { kind: "setValue", sheet: "Other", row: 0, col: 0, value: 2 },
    ]);
    expect(decision.verdict).toBe("needs-approval");
    expect(decision.outOfScope).toEqual(["Other!A1"]);
  });

  it("escalates a repair that raises the risk tier", () => {
    // Approved: writing into an EMPTY cell — medium risk.
    const book = workbook();
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 9, col: 9, value: 1 }],
    });
    expect(approved.risk).toBe("medium");

    // The "repair" clears that cell instead, which is high risk.
    const decision = classifyRepair(book, approved, [
      { kind: "clear", sheet: "S", row: 9, col: 9 },
    ]);
    expect(decision.verdict).toBe("needs-approval");
    expect(decision.repairRisk).toBe("high");
    expect(decision.reasons.join(" ")).toContain("above the medium risk you approved");
  });

  it("never lets a structural edit in on a cell approval", () => {
    const book = workbook();
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 1 }],
    });
    const decision = classifyRepair(book, approved, [
      { kind: "createSheet", name: "Helper" },
    ]);
    expect(decision.verdict).toBe("needs-approval");
    expect(decision.reasons.join(" ")).toContain("structural change");
  });

  it("counts cells approved by an earlier repair round as in scope", () => {
    // After a merge, the change set's snapshots are the union — which is what
    // "already approved" has to mean once a repair has been approved.
    const book = workbook();
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [
        { kind: "setFormula", sheet: "S", row: 0, col: 2, formula: "=B1" },
        { kind: "setFormula", sheet: "S", row: 1, col: 2, formula: "=B2" },
      ],
    });
    expect(approvedCells(approved)).toEqual(new Set(["S!0,2", "S!1,2"]));
    const decision = classifyRepair(book, approved, [
      { kind: "setFormula", sheet: "S", row: 1, col: 2, formula: "=B2*1" },
    ]);
    expect(decision.verdict).toBe("in-scope");
  });

  it("treats an empty repair as in scope", () => {
    const book = workbook();
    const approved = proposeChangeSet(book, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 1 }],
    });
    expect(classifyRepair(book, approved, [] as Edit[]).verdict).toBe("in-scope");
  });
});

describe("the runtime honours the boundary", () => {
  /** A first plan that fails verification, then a repair the test controls. */
  function providerWith(repair: unknown[]) {
    return new MockLlmProvider()
      .script(
        "The user's request",
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/Z9" } }])
      )
      .script("verification failed", planJson(repair));
  }

  it("applies an in-scope repair without asking", async () => {
    const book = workbook();
    Simulator.of(book).recalculate();
    const asked: string[] = [];

    const result = await runAgent(book, {
      intent: "add a ratio",
      provider: providerWith([
        { tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } },
      ]),
      approver: async (request) => {
        asked.push(request.text);
        return "approve";
      },
      executor: createDefaultExecutor(),
    });

    expect(result.outcome).toBe("applied");
    // One approval: the original plan. The repair did not need a second.
    expect(asked.filter((text) => text.includes("fix"))).toHaveLength(0);
    expect(result.transcript.join(" ")).toContain("within the change you already approved");
  });

  it("asks before a repair that reaches outside the approved cells", async () => {
    const book = workbook();
    Simulator.of(book).recalculate();
    const asked: string[] = [];

    await runAgent(book, {
      intent: "add a ratio",
      provider: providerWith([
        { tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } },
        { tool: "formula.set", params: { sheet: "S", a1: "F6", formula: "=1" } },
      ]),
      approver: async (request) => {
        asked.push(request.text);
        return "approve";
      },
      executor: createDefaultExecutor(),
    });

    const escalation = asked.find((text) => text.includes("goes beyond what you approved"));
    expect(escalation).toBeDefined();
    expect(escalation).toContain("S!F6");
  });

  it("does not apply the repair when the user rejects it", async () => {
    const book = workbook();
    Simulator.of(book).recalculate();

    const result = await runAgent(book, {
      intent: "add a ratio",
      provider: providerWith([
        { tool: "formula.set", params: { sheet: "S", a1: "F6", formula: "=1" } },
      ]),
      // Approve the plan, reject the repair, then reject the rollback offer so
      // the assertion is about the repair alone.
      approver: async (request) =>
        request.text.includes("goes beyond") || request.text.includes("Roll back")
          ? "reject"
          : "approve",
      executor: createDefaultExecutor(),
    });

    expect(book.sheet("S")!.get(5, 5)).toBeUndefined();
    expect(result.transcript.join(" ")).toContain("Fix rejected");
  });

  it("asks for every repair under repairApproval: always", async () => {
    const book = workbook();
    Simulator.of(book).recalculate();
    const asked: string[] = [];

    const result = await runAgent(book, {
      intent: "add a ratio",
      provider: providerWith([
        { tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } },
      ]),
      approver: async (request) => {
        asked.push(request.text);
        return "approve";
      },
      executor: createDefaultExecutor(),
      repairApproval: "always",
    });

    expect(result.outcome).toBe("applied");
    expect(asked.some((text) => text.includes("Apply this fix?"))).toBe(true);
  });

  it("still converges on the ordinary same-cell repair (no approval churn)", async () => {
    const book = workbook();
    Simulator.of(book).recalculate();
    let approvals = 0;

    const result = await runAgent(book, {
      intent: "add a ratio",
      provider: providerWith([
        { tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } },
      ]),
      approver: async () => {
        approvals++;
        return "approve";
      },
      executor: createDefaultExecutor(),
    });

    expect(result.outcome).toBe("applied");
    // The two the design always asks for — the plan and the change-set
    // preview — and nothing extra for the repair.
    expect(approvals).toBe(2);
    expect(book.sheet("S")!.get(0, 2)!.formula).toBe("=B1/2");
  });
});

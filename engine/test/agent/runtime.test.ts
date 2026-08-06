import { describe, expect, it, vi } from "vitest";
import { MockLlmProvider, CostMeter, tierFor } from "../../src/agent/llm";
import { isAutoApprovable, parsePlan, renderPlan } from "../../src/agent/plan";
import {
  createDefaultExecutor,
  runAgent,
  translateFormula,
} from "../../src/agent/runtime";
import { TOOLS, renderToolCatalogue, toolByName } from "../../src/agent/tools";
import { Simulator } from "../../src/sim/simulator";
import { workbookOf } from "../helpers/build";
import { threeStatementModel } from "../../src/corpus";

const approveAll = async () => "approve" as const;
const rejectAll = async () => "reject" as const;

function planJson(steps: unknown[], summary = "do the thing"): string {
  return JSON.stringify({ summary, steps });
}

describe("tool catalogue", () => {
  it("covers the handoff's ~60 tools across all three categories", () => {
    expect(TOOLS.length).toBeGreaterThanOrEqual(60);
    expect(TOOLS.filter((t) => t.category === "inspection").length).toBeGreaterThan(10);
    expect(TOOLS.filter((t) => t.category === "mutation").length).toBeGreaterThan(15);
    expect(TOOLS.filter((t) => t.category === "control").length).toBeGreaterThan(5);
  });

  it("gives every mutation tool a risk tier", () => {
    for (const tool of TOOLS.filter((t) => t.category === "mutation")) {
      expect(["low", "medium", "high"], tool.name).toContain(tool.risk);
    }
  });

  it("marks every inspection tool as no-risk and read access", () => {
    for (const tool of TOOLS.filter((t) => t.category === "inspection")) {
      expect(tool.risk, tool.name).toBe("none");
      expect(tool.access, tool.name).toBe("read");
    }
  });

  it("documents the known API gaps so the planner does not attempt them", () => {
    const catalogue = renderToolCatalogue();
    expect(catalogue).toContain("OLAP");
    expect(catalogue).toContain("VBA");
    expect(catalogue).toContain("Data Tables");
    expect(catalogue).toContain("external links");
  });

  it("names deletion and overwrite tools as HIGH risk", () => {
    for (const name of ["range.clear", "row.delete", "column.delete", "sheet.delete", "name.delete"]) {
      expect(toolByName(name)?.risk, name).toBe("high");
    }
  });
});

describe("plan parsing (INV-1)", () => {
  it("accepts a well-formed plan", () => {
    const result = parsePlan(
      planJson([{ tool: "formula.set", params: { sheet: "S", a1: "B1", formula: "=A1*2" }, rationale: "why" }])
    );
    expect(result.errors).toEqual([]);
    expect(result.plan!.steps).toHaveLength(1);
    expect(result.plan!.risk).toBe("high");
  });

  it("rejects an unknown tool rather than passing it through", () => {
    const result = parsePlan(planJson([{ tool: "workbook.exec", params: {} }]));
    expect(result.plan).toBeUndefined();
    expect(result.errors[0]!.message).toContain("Unknown tool");
  });

  it("rejects a step with no tool name", () => {
    expect(parsePlan(planJson([{ params: {} }])).plan).toBeUndefined();
  });

  it("rejects non-object params", () => {
    expect(parsePlan(planJson([{ tool: "range.read", params: "everything" }])).plan).toBeUndefined();
  });

  it("rejects a response with no JSON", () => {
    expect(parsePlan("I would rather not.").plan).toBeUndefined();
  });

  it("rejects an empty plan", () => {
    expect(parsePlan(planJson([])).plan).toBeUndefined();
  });

  it("extracts JSON from a fenced block with prose around it", () => {
    const text = "Here is my plan:\n```json\n" + planJson([{ tool: "audit.run", params: {} }]) + "\n```\nHope that helps.";
    expect(parsePlan(text).plan?.steps).toHaveLength(1);
  });

  it("takes the highest risk across steps", () => {
    const result = parsePlan(
      planJson([
        { tool: "number_format.set", params: {} },
        { tool: "range.clear", params: {} },
      ])
    );
    expect(result.plan!.risk).toBe("high");
  });

  it("only auto-approves plans that are entirely low risk", () => {
    const low = parsePlan(planJson([{ tool: "number_format.set", params: {} }])).plan!;
    const high = parsePlan(planJson([{ tool: "range.clear", params: {} }])).plan!;
    expect(isAutoApprovable(low)).toBe(true);
    expect(isAutoApprovable(high)).toBe(false);
  });

  it("renders a plan for human review", () => {
    const plan = parsePlan(
      planJson([{ tool: "formula.set", params: { a1: "B1" }, rationale: "recompute margin" }])
    ).plan!;
    const text = renderPlan(plan);
    expect(text).toContain("formula.set");
    expect(text).toContain("recompute margin");
    expect(text).toContain("Risk:");
  });
});

describe("formula translation for fills", () => {
  it.each([
    ["=A1*2", 1, 0, "=A2*2"],
    ["=A1*2", 0, 1, "=B1*2"],
    ["=$A$1*B1", 1, 0, "=$A$1*B2"],
    ["=A$1*$B1", 2, 1, "=B$1*$B3"],
    ["=SUM(A1:A10)", 1, 0, "=SUM(A2:A11)"],
    ["=Sheet2!B5", 3, 0, "=Sheet2!B8"],
  ])("translates %s by (%i,%i)", (formula, dr, dc, expected) => {
    expect(translateFormula(formula, dr, dc)).toBe(expected);
  });
});

describe("model routing (handoff §7)", () => {
  it("routes planner and critic to the strong tier, executor to fast", () => {
    expect(tierFor("planner")).toBe("strong");
    expect(tierFor("critic")).toBe("strong");
    expect(tierFor("executor")).toBe("fast");
    expect(tierFor("classifier")).toBe("cheap");
  });

  it("meters cost per tier", () => {
    const meter = new CostMeter();
    meter.record("strong", { inputTokens: 100, outputTokens: 50, costUsd: 0.01, model: "m" });
    meter.record("fast", { inputTokens: 40, outputTokens: 10, costUsd: 0.001, model: "m" });
    expect(meter.totalCalls).toBe(2);
    expect(meter.totalCostUsd).toBeCloseTo(0.011);
    expect(meter.byTier().strong.calls).toBe(1);
  });
});

describe("agent runtime loop", () => {
  const baseWorkbook = () =>
    workbookOf({
      S: { A1: 10, A2: 20, A3: 30, B1: "=A1*2", B2: "=A2*2", B3: "=A3*2" },
    });

  it("plans, previews, applies and explains a clean edit", async () => {
    const workbook = baseWorkbook();
    Simulator.of(workbook).recalculate();
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson(
        [
          {
            tool: "formula.set",
            params: { sheet: "S", a1: "C1", formula: "=B1+1" },
            rationale: "add a derived column",
          },
        ],
        "Add a derived column in C1"
      )
    );

    const result = await runAgent(workbook, {
      intent: "add a derived column",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });

    expect(result.outcome).toBe("applied");
    expect(result.verification!.ok).toBe(true);
    expect(workbook.sheet("S")!.get(0, 2)!.formula).toBe("=B1+1");
    expect(result.explanation).toContain("What changed:");
    // The planner must have gone to the strong tier.
    expect(provider.requestsForTier("strong")).toHaveLength(1);
  });

  it("sends the WIL, never a raw grid, to the planner (INV-6)", async () => {
    const { workbook } = threeStatementModel();
    const provider = new MockLlmProvider().script("INTENT", planJson([{ tool: "audit.run", params: {} }]));
    await runAgent(workbook, {
      intent: "look at it",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });
    const prompt = provider.requests[0]!.messages.map((m) => m.content).join("\n");
    expect(prompt).toContain("WORKBOOK SUMMARY");
    expect(prompt).toContain("SHEET Assumptions");
    // A raw grid dump would contain long runs of comma-separated numbers.
    expect(prompt).not.toMatch(/(\d+,){20}/);
  });

  it("stops when the user rejects the plan, changing nothing", async () => {
    const workbook = baseWorkbook();
    const before = JSON.stringify([...workbook.sheet("S")!.cells.entries()]);
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1+1" } }])
    );
    const result = await runAgent(workbook, {
      intent: "add a column",
      provider,
      approver: rejectAll,
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("rejected");
    expect(JSON.stringify([...workbook.sheet("S")!.cells.entries()])).toBe(before);
  });

  it("aborts before writing when the workbook drifted (INV-8)", async () => {
    const workbook = baseWorkbook();
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "formula.set", params: { sheet: "S", a1: "B1", formula: "=A1*5" } }])
    );

    // Simulate a co-author editing the cell AFTER the change set was proposed
    // (and its snapshot taken) but before it is applied — exactly the window
    // the drift check exists to cover.
    let approvals = 0;
    const approver = async () => {
      approvals++;
      if (approvals === 2) {
        workbook.sheet("S")!.set({ row: 0, col: 1, value: 0, formula: "=A1*999" });
      }
      return "approve" as const;
    };

    const result = await runAgent(workbook, {
      intent: "change the multiplier",
      provider,
      approver,
      executor: createDefaultExecutor(),
    });

    expect(result.outcome).toBe("aborted-drift");
    // The concurrent edit must be intact — we never wrote over it.
    expect(workbook.sheet("S")!.get(0, 1)!.formula).toBe("=A1*999");
    expect(result.transcript.join(" ")).toContain("Someone else may be editing");
  });

  it("auto-approves a low-risk plan in a trusted session", async () => {
    const workbook = baseWorkbook();
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([
        { tool: "number_format.set", params: { sheet: "S", a1: "A1:A3", format: "0.00" } },
      ])
    );
    const approver = vi.fn(approveAll);
    const result = await runAgent(workbook, {
      intent: "format the inputs",
      provider,
      approver,
      executor: createDefaultExecutor(),
      trustedSession: true,
    });
    expect(result.outcome).toBe("applied");
    expect(approver).not.toHaveBeenCalled();
  });

  it("still asks for approval on a HIGH-risk plan in a trusted session", async () => {
    const workbook = baseWorkbook();
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "range.clear", params: { sheet: "S", a1: "B1" } }])
    );
    const approver = vi.fn(approveAll);
    await runAgent(workbook, {
      intent: "clear it",
      provider,
      approver,
      executor: createDefaultExecutor(),
      trustedSession: true,
    });
    expect(approver).toHaveBeenCalled();
  });

  it("reports a plan failure instead of guessing", async () => {
    const provider = new MockLlmProvider().script("INTENT", "no json here");
    const result = await runAgent(baseWorkbook(), {
      intent: "do something",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("plan-failed");
  });

  it("refuses an unknown tool emitted by the model", async () => {
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "shell.exec", params: { cmd: "rm -rf /" } }])
    );
    const result = await runAgent(baseWorkbook(), {
      intent: "be helpful",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("plan-failed");
    expect(result.transcript.join(" ")).toContain("Unknown tool");
  });

  it("repairs a failed verification, then succeeds", async () => {
    const workbook = baseWorkbook();
    Simulator.of(workbook).recalculate();

    // First plan writes a formula referencing a missing cell -> #REF!-free but
    // divides by an empty cell, producing #DIV/0! in the blast radius.
    const provider = new MockLlmProvider()
      .script(
        "INTENT",
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/Z9" } }])
      )
      .script(
        "verification failed",
        planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1/2" } }])
      );

    const result = await runAgent(workbook, {
      intent: "add a ratio",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });

    expect(result.repairAttempts).toBeGreaterThanOrEqual(1);
    expect(result.outcome).toBe("applied");
    expect(workbook.sheet("S")!.get(0, 2)!.formula).toBe("=B1/2");
  });

  it("stops repairing after the budget and offers rollback", async () => {
    const workbook = baseWorkbook();
    Simulator.of(workbook).recalculate();
    const before = workbook.sheet("S")!.get(0, 1)!.formula;

    // Every plan produces the same broken formula, so repair never succeeds.
    const broken = planJson([
      { tool: "formula.set", params: { sheet: "S", a1: "B1", formula: "=B1/Z9" } },
    ]);
    const provider = new MockLlmProvider()
      .script("INTENT", broken)
      .script("verification failed", broken);

    const result = await runAgent(workbook, {
      intent: "break it",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
      maxRepairs: 2,
    });

    expect(result.repairAttempts).toBe(2);
    expect(result.outcome).toBe("rolled-back");
    // Rollback restored the original formula.
    expect(workbook.sheet("S")!.get(0, 1)!.formula).toBe(before);
    expect(result.rollbackReport!.restoredCells).toBeGreaterThan(0);
  });

  it("keeps the change when the user declines rollback, and says so", async () => {
    const workbook = baseWorkbook();
    Simulator.of(workbook).recalculate();
    const broken = planJson([
      { tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=1/Z9" } },
    ]);
    const provider = new MockLlmProvider().script("INTENT", broken).script("verification failed", broken);

    let calls = 0;
    const approver = async () => {
      calls++;
      // Approve plan and change set; decline the rollback offer.
      return calls <= 2 ? ("approve" as const) : ("reject" as const);
    };

    const result = await runAgent(workbook, {
      intent: "add a ratio",
      provider,
      approver,
      executor: createDefaultExecutor(),
      maxRepairs: 1,
    });
    expect(result.outcome).toBe("applied-with-warnings");
    expect(result.transcript.join(" ")).toContain("despite the warnings");
  });

  it("refuses to write to a protected sheet, and says why", async () => {
    const workbook = baseWorkbook();
    workbook.sheet("S")!.protectedSheet = true;
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1+1" } }])
    );
    const result = await runAgent(workbook, {
      intent: "add a column",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("blocked-hazard");
    expect(result.transcript.join(" ")).toContain("Unprotect it");
    // Nothing was written.
    expect(workbook.sheet("S")!.get(0, 2)).toBeUndefined();
  });

  it("refuses to write into a merged area's non-anchor cell", async () => {
    const workbook = baseWorkbook();
    workbook.sheet("S")!.merged = [[0, 2, 0, 5]];
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "formula.set", params: { sheet: "S", a1: "D1", formula: "=B1+1" } }])
    );
    const result = await runAgent(workbook, {
      intent: "write into the merged block",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("blocked-hazard");
    expect(result.transcript.join(" ")).toContain("silently ignored");
  });

  it("reports a no-op when the plan produces no edits", async () => {
    const provider = new MockLlmProvider().script("INTENT", planJson([{ tool: "audit.run", params: {} }]));
    const result = await runAgent(baseWorkbook(), {
      intent: "just audit",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
    });
    expect(result.outcome).toBe("no-op");
  });

  it("meters cost across the whole run", async () => {
    const meter = new CostMeter();
    const provider = new MockLlmProvider().script(
      "INTENT",
      planJson([{ tool: "formula.set", params: { sheet: "S", a1: "C1", formula: "=B1+1" } }])
    );
    const result = await runAgent(baseWorkbook(), {
      intent: "add a column",
      provider,
      approver: approveAll,
      executor: createDefaultExecutor(),
      costMeter: meter,
    });
    expect(result.llmCalls).toBeGreaterThanOrEqual(1);
    expect(meter.totalInputTokens).toBeGreaterThan(0);
  });

  it("verifies balance assertions after applying", async () => {
    const { workbook } = threeStatementModel();
    Simulator.of(workbook).recalculate();
    // Breaking equity breaks the tie-out, which the verifier must catch.
    const provider = new MockLlmProvider()
      .script(
        "INTENT",
        planJson([
          { tool: "formula.set", params: { sheet: "BS", a1: "C7", formula: "=B7+IS!C10+5000" } },
        ])
      )
      .script("verification failed", planJson([]));

    const result = await runAgent(workbook, {
      intent: "adjust equity",
      provider,
      approver: async (request) => (request.text.includes("Roll back") ? "reject" : "approve"),
      executor: createDefaultExecutor(),
      assertions: [
        { name: "Balance sheet ties", sheet: "BS", a1: "C9", expected: 0, tolerance: 1e-6 },
      ],
      maxRepairs: 1,
    });

    expect(result.verification!.ok).toBe(false);
    expect(result.verification!.issues.some((issue) => issue.kind === "assertion-failed")).toBe(true);
  });
});

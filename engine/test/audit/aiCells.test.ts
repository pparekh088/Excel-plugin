import { describe, expect, it } from "vitest";
import { inventoryAiCells } from "../../src/audit/aiCells";
import { runAudit } from "../../src/audit/engine";
import { DependencyGraph } from "../../src/graph/graph";
import { workbookOf } from "../helpers/build";

function analyze(cells: Parameters<typeof workbookOf>[0]["S"]) {
  const workbook = workbookOf({ S: cells });
  const graph = DependencyGraph.build(workbook);
  return { workbook, graph, inventory: inventoryAiCells(workbook, graph) };
}

describe("AI-derived value inventory (handoff §8)", () => {
  it("finds generative AI cells", () => {
    const { inventory } = analyze({
      A1: "revenue line",
      B1: '=AI.CLASSIFY(A1,"revenue,cost")',
    });
    expect(inventory.generative).toHaveLength(1);
    expect(inventory.generative[0]!.fn).toBe("AI.CLASSIFY");
    expect(inventory.totalAiCells).toBe(1);
  });

  it("separates AI.FORECAST as computed rather than generated", () => {
    // The distinction is the point: a forecast is statistics, not inference.
    const { inventory } = analyze({
      A1: 10,
      A2: 20,
      A3: 30,
      B1: "=AI.FORECAST(A1:A3,2)",
      C1: '=AI.ASK(A1,"what is this")',
    });
    expect(inventory.computed.map((entry) => entry.fn)).toEqual(["AI.FORECAST"]);
    expect(inventory.generative.map((entry) => entry.fn)).toEqual(["AI.ASK"]);
  });

  it("counts how far an inferred value propagates", () => {
    const { inventory } = analyze({
      A1: "text",
      B1: '=AI.CLASSIFY(A1,"a,b")',
      C1: "=IF(B1=\"a\",100,200)",
      D1: "=C1*2",
    });
    expect(inventory.generative[0]!.dependents).toBeGreaterThan(0);
    expect(inventory.downstreamCells).toBeGreaterThanOrEqual(2);
  });

  it("lists cells stranded by the budget breaker", () => {
    const workbook = workbookOf({ S: { A1: "x" } });
    workbook.sheet("S")!.set({
      row: 1,
      col: 0,
      value: "#AI_BUDGET!",
      formula: '=AI.CLASSIFY(A1,"a,b")',
    });
    const inventory = inventoryAiCells(workbook, DependencyGraph.build(workbook));
    expect(inventory.budgetExhausted).toEqual(["S!A2"]);
  });

  it("reports nothing on a workbook with no AI functions", () => {
    const { inventory } = analyze({ A1: 1, B1: "=A1*2" });
    expect(inventory.totalAiCells).toBe(0);
    expect(inventory.generative).toEqual([]);
  });
});

describe("AUD-013", () => {
  it("is informational when nothing reads the AI value", () => {
    const workbook = workbookOf({
      S: { A1: "text", B1: '=AI.CLASSIFY(A1,"a,b")' },
    });
    const report = runAudit(workbook, { rules: ["AUD-013"] });
    const finding = report.findings.find((item) => item.rule === "AUD-013");
    expect(finding).toBeDefined();
    expect(finding!.severity).toBe("info");
    expect(finding!.explanation).toContain("presentational");
  });

  it("escalates when an inferred value feeds calculations", () => {
    const workbook = workbookOf({
      S: { A1: "text", B1: '=AI.CLASSIFY(A1,"a,b")', C1: '=IF(B1="a",1,2)', D1: "=C1*100" },
    });
    const report = runAudit(workbook, { rules: ["AUD-013"] });
    const finding = report.findings.find((item) => item.rule === "AUD-013");
    expect(finding!.severity).toBe("medium");
    expect(finding!.explanation).toContain("propagating into calculated results");
  });

  it("flags a stranded #AI_BUDGET! cell as high severity", () => {
    const workbook = workbookOf({ S: { A1: "x" } });
    workbook.sheet("S")!.set({
      row: 1,
      col: 0,
      value: "#AI_BUDGET!",
      formula: '=AI.CLASSIFY(A1,"a,b")',
    });
    const report = runAudit(workbook, { rules: ["AUD-013"] });
    const budget = report.findings.find((item) => item.title === "AI budget exhausted");
    expect(budget).toBeDefined();
    expect(budget!.severity).toBe("high");
    expect(budget!.explanation).toContain("Recalculate to start a new cycle");
  });

  it("resolves a real cell position so Trace works", () => {
    const workbook = workbookOf({
      S: { A1: "text", D5: '=AI.CLASSIFY(A1,"a,b")' },
    });
    const report = runAudit(workbook, { rules: ["AUD-013"] });
    const finding = report.findings.find((item) => item.rule === "AUD-013")!;
    expect(finding.row).toBe(4);
    expect(finding.col).toBe(3);
    expect(finding.trace).toEqual([{ sheet: "S", row: 4, col: 3 }]);
  });

  it("still uses zero LLM calls", () => {
    const workbook = workbookOf({
      S: { A1: "text", B1: '=AI.ASK(A1,"summarize")' },
    });
    expect(runAudit(workbook).stats.llmCallsUsed).toBe(0);
  });
});

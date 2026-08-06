import { describe, expect, it } from "vitest";
import { DependencyGraph } from "../../src/graph/graph";
import { buildWil, estimateTokens } from "../../src/wil/serialize";
import { buildSemanticMap, looksLikePeriod } from "../../src/wil/semantic";
import { threeStatementModel, dcfModel } from "../../src/corpus/models";
import { workbookOf } from "../helpers/build";
import { Workbook } from "../../src/model/workbook";

function analyze(workbook: Workbook) {
  const graph = DependencyGraph.build(workbook);
  return { graph, wil: buildWil(workbook, graph) };
}

describe("period detection", () => {
  it.each(["2024", "FY24", "FY2026", "Q1", "Q3-25", "Jan", "January", "Month 3", 2024])(
    "recognizes %s as a period",
    (value) => {
      expect(looksLikePeriod(value as string | number)).toBe(true);
    }
  );

  it.each(["Revenue", "Total", "", "1234567", 0.05, 12, "Q5", "Item 3"])(
    "does not treat %s as a period",
    (value) => {
      expect(looksLikePeriod(value as string | number)).toBe(false);
    }
  );

  it("requires a run of three to call something a time axis", () => {
    const single = workbookOf({ S: { A1: "FY24", A2: "Revenue" } });
    const singleMap = buildSemanticMap(single, DependencyGraph.build(single));
    expect(singleMap.timeAxes).toHaveLength(0);

    const run = workbookOf({ S: { B1: "FY24", C1: "FY25", D1: "FY26" } });
    const runMap = buildSemanticMap(run, DependencyGraph.build(run));
    expect(runMap.timeAxes.length).toBeGreaterThan(0);
  });
});

describe("semantic classification", () => {
  it("separates inputs, calculations, outputs and labels", () => {
    const workbook = workbookOf({
      S: {
        A1: "Growth", // label
        B1: 0.05, // input (read by B2)
        A2: "Revenue",
        B2: "=100*(1+B1)", // calculation (read by B3)
        A3: "Total",
        B3: "=B2*2", // output (nothing reads it)
        D9: 42, // constant with no dependents
      },
    });
    const map = buildSemanticMap(workbook, DependencyGraph.build(workbook));
    const kindAt = (row: number, col: number) =>
      map.regions.find(
        (region) =>
          row >= region.startRow &&
          row <= region.endRow &&
          col >= region.startCol &&
          col <= region.endCol
      )?.kind;

    expect(kindAt(0, 0)).toBe("label");
    expect(kindAt(0, 1)).toBe("input");
    expect(kindAt(1, 1)).toBe("calculation");
    expect(kindAt(2, 1)).toBe("output");
    expect(kindAt(8, 3)).toBe("constant");
  });

  it("records a reason for every region", () => {
    const { workbook } = threeStatementModel();
    const map = buildSemanticMap(workbook, DependencyGraph.build(workbook));
    expect(map.regions.length).toBeGreaterThan(0);
    for (const region of map.regions) {
      expect(region.reason.length).toBeGreaterThan(0);
    }
  });

  it("finds the driver block in a 3-statement model", () => {
    const { workbook } = threeStatementModel();
    const map = buildSemanticMap(workbook, DependencyGraph.build(workbook));
    const assumptionInputs = map.inputs.filter((region) => region.sheet === "Assumptions");
    expect(assumptionInputs.length).toBeGreaterThan(0);
    const totalInputCells = assumptionInputs.reduce((sum, r) => sum + r.cellCount, 0);
    expect(totalInputCells).toBeGreaterThanOrEqual(30);
  });
});

describe("WIL serialization", () => {
  it("summarizes a 3-statement model without any raw grid (INV-6)", () => {
    const { workbook } = threeStatementModel();
    const { wil } = analyze(workbook);
    expect(wil.text).toContain("WORKBOOK");
    expect(wil.text).toContain("SHEET Assumptions");
    expect(wil.text).toContain("SHEET IS");
    expect(wil.text).toContain("KEY CHAINS");
    expect(wil.stats.formulaCells).toBeGreaterThan(50);
    // Run collapsing must actually compress: far fewer nodes than formulas.
    expect(wil.stats.runNodes).toBeLessThan(wil.stats.formulaCells / 2);
  });

  it("states coverage honestly when everything resolved", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const { wil } = analyze(workbook);
    expect(wil.text).toContain("coverage: full");
  });

  it("leads with caveats when coverage is incomplete", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=INDIRECT(\"A1\")", C1: "=[Other.xlsx]S!A1", D1: "#REF!" },
    });
    const { wil } = analyze(workbook);
    expect(wil.text).toContain("COVERAGE CAVEATS");
    expect(wil.text).toContain("opaque nodes");
    expect(wil.text).toContain("NOT complete");
  });

  it("reports OLAP pivots as un-inspectable", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.pivots.push({ name: "P1", sheet: "S", sourceRange: "S!A1:B2", olap: true });
    const { wil } = analyze(workbook);
    expect(wil.text).toContain("OLAP");
  });

  it("respects the token budget and says when it truncated", () => {
    const { workbook } = threeStatementModel();
    const graph = DependencyGraph.build(workbook);
    const tight = buildWil(workbook, graph, { tokenBudget: 400 });
    expect(tight.approxTokens).toBeLessThanOrEqual(500);
    expect(tight.truncated).toBe(true);
    expect(tight.text).toContain("truncated");
  });

  it("stays within the default 6k budget on a multi-sheet model", () => {
    const { workbook } = threeStatementModel();
    const { wil } = analyze(workbook);
    expect(wil.approxTokens).toBeLessThanOrEqual(6000);
  });

  it("lists defined names and tables", () => {
    const { workbook } = dcfModel();
    const { wil } = analyze(workbook);
    expect(wil.text).toContain("DEFINED NAMES");
    expect(wil.text).toContain("WACC");
  });

  it("traces key chains from outputs back toward inputs", () => {
    const { workbook } = dcfModel();
    const { wil } = analyze(workbook);
    const chainSection = wil.text.slice(wil.text.indexOf("KEY CHAINS"));
    expect(chainSection).toContain("<-");
  });

  it("estimates tokens conservatively", () => {
    expect(estimateTokens("a".repeat(360))).toBeGreaterThanOrEqual(100);
  });
});

describe("corpus models", () => {
  it("3-statement model builds a connected graph with a balance check", () => {
    const { workbook } = threeStatementModel();
    const graph = DependencyGraph.build(workbook);
    expect(graph.stats.cycleCount).toBe(0);
    expect(graph.stats.unresolvedRefs).toBe(0);
    // The check row must depend on both totals.
    const check = graph.nodeAt("BS", 8, 1)!;
    expect(graph.trace(check.id).length).toBeGreaterThan(5);
  });

  it("DCF model resolves its named WACC through to per-share value", () => {
    const { workbook } = dcfModel();
    const graph = DependencyGraph.build(workbook);
    const wacc = graph.nodeAt("Inputs", 0, 1)!;
    const impacted = graph.impact(wacc.id);
    const perShare = graph.nodeAt("DCF", 16, 1)!;
    expect(impacted).toContain(perShare.id);
  });

  it("clean models have no error cells", () => {
    for (const { workbook } of [threeStatementModel(), dcfModel()]) {
      expect(DependencyGraph.build(workbook).stats.errorCells).toBe(0);
    }
  });
});

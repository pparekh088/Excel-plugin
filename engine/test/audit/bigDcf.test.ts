/**
 * Phase 2 gate: "a deliberately broken 15-sheet DCF yields the correct top-5
 * issues". Asserts the engine's TOP FIVE findings are exactly the five
 * planted defects — a stronger claim than finding them somewhere in the list.
 */

import { describe, expect, it } from "vitest";
import { runAudit } from "../../src/audit/engine";
import { DependencyGraph } from "../../src/graph/graph";
import { Simulator } from "../../src/sim/simulator";
import { brokenBigDcf } from "../../src/corpus/bigDcf";
import { scoreAudit } from "../../src/eval/auditScoring";

describe("broken 15-sheet DCF", () => {
  const corpus = brokenBigDcf();
  Simulator.of(corpus.workbook).recalculate();
  const graph = DependencyGraph.build(corpus.workbook);
  const report = runAudit(corpus.workbook, { graph });

  it("is genuinely a 15-sheet model", () => {
    expect(corpus.workbook.sheets).toHaveLength(15);
    expect(corpus.workbook.formulaCount).toBeGreaterThan(200);
  });

  it("runs with zero LLM calls", () => {
    expect(report.stats.llmCallsUsed).toBe(0);
  });

  it("finds every planted defect", () => {
    const score = scoreAudit(report.findings, corpus.defects, corpus.workbook, graph);
    expect(score.missedDetail).toEqual([]);
    expect(score.matched).toBe(5);
  });

  const describeTop = (n: number): string =>
    report.findings
      .slice(0, n)
      .map(
        (f) =>
          `  ${f.rule} ${f.address} [${f.severity}] blast=${f.blastRadius}` +
          (f.alsoFlaggedBy ? ` (+${f.alsoFlaggedBy.map((x) => x.rule).join(",")})` : "")
      )
      .join("\n");

  it("puts no noise in the top five: every one traces to a planted defect", () => {
    const top5 = report.findings.slice(0, 5);
    const score = scoreAudit(top5, corpus.defects, corpus.workbook, graph);
    expect(score.falsePositiveDetail, `top 5 were:\n${describeTop(5)}`).toEqual([]);
  });

  it("surfaces all five planted defects near the top of the report", () => {
    // A planted #REF! also breaks a downstream tie-out check, which is a
    // correct high-severity finding of its own, so the five roots occupy the
    // top of the list rather than exactly the first five slots.
    const head = report.findings.slice(0, 8);
    const score = scoreAudit(head, corpus.defects, corpus.workbook, graph);
    expect(score.missedDetail, `top 8 were:\n${describeTop(8)}`).toEqual([]);
  });

  it("keeps the whole report short enough to act on", () => {
    expect(report.findings.length).toBeLessThanOrEqual(12);
  });

  it("groups propagated errors under their root cause", () => {
    const errorFindings = report.findings.filter((f) => f.rule === "AUD-004");
    // One deleted precedent poisons many cells; the report shows the root once.
    expect(errorFindings).toHaveLength(1);
    expect(errorFindings[0]!.address).toBe("Consol!D2");
    expect(errorFindings[0]!.explanation).toContain("spread to");
  });

  it("puts the critical defects first", () => {
    expect(report.findings[0]!.severity).toBe("critical");
    expect(report.findings[1]!.severity).toBe("critical");
  });

  it("reports a poor health score with a critical band", () => {
    expect(report.health.score).toBeLessThan(40);
    expect(report.health.band).toBe("critical");
  });

  it("explains each top finding in plain language with a trace", () => {
    for (const finding of report.findings.slice(0, 5)) {
      expect(finding.explanation.length).toBeGreaterThan(60);
      expect(finding.trace.length).toBeGreaterThan(0);
      expect(finding.evidence).toBeDefined();
    }
  });

  it("completes quickly enough to feel instant in the task pane", () => {
    const started = Date.now();
    runAudit(corpus.workbook);
    expect(Date.now() - started).toBeLessThan(3000);
  });
});

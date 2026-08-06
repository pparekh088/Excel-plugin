/**
 * Verifies the add-in consumes the engine correctly — same import path, same
 * call sequence the task pane uses (extract -> DependencyGraph.build ->
 * runAudit), just with the Office.js extraction replaced by a fixture, since
 * Office.js cannot run headlessly.
 *
 * The point is to catch integration drift (bad exports, alias misconfig,
 * engine API changes) in CI rather than at sideload time.
 */

import { describe, expect, it } from "vitest";
import {
  DependencyGraph,
  brokenBigDcf,
  runAudit,
  Simulator,
  threeStatementModel,
} from "ledger-engine";

describe("add-in -> engine integration", () => {
  it("imports the engine through the package alias", () => {
    expect(typeof runAudit).toBe("function");
    expect(typeof DependencyGraph.build).toBe("function");
  });

  it("audits a clean model with no findings and no LLM calls", () => {
    const { workbook } = threeStatementModel();
    Simulator.of(workbook).recalculate();
    const report = runAudit(workbook, { graph: DependencyGraph.build(workbook) });
    expect(report.findings).toEqual([]);
    expect(report.health.score).toBe(100);
    expect(report.stats.llmCallsUsed).toBe(0);
  });

  it("audits a broken model and returns UI-renderable findings", () => {
    const corpus = brokenBigDcf();
    Simulator.of(corpus.workbook).recalculate();
    const report = runAudit(corpus.workbook, {
      graph: DependencyGraph.build(corpus.workbook),
    });

    expect(report.findings.length).toBeGreaterThan(0);
    for (const finding of report.findings) {
      // Everything the AuditPanel renders must be present.
      expect(finding.address).toMatch(/.+!.+/);
      expect(finding.severity).toBeTruthy();
      expect(finding.explanation.length).toBeGreaterThan(20);
      expect(finding.rule).toMatch(/^AUD-\d{3}$/);
      expect(Array.isArray(finding.trace)).toBe(true);
      // Trace cells drive highlightTrace(), which needs real coordinates.
      for (const cell of finding.trace) {
        expect(typeof cell.sheet).toBe("string");
        expect(cell.row).toBeGreaterThanOrEqual(0);
        expect(cell.col).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("exposes the coverage caveat the UI surfaces as a warning", () => {
    const corpus = brokenBigDcf();
    const report = runAudit(corpus.workbook);
    expect(typeof report.coverage.complete).toBe("boolean");
    expect(report.coverage.caveat.length).toBeGreaterThan(20);
  });
});

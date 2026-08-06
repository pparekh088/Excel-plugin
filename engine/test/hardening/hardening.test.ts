/**
 * Phase 5 hardening: co-authoring chaos, hazards, locale, and scale.
 *
 * The theme is that none of these may fail SILENTLY. A protected sheet, a
 * merged cell, a concurrent edit or a localized formula must each produce a
 * clear message, never a partial write or a quietly wrong result.
 */

import { describe, expect, it } from "vitest";
import {
  applyToWorkbook,
  checkDrift,
  proposeChangeSet,
  rollback,
} from "../../src/changeset/engine";
import { checkHazards, describeHazards } from "../../src/changeset/hazards";
import { DependencyGraph } from "../../src/graph/graph";
import { DE_DE, EN_US, FR_FR, fromEnUs, looksLocalized, toEnUs } from "../../src/locale/normalize";
import { Workbook } from "../../src/model/workbook";
import { parseFormula } from "../../src/parser/parser";
import { serialize } from "../../src/parser/serialize";
import { Simulator } from "../../src/sim/simulator";
import { workbookOf } from "../helpers/build";
import { threeStatementModel } from "../../src/corpus";

describe("co-authoring chaos", () => {
  /** Someone else edits a cell we are about to write. */
  it("catches a concurrent formula change before writing", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 0, formula: "=A1*99" });
    expect(checkDrift(workbook, changeSet).clean).toBe(false);
  });

  it("catches a concurrent DELETE of a cell we planned to change", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    workbook.sheet("S")!.delete(0, 1);
    const drift = checkDrift(workbook, changeSet);
    expect(drift.clean).toBe(false);
    expect(drift.entries[0]!.address).toBe("S!B1");
  });

  it("catches a cell APPEARING where we expected an empty one", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 5, col: 5, formula: "=A1" }],
    });
    workbook.sheet("S")!.set({ row: 5, col: 5, value: "someone else's work" });
    expect(checkDrift(workbook, changeSet).clean).toBe(false);
  });

  it("survives interleaved edits: rollback still restores the pre-change state", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=A1*2", C1: "=B1+1", D1: "=C1*3" },
    });
    Simulator.of(workbook).recalculate();

    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 2, formula: "=B1+100" }],
    });
    applyToWorkbook(workbook, changeSet);
    Simulator.of(workbook).recalculate();

    // A co-author edits a DIFFERENT cell after we applied.
    workbook.sheet("S")!.set({ row: 0, col: 0, value: 50 });
    Simulator.of(workbook).recalculate();

    rollback(workbook, changeSet);
    Simulator.of(workbook).recalculate();

    // Our change is undone; their unrelated edit is untouched.
    expect(workbook.sheet("S")!.get(0, 2)!.formula).toBe("=B1+1");
    expect(workbook.sheet("S")!.get(0, 0)!.value).toBe(50);
  });

  it("does not mistake a recalculation for a co-author edit", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    Simulator.of(workbook).recalculate();
    expect(checkDrift(workbook, changeSet).clean).toBe(true);
  });

  it("detects drift across many cells and reports each one", () => {
    const cells: Record<string, string | number> = {};
    for (let row = 1; row <= 20; row++) cells[`A${row}`] = row;
    for (let row = 1; row <= 20; row++) cells[`B${row}`] = `=A${row}*2`;
    const workbook = workbookOf({ S: cells });

    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: Array.from({ length: 20 }, (_, index) => ({
        kind: "setFormula" as const,
        sheet: "S",
        row: index,
        col: 1,
        formula: "=A1*5",
      })),
    });
    // Three cells changed underneath us.
    for (const row of [3, 7, 11]) {
      workbook.sheet("S")!.set({ row, col: 1, value: 0, formula: "=A1*777" });
    }
    const drift = checkDrift(workbook, changeSet);
    expect(drift.entries).toHaveLength(3);
  });
});

describe("workbook hazards", () => {
  it("blocks a write to a protected sheet, and says to unprotect it", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.sheet("S")!.protectedSheet = true;
    const report = checkHazards(workbook, [
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 },
    ]);
    expect(report.blocked).toBe(true);
    expect(report.hazards[0]!.message).toContain("Unprotect it");
    // Never silently attempt to remove protection.
    expect(report.hazards[0]!.message).toContain("will not attempt");
  });

  it("blocks a write to a non-anchor cell of a merged area", () => {
    const workbook = workbookOf({ S: { A1: "merged title" } });
    workbook.sheet("S")!.merged = [[0, 0, 0, 3]];
    const report = checkHazards(workbook, [
      { kind: "setValue", sheet: "S", row: 0, col: 2, value: "x" },
    ]);
    expect(report.blocked).toBe(true);
    expect(report.hazards[0]!.message).toContain("silently ignored");
  });

  it("allows a write to the anchor cell of a merged area", () => {
    const workbook = workbookOf({ S: { A1: "merged title" } });
    workbook.sheet("S")!.merged = [[0, 0, 0, 3]];
    const report = checkHazards(workbook, [
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: "x" },
    ]);
    expect(report.blocked).toBe(false);
  });

  it("blocks a write to a sheet that no longer exists", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const report = checkHazards(workbook, [
      { kind: "setValue", sheet: "Gone", row: 0, col: 0, value: 1 },
    ]);
    expect(report.blocked).toBe(true);
    expect(report.hazards[0]!.message).toContain("renamed or deleted");
  });

  it("warns about a table's calculated column without blocking", () => {
    const workbook = workbookOf({
      Data: {
        A1: "Qty", B1: "Total",
        A2: 2, B2: "=A2*10",
        A3: 3, B3: "=A3*10",
        A4: 4, B4: "=A4*10",
      },
    });
    workbook.tables.push({
      name: "T", sheet: "Data", headerRow: 0, startRow: 0, endRow: 3,
      startCol: 0, endCol: 1, hasTotals: false,
      columns: [{ name: "Qty", col: 0 }, { name: "Total", col: 1 }],
    });
    const report = checkHazards(workbook, [
      { kind: "setFormula", sheet: "Data", row: 2, col: 1, formula: "=A3*20" },
    ]);
    expect(report.blocked).toBe(false);
    expect(report.hazards[0]!.severity).toBe("warning");
    expect(report.hazards[0]!.message).toContain("calculated column");
  });

  it("reports nothing on an ordinary write", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const report = checkHazards(workbook, [
      { kind: "setValue", sheet: "S", row: 5, col: 5, value: 1 },
    ]);
    expect(report.hazards).toEqual([]);
    expect(describeHazards(report)).toBe("");
  });

  it("summarizes hazards for the preview", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.sheet("S")!.protectedSheet = true;
    const text = describeHazards(
      checkHazards(workbook, [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 }])
    );
    expect(text).toContain("Cannot apply");
  });
});

describe("locale normalization (Q-003)", () => {
  it.each([
    ["=SUMME(A1:A10)", "=SUM(A1:A10)"],
    ["=WENN(A1>0;1;2)", "=IF(A1>0,1,2)"],
    ["=RUNDEN(A1;2)", "=ROUND(A1,2)"],
    ["=SVERWEIS(A1;B:D;2;FALSCH)", "=VLOOKUP(A1,B:D,2,FALSCH)"],
    ["=A1*1,5", "=A1*1.5"],
    ["=SUMME(A1;A2)*2,25", "=SUM(A1,A2)*2.25"],
  ])("converts German %s to %s", (input, expected) => {
    expect(toEnUs(input, DE_DE)).toBe(expected);
  });

  it.each([
    ["=SOMME(A1:A10)", "=SUM(A1:A10)"],
    ["=SI(A1>0;1;2)", "=IF(A1>0,1,2)"],
    ["=SOMME.SI(A:A;\">5\")", '=SUMIF(A:A,">5")'],
  ])("converts French %s to %s", (input, expected) => {
    expect(toEnUs(input, FR_FR)).toBe(expected);
  });

  it("never touches separators inside string literals", () => {
    // The comma and period here are the user's TEXT, not syntax.
    expect(toEnUs('=WENN(A1;"a;b";"c,d")', DE_DE)).toBe('=IF(A1,"a;b","c,d")');
  });

  it("never touches quoted sheet names", () => {
    expect(toEnUs("=SUMME('My; Sheet'!A1:A2)", DE_DE)).toBe("=SUM('My; Sheet'!A1:A2)");
  });

  it("passes unknown function names through rather than corrupting them", () => {
    expect(toEnUs("=NICHTBEKANNT(A1;2)", DE_DE)).toBe("=NICHTBEKANNT(A1,2)");
  });

  it("converts back for display", () => {
    expect(fromEnUs("=SUM(A1,A2)", DE_DE)).toBe("=SUMME(A1;A2)");
    expect(fromEnUs("=IF(A1>0,1,2)", FR_FR)).toBe("=SI(A1>0;1;2)");
  });

  it("is a no-op for en-US in both directions", () => {
    expect(toEnUs("=SUM(A1,A2)", EN_US)).toBe("=SUM(A1,A2)");
    expect(fromEnUs("=SUM(A1,A2)", EN_US)).toBe("=SUM(A1,A2)");
  });

  it("round-trips a formula through a locale and back", () => {
    const original = "=IF(A1>0,SUM(B1:B10),0)";
    const german = fromEnUs(original, DE_DE);
    expect(german).toContain("WENN");
    expect(toEnUs(german, DE_DE)).toBe(original);
  });

  it("normalized formulas parse cleanly", () => {
    for (const localized of ["=SUMME(A1:A10)", "=WENN(A1>0;1;2)", "=MITTELWERT(A1;A2)"]) {
      const parsed = parseFormula(toEnUs(localized, DE_DE));
      expect(parsed.ok, localized).toBe(true);
      expect(serialize(parsed.ast).length).toBeGreaterThan(0);
    }
  });

  it("detects a formula that looks localized", () => {
    expect(looksLocalized("=SUMME(A1)", DE_DE)).toBe(true);
    expect(looksLocalized("=SUM(A1)", DE_DE)).toBe(false);
  });
});

describe("scale and stress", () => {
  it("builds a graph over 100k formulas within the time budget", () => {
    const workbook = new Workbook("stress");
    const sheet = workbook.addSheet("Big");
    for (let row = 0; row < 100_000; row++) {
      sheet.set({ row, col: 0, value: row });
      sheet.set({ row, col: 1, value: 0, formula: `=A${row + 1}*2` });
    }
    const started = Date.now();
    const graph = DependencyGraph.build(workbook);
    const elapsed = Date.now() - started;

    expect(graph.stats.formulaCells).toBe(100_000);
    // The whole column is one fill pattern: it must collapse to one node.
    expect(graph.stats.runCount).toBe(1);
    expect(elapsed).toBeLessThan(30_000);
  });

  it("handles a workbook with many sheets", () => {
    const workbook = new Workbook("wide");
    for (let index = 0; index < 60; index++) {
      const sheet = workbook.addSheet(`S${index}`);
      sheet.set({ row: 0, col: 0, value: index });
      sheet.set({ row: 0, col: 1, value: 0, formula: "=A1*2" });
    }
    const graph = DependencyGraph.build(workbook);
    expect(graph.stats.formulaCells).toBe(60);
    expect(graph.stats.cycleCount).toBe(0);
  });

  it("handles a deep dependency chain without stack overflow", () => {
    const workbook = new Workbook("deep");
    const sheet = workbook.addSheet("S");
    sheet.set({ row: 0, col: 0, value: 1 });
    for (let row = 1; row < 20_000; row++) {
      sheet.set({ row, col: 0, value: 0, formula: `=A${row}+1` });
    }
    const graph = DependencyGraph.build(workbook);
    expect(graph.stats.cycleCount).toBe(0);
    expect(() => graph.impact(0)).not.toThrow();
  });

  it("keeps a change set over 10k cells manageable", () => {
    const { workbook } = threeStatementModel();
    const edits = Array.from({ length: 10_000 }, (_, index) => ({
      kind: "setNumberFormat" as const,
      sheet: "Assumptions",
      row: index % 100,
      col: Math.floor(index / 100),
      numberFormat: "0.0%",
    }));
    const started = Date.now();
    const changeSet = proposeChangeSet(workbook, {
      intent: "format everything",
      summary: "format",
      edits,
    });
    expect(Date.now() - started).toBeLessThan(10_000);
    expect(changeSet.risk).toBe("low");
    expect(changeSet.snapshots.length).toBeGreaterThan(0);
  });
});

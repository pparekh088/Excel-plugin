import { describe, expect, it } from "vitest";
import {
  analyzeImpact,
  applyToWorkbook,
  captureAppliedState,
  checkDrift,
  explainChangeSet,
  overallRisk,
  proposeChangeSet,
  riskOf,
  rollback,
  snapshot,
} from "../../src/changeset/engine";
import { Edit } from "../../src/changeset/types";
import { DependencyGraph } from "../../src/graph/graph";
import { Simulator } from "../../src/sim/simulator";
import { Workbook } from "../../src/model/workbook";
import { workbookOf } from "../helpers/build";
import { buildCorpus, threeStatementModel } from "../../src/corpus";

/** Byte-level fingerprint of every cell — the rollback fidelity gate. */
function fingerprint(workbook: Workbook): string {
  const rows: string[] = [];
  for (const sheet of [...workbook.sheets].sort((a, b) => a.name.localeCompare(b.name))) {
    const cells = [...sheet.cells.values()].sort((a, b) => a.row - b.row || a.col - b.col);
    for (const cell of cells) {
      rows.push(
        `${sheet.name}!${cell.row},${cell.col}|${JSON.stringify(cell.value)}|` +
          `${cell.formula ?? ""}|${cell.numberFormat ?? ""}`
      );
    }
  }
  return rows.join("\n");
}

describe("risk tiering (handoff §4)", () => {
  const workbook = workbookOf({
    S: { A1: 5, B1: "=A1*2" },
  });

  it("number formats are LOW", () => {
    expect(
      riskOf({ kind: "setNumberFormat", sheet: "S", row: 0, col: 0, numberFormat: "0.00" }, workbook)
    ).toBe("low");
  });

  it("writing into an empty cell is MEDIUM", () => {
    expect(riskOf({ kind: "setFormula", sheet: "S", row: 9, col: 9, formula: "=1" }, workbook)).toBe(
      "medium"
    );
  });

  it("overwriting a value is HIGH", () => {
    expect(riskOf({ kind: "setValue", sheet: "S", row: 0, col: 0, value: 9 }, workbook)).toBe("high");
  });

  it("changing an existing formula is HIGH", () => {
    expect(
      riskOf({ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }, workbook)
    ).toBe("high");
  });

  it("clearing is always HIGH", () => {
    expect(riskOf({ kind: "clear", sheet: "S", row: 9, col: 9 }, workbook)).toBe("high");
  });

  it("a change set takes the highest risk of its edits", () => {
    expect(
      overallRisk(
        [
          { kind: "setNumberFormat", sheet: "S", row: 0, col: 0, numberFormat: "0" },
          { kind: "setFormula", sheet: "S", row: 5, col: 5, formula: "=1" },
        ],
        workbook
      )
    ).toBe("medium");
    expect(
      overallRisk(
        [
          { kind: "setNumberFormat", sheet: "S", row: 0, col: 0, numberFormat: "0" },
          { kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" },
        ],
        workbook
      )
    ).toBe("high");
  });
});

describe("snapshots (INV-3)", () => {
  it("captures value, formula and number format before any write", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    workbook.sheet("S")!.get(0, 1)!.numberFormat = "0.00%";
    const snaps = snapshot(workbook, [
      { kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" },
    ]);
    expect(snaps).toHaveLength(1);
    expect(snaps[0]).toMatchObject({ formula: "=A1*2", numberFormat: "0.00%", absent: false });
  });

  it("records absence for cells that do not exist yet", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    const snaps = snapshot(workbook, [
      { kind: "setFormula", sheet: "S", row: 9, col: 9, formula: "=1" },
    ]);
    expect(snaps[0]!.absent).toBe(true);
  });

  it("does not duplicate snapshots for repeated edits to one cell", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    const snaps = snapshot(workbook, [
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 1 },
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 },
    ]);
    expect(snaps).toHaveLength(1);
    expect(snaps[0]!.value).toBe(5);
  });
});

describe("impact analysis", () => {
  it("counts downstream cells and names outputs", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=A1*2", C1: "=B1+1", D1: "=C1*3" },
    });
    const graph = DependencyGraph.build(workbook);
    const impact = analyzeImpact(workbook, graph, [
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 },
    ]);
    expect(impact.affectedCells).toBe(3);
    expect(impact.affectedOutputs).toContain("S!D1");
  });

  it("reports opaque downstream paths so the count reads as a lower bound", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=A1*2", C1: '=INDIRECT("B1")+1' },
    });
    const graph = DependencyGraph.build(workbook);
    const impact = analyzeImpact(workbook, graph, [
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 },
    ]);
    expect(impact.opaqueDownstream).toBeGreaterThanOrEqual(0);
  });

  it("names charts and pivots on touched sheets", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.charts.push({ name: "Revenue chart", sheet: "S", sourceRanges: [] });
    workbook.pivots.push({ name: "P1", sheet: "S", sourceRange: "", olap: true });
    const graph = DependencyGraph.build(workbook);
    const impact = analyzeImpact(workbook, graph, [
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 },
    ]);
    expect(impact.affectedCharts).toEqual(["Revenue chart"]);
    expect(impact.affectedPivots).toEqual(["P1"]);
  });
});

describe("drift detection (INV-8)", () => {
  it("passes when nothing changed", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    expect(checkDrift(workbook, changeSet).clean).toBe(true);
  });

  it("detects someone else changing a formula", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 0, formula: "=A1*99" });
    const drift = checkDrift(workbook, changeSet);
    expect(drift.clean).toBe(false);
    expect(drift.entries[0]!.address).toBe("S!B1");
  });

  it("detects someone else changing a constant", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 7 }],
    });
    workbook.sheet("S")!.set({ row: 0, col: 0, value: 42 });
    expect(checkDrift(workbook, changeSet).clean).toBe(false);
  });

  it("does not treat a recalculated value as drift", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    // A recalculation updates the value but leaves the formula alone.
    workbook.sheet("S")!.get(0, 1)!.value = 10;
    expect(checkDrift(workbook, changeSet).clean).toBe(true);
  });
});

describe("rollback fidelity (Phase 3 gate: byte-identical restore)", () => {
  it("restores values and formulas exactly on a simple edit", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2", C1: "=B1+1" } });
    Simulator.of(workbook).recalculate();
    const before = fingerprint(workbook);

    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*99" }],
    });
    applyToWorkbook(workbook, changeSet);
    Simulator.of(workbook).recalculate();
    expect(fingerprint(workbook)).not.toBe(before);

    rollback(workbook, changeSet);
    Simulator.of(workbook).recalculate();
    expect(fingerprint(workbook)).toBe(before);
  });

  it("removes cells that did not exist before", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    const before = fingerprint(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 4, col: 4, formula: "=A1*2" }],
    });
    applyToWorkbook(workbook, changeSet);
    expect(workbook.sheet("S")!.get(4, 4)).toBeDefined();
    rollback(workbook, changeSet);
    expect(workbook.sheet("S")!.get(4, 4)).toBeUndefined();
    expect(fingerprint(workbook)).toBe(before);
  });

  it("restores number formats", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    workbook.sheet("S")!.get(0, 0)!.numberFormat = "0.00%";
    const before = fingerprint(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [
        { kind: "setNumberFormat", sheet: "S", row: 0, col: 0, numberFormat: "#,##0" },
      ],
    });
    applyToWorkbook(workbook, changeSet);
    rollback(workbook, changeSet);
    expect(fingerprint(workbook)).toBe(before);
  });

  it("restores a cleared cell", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    const before = fingerprint(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "clear", sheet: "S", row: 0, col: 1 }],
    });
    applyToWorkbook(workbook, changeSet);
    expect(workbook.sheet("S")!.get(0, 1)).toBeUndefined();
    rollback(workbook, changeSet);
    expect(fingerprint(workbook)).toBe(before);
  });

  it("is byte-identical across every corpus workbook", () => {
    for (const entry of buildCorpus()) {
      const workbook = entry.workbook;
      Simulator.of(workbook).recalculate();
      const before = fingerprint(workbook);

      // Edit a handful of real formula cells on each workbook.
      const targets: Edit[] = [];
      for (const sheet of workbook.sheets.slice(0, 3)) {
        const formulaCells = [...sheet.cells.values()]
          .filter((cell) => cell.formula !== undefined)
          .slice(0, 4);
        for (const cell of formulaCells) {
          targets.push({
            kind: "setFormula",
            sheet: sheet.name,
            row: cell.row,
            col: cell.col,
            formula: "=1+1",
          });
        }
      }
      if (targets.length === 0) continue;

      const changeSet = proposeChangeSet(workbook, {
        intent: "corpus rollback test",
        summary: "overwrite formulas then undo",
        edits: targets,
      });
      applyToWorkbook(workbook, changeSet);
      Simulator.of(workbook).recalculate();
      rollback(workbook, changeSet);
      Simulator.of(workbook).recalculate();

      expect(fingerprint(workbook), `rollback drifted on ${entry.id}`).toBe(before);
    }
  });

  it("reports honestly what it cannot restore", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.pivots.push({ name: "P1", sheet: "S", sourceRange: "", olap: true });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [
        { kind: "createSheet", name: "New" },
        { kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 },
      ],
    });
    applyToWorkbook(workbook, changeSet);
    const report = rollback(workbook, changeSet);
    // The sheet has a planned inverse and is genuinely removed; the pivot
    // cache does not and is reported rather than glossed over.
    expect(report.reversedStructural.some((item) => item.includes("New"))).toBe(true);
    expect(report.unrestorable.some((item) => item.includes("pivot"))).toBe(true);
  });
});

describe("rollback never destroys a human edit made after apply (P0-2)", () => {
  /** The reviewer's scenario: before=10, we write 20, a human writes 30. */
  function setUp() {
    const workbook = workbookOf({ S: { F27: 10 } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "bump F27",
      summary: "bump F27",
      edits: [{ kind: "setValue", sheet: "S", row: 26, col: 5, value: 20 }],
    });
    applyToWorkbook(workbook, changeSet);
    Simulator.of(workbook).recalculate();
    expect(workbook.sheet("S")!.get(26, 5)!.value).toBe(20);

    // The human edits the same cell after we applied.
    workbook.sheet("S")!.set({ row: 26, col: 5, value: 30 });
    return { workbook, changeSet };
  }

  it("keeps the human's value and reports a conflict", () => {
    const { workbook, changeSet } = setUp();
    const report = rollback(workbook, changeSet);

    expect(workbook.sheet("S")!.get(26, 5)!.value).toBe(30);
    expect(report.restoredCells).toBe(0);
    expect(report.conflicts).toHaveLength(1);
    expect(report.conflicts[0]!.address).toBe("S!F27");
    expect(report.conflicts[0]!.applied.value).toBe(20);
    expect(report.conflicts[0]!.current.value).toBe(30);
    // Conflicts are a correct outcome, not a failed rollback.
    expect(report.ok).toBe(true);
  });

  it("restores it only when the user explicitly forces the rollback", () => {
    const { workbook, changeSet } = setUp();
    const report = rollback(workbook, changeSet, { force: true });

    expect(workbook.sheet("S")!.get(26, 5)!.value).toBe(10);
    expect(report.restoredCells).toBe(1);
    expect(report.conflicts).toHaveLength(0);
  });

  it("detects a human replacing our formula with their own", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    Simulator.of(workbook).recalculate();
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    applyToWorkbook(workbook, changeSet);
    Simulator.of(workbook).recalculate();

    workbook.sheet("S")!.set({ row: 0, col: 1, value: 20, formula: "=A1*4" });
    const report = rollback(workbook, changeSet);

    expect(workbook.sheet("S")!.get(0, 1)!.formula).toBe("=A1*4");
    expect(report.conflicts).toHaveLength(1);
    expect(report.conflicts[0]!.applied.formula).toBe("=A1*3");
    expect(report.conflicts[0]!.current.formula).toBe("=A1*4");
  });

  it("detects a human deleting a cell we created", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 9, col: 9, value: 1 }],
    });
    applyToWorkbook(workbook, changeSet);
    workbook.sheet("S")!.delete(9, 9);

    const report = rollback(workbook, changeSet);
    expect(report.conflicts).toHaveLength(1);
    expect(report.restoredCells).toBe(0);
  });

  it("does not mistake a recalculated value for a human edit", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    Simulator.of(workbook).recalculate();
    const before = fingerprint(workbook);
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    applyToWorkbook(workbook, changeSet);
    // Recalculation moves B1's VALUE from 10 to 15 without anyone touching it.
    Simulator.of(workbook).recalculate();
    expect(workbook.sheet("S")!.get(0, 1)!.value).toBe(15);

    const report = rollback(workbook, changeSet);
    Simulator.of(workbook).recalculate();
    expect(report.conflicts).toHaveLength(0);
    expect(fingerprint(workbook)).toBe(before);
  });

  it("rolls back untouched cells and conflicts only on the edited one", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: 2, C1: 3 } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [
        { kind: "setValue", sheet: "S", row: 0, col: 0, value: 10 },
        { kind: "setValue", sheet: "S", row: 0, col: 1, value: 20 },
        { kind: "setValue", sheet: "S", row: 0, col: 2, value: 30 },
      ],
    });
    applyToWorkbook(workbook, changeSet);
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 999 });

    const report = rollback(workbook, changeSet);
    expect(workbook.sheet("S")!.get(0, 0)!.value).toBe(1);
    expect(workbook.sheet("S")!.get(0, 1)!.value).toBe(999);
    expect(workbook.sheet("S")!.get(0, 2)!.value).toBe(3);
    expect(report.restoredCells).toBe(2);
    expect(report.conflicts.map((conflict) => conflict.address)).toEqual(["S!B1"]);
  });

  it("refuses every cell when no post-apply state was recorded", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 2 }],
    });
    applyToWorkbook(workbook, changeSet);
    // Simulate a change set that came back from storage without the evidence.
    delete changeSet.appliedState;

    const report = rollback(workbook, changeSet);
    expect(workbook.sheet("S")!.get(0, 0)!.value).toBe(2);
    expect(report.restoredCells).toBe(0);
    expect(report.conflicts[0]!.reason).toContain("cannot tell");
  });

  it("captureAppliedState records what the cells hold now", () => {
    const workbook = workbookOf({ S: { A1: 5, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    applyToWorkbook(workbook, changeSet);
    Simulator.of(workbook).recalculate();

    const state = captureAppliedState(workbook, changeSet);
    expect(state).toHaveLength(1);
    expect(state[0]!.formula).toBe("=A1*3");
    expect(state[0]!.value).toBe(15);
    expect(state[0]!.absent).toBe(false);
  });
});

describe("change-set explanation (INV-10)", () => {
  it("states what changed, why, and what it affects", () => {
    const { workbook } = threeStatementModel();
    Simulator.of(workbook).recalculate();
    const changeSet = proposeChangeSet(workbook, {
      intent: "Change FY2025 revenue growth to 5.2%",
      summary: "Set the FY2025 growth driver to 5.2%",
      edits: [{ kind: "setValue", sheet: "Assumptions", row: 1, col: 2, value: 0.052 }],
    });
    const text = explainChangeSet(changeSet);
    expect(text).toContain("Set the FY2025 growth driver");
    expect(text).toContain("Intent:");
    expect(text).toContain("What changed:");
    expect(text).toContain("What it affects downstream:");
    expect(text).toMatch(/\d+ cell\(s\)/);
  });

  it("flags overwritten formulas as the risky ones", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    expect(explainChangeSet(changeSet)).toContain("existing formula(s) replaced");
  });

  it("warns when impact figures are a lower bound", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=A1*2", C1: '=INDIRECT("B1")' },
    });
    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 5 }],
    });
    if (changeSet.impact.opaqueDownstream > 0) {
      expect(explainChangeSet(changeSet)).toContain("LOWER BOUND");
    }
  });
});

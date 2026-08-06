import { describe, expect, it } from "vitest";
import {
  accountDestruction,
  gradeCells,
  scoreDetection,
  snapshotWorkbook,
  summarize,
} from "../../src/eval/harness";
import { workbookOf } from "../helpers/build";

describe("destruction accounting (cells-destroyed must be 0)", () => {
  it("reports nothing when the workbook is untouched", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const before = snapshotWorkbook(workbook);
    expect(accountDestruction(before, workbook, new Set()).destroyed).toEqual([]);
  });

  it("flags a formula replaced by a constant", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const before = snapshotWorkbook(workbook);
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 42 });
    expect(accountDestruction(before, workbook, new Set()).destroyed).toEqual(["S!0,1"]);
  });

  it("flags a formula that was changed", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const before = snapshotWorkbook(workbook);
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 0, formula: "=A1*3" });
    expect(accountDestruction(before, workbook, new Set()).destroyed).toEqual(["S!0,1"]);
  });

  it("flags a deleted cell", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const before = snapshotWorkbook(workbook);
    workbook.sheet("S")!.delete(0, 1);
    expect(accountDestruction(before, workbook, new Set()).destroyed).toEqual(["S!0,1"]);
  });

  it("does not flag changes the task declared as intentional", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const before = snapshotWorkbook(workbook);
    workbook.sheet("S")!.set({ row: 0, col: 1, value: 0, formula: "=A1*3" });
    const report = accountDestruction(before, workbook, new Set(["S!0,1"]));
    expect(report.destroyed).toEqual([]);
    expect(report.intentional).toEqual(["S!0,1"]);
  });

  it("does not treat recalculated values as destruction", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2" } });
    const before = snapshotWorkbook(workbook);
    workbook.sheet("S")!.get(0, 1)!.value = 2; // recalc result
    expect(accountDestruction(before, workbook, new Set()).destroyed).toEqual([]);
  });

  it("does not treat editing a constant as destruction", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const before = snapshotWorkbook(workbook);
    workbook.sheet("S")!.set({ row: 0, col: 0, value: 99 });
    expect(accountDestruction(before, workbook, new Set()).destroyed).toEqual([]);
  });
});

describe("cell grading", () => {
  const workbook = workbookOf({ Model: { B2: "=A1*1.05", C3: 42 } });

  it("passes when formula and value match", () => {
    const result = gradeCells(workbook, [
      { sheet: "Model", a1: "B2", formula: "=A1*1.05" },
      { sheet: "Model", a1: "C3", value: 42 },
    ]);
    expect(result.passed).toBe(true);
  });

  it("fails on a formula mismatch and explains why", () => {
    const result = gradeCells(workbook, [
      { sheet: "Model", a1: "B2", formula: "=A1*1.07" },
    ]);
    expect(result.passed).toBe(false);
    expect(result.checks[0]!.detail).toContain("expected =A1*1.07");
  });

  it("compares numbers within tolerance", () => {
    const wb = workbookOf({ S: { A1: 0.1 + 0.2 } });
    expect(gradeCells(wb, [{ sheet: "S", a1: "A1", value: 0.3, tolerance: 1e-9 }]).passed).toBe(
      true
    );
  });

  it("ignores whitespace differences in formulas", () => {
    const wb = workbookOf({ S: { A1: "=SUM(B1, C1)" } });
    expect(gradeCells(wb, [{ sheet: "S", a1: "A1", formula: "=SUM(B1,C1)" }]).passed).toBe(true);
  });

  it("fails cleanly for a missing cell", () => {
    expect(gradeCells(workbook, [{ sheet: "Model", a1: "Z9", value: 1 }]).passed).toBe(false);
  });
});

describe("detection scoring", () => {
  it("scores a perfect run", () => {
    const score = scoreDetection(
      [{ rule: "AUD-002", address: "IS!E2" }],
      [{ rule: "AUD-002", address: "IS!E2" }]
    );
    expect(score).toMatchObject({ truePositives: 1, falsePositives: 0, falseNegatives: 0 });
    expect(score.precision).toBe(1);
    expect(score.recall).toBe(1);
    expect(score.f1).toBe(1);
  });

  it("counts false positives and negatives", () => {
    const score = scoreDetection(
      [
        { rule: "AUD-002", address: "IS!E2" },
        { rule: "AUD-002", address: "IS!F9" }, // spurious
      ],
      [
        { rule: "AUD-002", address: "IS!E2" },
        { rule: "AUD-003", address: "BS!D6" }, // missed
      ]
    );
    expect(score.truePositives).toBe(1);
    expect(score.falsePositives).toBe(1);
    expect(score.falseNegatives).toBe(1);
    expect(score.precision).toBe(0.5);
    expect(score.recall).toBe(0.5);
  });

  it("treats a clean workbook with no findings as perfect", () => {
    const score = scoreDetection([], []);
    expect(score.precision).toBe(1);
    expect(score.recall).toBe(1);
  });

  it("scores any finding on a clean workbook as zero precision", () => {
    const score = scoreDetection([{ rule: "AUD-001", address: "S!A1" }], []);
    expect(score.precision).toBe(0);
    expect(score.falsePositives).toBe(1);
  });

  it("matches addresses case-insensitively", () => {
    const score = scoreDetection(
      [{ rule: "AUD-004", address: "cf!d2" }],
      [{ rule: "AUD-004", address: "CF!D2" }]
    );
    expect(score.truePositives).toBe(1);
  });
});

describe("run summarization", () => {
  it("aggregates success rate, destruction and cost", () => {
    const summary = summarize([
      {
        taskId: "t1",
        taskClass: "edit",
        success: true,
        cellsDestroyed: 0,
        destroyedAddresses: [],
        toolCalls: 4,
        latencyMs: 1200,
        costUsd: 0.02,
        detail: "",
      },
      {
        taskId: "t2",
        taskClass: "edit",
        success: false,
        cellsDestroyed: 2,
        destroyedAddresses: ["S!1,1", "S!1,2"],
        toolCalls: 8,
        latencyMs: 2400,
        costUsd: 0.05,
        detail: "",
      },
    ]);
    expect(summary.successRate).toBe(0.5);
    expect(summary.totalCellsDestroyed).toBe(2);
    expect(summary.meanToolCalls).toBe(6);
    expect(summary.totalCostUsd).toBeCloseTo(0.07);
  });
});

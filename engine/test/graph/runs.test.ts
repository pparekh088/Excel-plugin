import { describe, expect, it } from "vitest";
import { RunCell, findRuns } from "../../src/graph/runs";

function cells(spec: Array<[number, number, string]>): RunCell[] {
  return spec.map(([row, col, signature]) => ({ row, col, signature }));
}

describe("findRuns", () => {
  it("collapses a filled column into one run", () => {
    const runs = findRuns(
      cells([
        [0, 2, "RC[-1]*2"],
        [1, 2, "RC[-1]*2"],
        [2, 2, "RC[-1]*2"],
        [3, 2, "RC[-1]*2"],
      ])
    );
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ startRow: 0, endRow: 3, startCol: 2, endCol: 2, cellCount: 4 });
  });

  it("collapses a filled row into one run", () => {
    const runs = findRuns(
      cells([
        [5, 0, "R[-1]C"],
        [5, 1, "R[-1]C"],
        [5, 2, "R[-1]C"],
      ])
    );
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ startCol: 0, endCol: 2, startRow: 5, endRow: 5 });
  });

  it("collapses a rectangular block", () => {
    const spec: Array<[number, number, string]> = [];
    for (let row = 0; row < 3; row++) {
      for (let col = 0; col < 4; col++) spec.push([row, col, "RC[-1]+R[-1]C"]);
    }
    const runs = findRuns(cells(spec));
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ startRow: 0, endRow: 2, startCol: 0, endCol: 3, cellCount: 12 });
  });

  it("splits a run where one cell breaks the pattern (AUD-001 basis)", () => {
    const runs = findRuns(
      cells([
        [0, 1, "RC[-1]*2"],
        [1, 1, "RC[-1]*2"],
        [2, 1, "HARDCODED"], // the odd one out
        [3, 1, "RC[-1]*2"],
        [4, 1, "RC[-1]*2"],
      ])
    );
    expect(runs).toHaveLength(3);
    expect(runs.map((r) => [r.startRow, r.endRow, r.cellCount])).toEqual([
      [0, 1, 2],
      [2, 2, 1],
      [3, 4, 2],
    ]);
  });

  it("does not merge non-adjacent cells with the same signature", () => {
    const runs = findRuns(
      cells([
        [0, 0, "SAME"],
        [1, 0, "SAME"],
        [10, 0, "SAME"], // gap
        [11, 0, "SAME"],
      ])
    );
    expect(runs).toHaveLength(2);
    expect(runs.map((r) => r.cellCount)).toEqual([2, 2]);
  });

  it("does not merge across columns when rows differ", () => {
    const runs = findRuns(
      cells([
        [0, 0, "SAME"],
        [1, 1, "SAME"], // diagonal, not contiguous
      ])
    );
    expect(runs).toHaveLength(2);
  });

  it("keeps different signatures in separate runs even when adjacent", () => {
    const runs = findRuns(
      cells([
        [0, 0, "A"],
        [0, 1, "B"],
      ])
    );
    expect(runs).toHaveLength(2);
  });

  it("covers every input cell exactly once", () => {
    const spec: Array<[number, number, string]> = [];
    // Irregular L-shape plus outliers.
    for (let row = 0; row < 6; row++) spec.push([row, 0, "S"]);
    for (let col = 1; col < 4; col++) spec.push([5, col, "S"]);
    spec.push([2, 3, "S"]);
    const runs = findRuns(cells(spec));
    const covered = new Set<string>();
    for (const run of runs) {
      for (let row = run.startRow; row <= run.endRow; row++) {
        for (let col = run.startCol; col <= run.endCol; col++) {
          const key = `${row},${col}`;
          expect(covered.has(key), `cell ${key} covered twice`).toBe(false);
          covered.add(key);
        }
      }
    }
    expect(covered.size).toBe(spec.length);
  });

  it("handles an empty input", () => {
    expect(findRuns([])).toEqual([]);
  });

  it("scales to a 20k-cell column without quadratic blowup", () => {
    const spec: Array<[number, number, string]> = [];
    for (let row = 0; row < 20_000; row++) spec.push([row, 0, "RC[-1]*2"]);
    const started = Date.now();
    const runs = findRuns(cells(spec));
    expect(runs).toHaveLength(1);
    expect(Date.now() - started).toBeLessThan(2000);
  });
});

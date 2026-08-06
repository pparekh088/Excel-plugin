import { describe, expect, it } from "vitest";
import { planChunks } from "../src/excel/chunks";

function coverage(chunks: ReturnType<typeof planChunks>): Set<string> {
  const cells = new Set<string>();
  for (const chunk of chunks) {
    for (let r = chunk.rowOffset; r < chunk.rowOffset + chunk.rowCount; r++) {
      for (let c = chunk.colOffset; c < chunk.colOffset + chunk.colCount; c++) {
        const key = `${r},${c}`;
        expect(cells.has(key), `cell ${key} covered twice`).toBe(false);
        cells.add(key);
      }
    }
  }
  return cells;
}

describe("planChunks", () => {
  it("returns a single chunk when the range fits the budget", () => {
    expect(planChunks(10, 4, 10_000)).toEqual([
      { rowOffset: 0, rowCount: 10, colOffset: 0, colCount: 4 },
    ]);
  });

  it("splits tall ranges into row bands within the cell budget", () => {
    const chunks = planChunks(25_000, 3, 10_000);
    for (const chunk of chunks) {
      expect(chunk.rowCount * chunk.colCount).toBeLessThanOrEqual(10_000);
    }
    expect(coverage(chunks).size).toBe(75_000);
  });

  it("handles rows wider than the budget by splitting columns too", () => {
    const chunks = planChunks(2, 16_384, 10_000);
    for (const chunk of chunks) {
      expect(chunk.rowCount * chunk.colCount).toBeLessThanOrEqual(10_000);
    }
    expect(coverage(chunks).size).toBe(2 * 16_384);
  });

  it("covers exact multiples without a trailing empty chunk", () => {
    const chunks = planChunks(20_000, 1, 10_000);
    expect(chunks).toHaveLength(2);
    expect(coverage(chunks).size).toBe(20_000);
  });

  it("covers a 200k-cell worst case exactly once", () => {
    const chunks = planChunks(200, 1_000, 10_000);
    expect(coverage(chunks).size).toBe(200_000);
  });

  it("rejects degenerate dimensions", () => {
    expect(() => planChunks(0, 5)).toThrow();
    expect(() => planChunks(5, 0)).toThrow();
    expect(() => planChunks(5, 5, 0)).toThrow();
  });
});

/**
 * Precedent-agreement verification (handoff §5 acceptance).
 *
 * The stated gate compares our graph against Office.js `getDirectPrecedents`
 * on 100 random formula cells (>= 99% agreement). That comparison needs a
 * live Excel host, so it runs on the sideload checklist and is recorded in
 * the gate report.
 *
 * What we CAN verify headlessly — and what actually protects the invariant —
 * is that run collapsing and reference expansion never change the answer:
 * a naive per-cell reference implementation (parse each cell, resolve its own
 * refs, no collapsing, no expansion) must agree with the optimized graph on
 * every sampled cell. Any disagreement is a bug in the optimization, which is
 * exactly the failure mode the Excel comparison would catch.
 */

import { describe, expect, it } from "vitest";
import { DependencyGraph } from "../../src/graph/graph";
import { CellIndex } from "../../src/graph/cellIndex";
import { resolveExtracted } from "../../src/graph/resolve";
import { Workbook, a1, fullAddress } from "../../src/model/workbook";
import { parseFormula } from "../../src/parser/parser";
import { extractRefs } from "../../src/parser/refs";
import {
  budgetVsActual,
  cohortAnalysis,
  dcfModel,
  threeStatementModel,
} from "../../src/corpus/models";
import { buildCorpus } from "../../src/corpus";

/** Reference implementation: direct precedents of ONE cell, no optimizations. */
function naiveDirectPrecedents(
  workbook: Workbook,
  sheetName: string,
  row: number,
  col: number
): Set<string> {
  const sheet = workbook.sheet(sheetName);
  const cell = sheet?.get(row, col);
  const out = new Set<string>();
  if (!sheet || !cell?.formula) return out;

  const parse = parseFormula(cell.formula);
  const resolution = resolveExtracted(extractRefs(parse.ast), {
    workbook,
    hostSheet: sheetName,
    hostRow: row,
  });
  for (const area of resolution.areas) {
    const target = workbook.sheet(area.sheet);
    if (!target) continue;
    const index = new CellIndex(target);
    for (const [pRow, pCol] of index.cellsIn(
      area.startRow,
      area.endRow,
      area.startCol,
      area.endCol
    )) {
      out.add(fullAddress(target.name, pRow, pCol));
    }
  }
  return out;
}

/**
 * Precedent cells the optimized graph attributes to a given cell.
 *
 * The graph answers at NODE granularity, so a cell's precedent set is the
 * union of the cells covered by its precedent nodes, plus — when the node is
 * a cascading fill — the cells of the node itself (`readsOwnRange`).
 */
function graphPrecedentCells(
  graph: DependencyGraph,
  sheetName: string,
  row: number,
  col: number
): Set<string> {
  const node = graph.nodeAt(sheetName, row, col);
  const out = new Set<string>();
  if (!node) return out;
  if (node.readsOwnRange) {
    for (const ref of graph.nodeCells(node)) {
      out.add(fullAddress(ref.sheet, ref.row, ref.col));
    }
  }
  for (const id of graph.precedents(node.id)) {
    for (const ref of graph.nodeCells(graph.node(id)!)) {
      out.add(fullAddress(ref.sheet, ref.row, ref.col));
    }
  }
  return out;
}

/** Precedent nodes, as address strings, for spurious-dependency checks. */
function graphPrecedentNodes(
  graph: DependencyGraph,
  sheetName: string,
  row: number,
  col: number
): Array<{ address: string; cells: Set<string> }> {
  const node = graph.nodeAt(sheetName, row, col);
  if (!node) return [];
  return graph.precedents(node.id).map((id) => {
    const precedent = graph.node(id)!;
    const cells = new Set<string>();
    for (const ref of graph.nodeCells(precedent)) {
      cells.add(fullAddress(ref.sheet, ref.row, ref.col));
    }
    return { address: graph.nodeAddress(precedent), cells };
  });
}

/** Deterministic sample of formula cells across a workbook. */
function sampleFormulaCells(
  workbook: Workbook,
  count: number,
  seed = 20260806
): Array<{ sheet: string; row: number; col: number }> {
  const all: Array<{ sheet: string; row: number; col: number }> = [];
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.formula !== undefined) {
        all.push({ sheet: sheet.name, row: cell.row, col: cell.col });
      }
    }
  }
  all.sort((a, b) => a.sheet.localeCompare(b.sheet) || a.row - b.row || a.col - b.col);

  let state = seed;
  const random = () => {
    state = (state * 1103515245 + 12345) & 0x7fffffff;
    return state / 0x7fffffff;
  };
  const picked: typeof all = [];
  const used = new Set<number>();
  while (picked.length < Math.min(count, all.length)) {
    const index = Math.floor(random() * all.length);
    if (used.has(index)) continue;
    used.add(index);
    picked.push(all[index]!);
  }
  return picked;
}

describe("graph agreement with a naive per-cell reference implementation", () => {
  const models = [
    ["3-statement", threeStatementModel()],
    ["DCF", dcfModel()],
    ["budget vs actual", budgetVsActual()],
    ["cohort", cohortAnalysis()],
  ] as const;

  it.each(models)("%s: every sampled cell's precedents are covered", (_name, corpus) => {
    const { workbook } = corpus;
    const graph = DependencyGraph.build(workbook);
    const samples = sampleFormulaCells(workbook, 100);
    expect(samples.length).toBeGreaterThan(20);

    let agreed = 0;
    const disagreements: string[] = [];
    for (const sample of samples) {
      const naive = naiveDirectPrecedents(workbook, sample.sheet, sample.row, sample.col);
      const optimized = graphPrecedentCells(graph, sample.sheet, sample.row, sample.col);
      // A run node reports the union for its members, so it must be a
      // superset of any single member's precedents — never miss one.
      const missing = [...naive].filter((address) => !optimized.has(address));
      if (missing.length === 0) agreed++;
      else {
        disagreements.push(
          `${sample.sheet}!${a1(sample.row, sample.col)} missing ${missing.join(", ")}`
        );
      }
    }
    expect(disagreements).toEqual([]);
    expect(agreed / samples.length).toBeGreaterThanOrEqual(0.99);
  });

  it("reports no spurious precedents: every precedent node is genuinely read", () => {
    // A precedent node may cover more cells than the queried cell reads (it is
    // a run), but it must contain at least one cell that cell actually reads —
    // otherwise the graph invented a dependency.
    for (const [, corpus] of models) {
      const { workbook } = corpus;
      const graph = DependencyGraph.build(workbook);
      for (const node of graph.nodes) {
        if (node.kind !== "formula" || node.cellCount !== 1) continue;
        const naive = naiveDirectPrecedents(workbook, node.sheet, node.startRow, node.startCol);
        for (const precedent of graphPrecedentNodes(
          graph,
          node.sheet,
          node.startRow,
          node.startCol
        )) {
          const overlaps = [...precedent.cells].some((address) => naive.has(address));
          expect(
            overlaps,
            `${graph.nodeAddress(node)} lists ${precedent.address} but never reads it`
          ).toBe(true);
        }
      }
    }
  });

  it("flags cascading runs as reading their own range", () => {
    // C1:E1 each read the cell to their left, so the run overlaps itself.
    const workbook = new Workbook("cascade");
    const sheet = workbook.addSheet("S");
    sheet.set({ row: 0, col: 1, value: 1 });
    for (let col = 2; col <= 4; col++) {
      sheet.set({ row: 0, col, value: 0, formula: `=${a1(0, col - 1)}+1` });
    }
    const graph = DependencyGraph.build(workbook);
    const run = graph.nodeAt("S", 0, 3)!;
    expect(graph.nodeAddress(run)).toBe("S!C1:E1");
    expect(run.readsOwnRange).toBe(true);
    expect(graph.stats.cycleCount).toBe(0);
  });

  it("holds across every corpus workbook including broken variants", () => {
    for (const corpus of buildCorpus()) {
      const graph = DependencyGraph.build(corpus.workbook);
      for (const sample of sampleFormulaCells(corpus.workbook, 40)) {
        const naive = naiveDirectPrecedents(
          corpus.workbook,
          sample.sheet,
          sample.row,
          sample.col
        );
        const optimized = graphPrecedentCells(graph, sample.sheet, sample.row, sample.col);
        const missing = [...naive].filter((address) => !optimized.has(address));
        expect(
          missing,
          `${corpus.id} ${sample.sheet}!${a1(sample.row, sample.col)}`
        ).toEqual([]);
      }
    }
  });
});

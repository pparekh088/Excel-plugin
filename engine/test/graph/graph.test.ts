import { describe, expect, it } from "vitest";
import { DependencyGraph } from "../../src/graph/graph";
import { Workbook } from "../../src/model/workbook";
import { fillDown, workbookOf } from "../helpers/build";

function graphOf(sheets: Parameters<typeof workbookOf>[0]): DependencyGraph {
  return DependencyGraph.build(workbookOf(sheets));
}

/** Node covering a cell, as "Sheet!A1:B2". */
function addressAt(graph: DependencyGraph, sheet: string, row: number, col: number): string {
  const node = graph.nodeAt(sheet, row, col);
  if (!node) throw new Error(`no node at ${sheet} r${row}c${col}`);
  return graph.nodeAddress(node);
}

describe("DependencyGraph construction", () => {
  it("collapses a filled column into a single node", () => {
    const graph = graphOf({
      Model: {
        A1: 100,
        ...fillDown((row) => `=A${row}*1.05`, 1, 50, "B"),
      },
    });
    expect(addressAt(graph, "Model", 0, 1)).toBe("Model!B1:B50");
    expect(graph.stats.formulaCells).toBe(50);
    expect(graph.stats.runCount).toBe(1);
  });

  it("links formulas to the constants they read", () => {
    const graph = graphOf({ S: { A1: 10, B1: "=A1*2" } });
    const input = graph.nodeAt("S", 0, 0)!;
    const formula = graph.nodeAt("S", 0, 1)!;
    expect(input.kind).toBe("value");
    expect(graph.dependents(input.id)).toContain(formula.id);
    expect(graph.precedents(formula.id)).toContain(input.id);
  });

  it("resolves cross-sheet references", () => {
    const graph = graphOf({
      Inputs: { B2: 0.05 },
      Model: { A1: "=Inputs!B2*100" },
    });
    const source = graph.nodeAt("Inputs", 1, 1)!;
    const target = graph.nodeAt("Model", 0, 0)!;
    expect(graph.dependents(source.id)).toContain(target.id);
  });

  it("resolves quoted sheet names", () => {
    const graph = graphOf({
      "My Inputs": { A1: 7 },
      Model: { A1: "='My Inputs'!A1+1" },
    });
    expect(graph.precedents(graph.nodeAt("Model", 0, 0)!.id)).toHaveLength(1);
  });

  it("only links populated cells inside a large range", () => {
    const graph = graphOf({
      S: { A1: 1, A5: 2, A9: 3, C1: "=SUM(A:A)" },
    });
    const total = graph.nodeAt("S", 0, 2)!;
    expect(graph.precedents(total.id)).toHaveLength(3);
  });

  it("traverses a multi-step chain", () => {
    const graph = graphOf({
      S: { A1: 1, B1: "=A1*2", C1: "=B1+1", D1: "=C1*3" },
    });
    const input = graph.nodeAt("S", 0, 0)!;
    const impacted = graph.impact(input.id).map((id) => graph.nodeAddress(graph.node(id)!));
    expect(impacted).toEqual(["S!B1", "S!C1", "S!D1"]);

    const output = graph.nodeAt("S", 0, 3)!;
    const traced = graph.trace(output.id).map((id) => graph.nodeAddress(graph.node(id)!));
    expect(traced).toEqual(["S!C1", "S!B1", "S!A1"]);
  });

  it("limits traversal depth when asked", () => {
    const graph = graphOf({
      S: { A1: 1, B1: "=A1*2", C1: "=B1+1", D1: "=C1*3" },
    });
    const input = graph.nodeAt("S", 0, 0)!;
    expect(graph.impact(input.id, 1)).toHaveLength(1);
    expect(graph.impact(input.id, 2)).toHaveLength(2);
  });

  it("computes the blast radius of an arbitrary area", () => {
    const graph = graphOf({
      S: { A1: 1, A2: 2, B1: "=A1*2", B2: "=A2*2", C1: "=B1+B2" },
    });
    const impacted = graph
      .impactOfArea("S", 0, 0, 1, 0)
      .map((id) => graph.nodeAddress(graph.node(id)!))
      .sort();
    // B1 and B2 share the fill signature RC[-1]*2, so they are one run node.
    expect(impacted).toEqual(["S!B1:B2", "S!C1"]);
  });

  it("resolves defined names", () => {
    const workbook = workbookOf({ Assumptions: { B2: 0.05 }, Model: { A1: "=Growth*100" } });
    workbook.names.push({ name: "Growth", scope: null, refersTo: "=Assumptions!$B$2" });
    const graph = DependencyGraph.build(workbook);
    const source = graph.nodeAt("Assumptions", 1, 1)!;
    expect(graph.dependents(source.id)).toHaveLength(1);
    expect(graph.stats.unresolvedRefs).toBe(0);
  });

  it("resolves a name that refers to another name", () => {
    const workbook = workbookOf({ S: { A1: 3, B1: "=Outer" } });
    workbook.names.push({ name: "Inner", scope: null, refersTo: "=S!$A$1" });
    workbook.names.push({ name: "Outer", scope: null, refersTo: "=Inner" });
    const graph = DependencyGraph.build(workbook);
    expect(graph.precedents(graph.nodeAt("S", 0, 1)!.id)).toHaveLength(1);
  });

  it("counts unresolvable names instead of silently dropping them", () => {
    const graph = graphOf({ S: { A1: "=MissingName*2" } });
    expect(graph.stats.unresolvedRefs).toBe(1);
    expect(graph.nodeAt("S", 0, 0)!.opaque).toBe(true);
  });

  it("expands 3D references across the sheet span", () => {
    const graph = graphOf({
      Jan: { B2: 1 },
      Feb: { B2: 2 },
      Mar: { B2: 3 },
      Summary: { A1: "=SUM(Jan:Mar!B2)" },
    });
    const total = graph.nodeAt("Summary", 0, 0)!;
    expect(graph.precedents(total.id)).toHaveLength(3);
  });

  it("marks external workbook references unresolved but keeps the node", () => {
    const graph = graphOf({ S: { A1: "=[Other.xlsx]Sheet1!A1*2" } });
    expect(graph.stats.externalRefNodes).toBe(1);
    expect(graph.stats.unresolvedRefs).toBe(1);
    expect(graph.nodeAt("S", 0, 0)).toBeDefined();
  });

  it("flags opaque formulas and still records their literal refs", () => {
    const graph = graphOf({ S: { A1: 5, B1: "=SUM(OFFSET(A1,0,0,3,1))" } });
    const node = graph.nodeAt("S", 0, 1)!;
    expect(node.opaque).toBe(true);
    expect(node.volatile).toBe(true);
    expect(graph.precedents(node.id)).toHaveLength(1);
    expect(graph.stats.opaqueNodes).toBe(1);
  });

  it("counts error cells", () => {
    const graph = graphOf({ S: { A1: "#REF!", B1: 1 } });
    expect(graph.stats.errorCells).toBe(1);
  });

  it("keeps unparsed formulas out of neighbouring runs", () => {
    const graph = graphOf({
      S: { A1: 1, A2: 2, B1: "=A1*2", B2: "=A2*2", B3: "=SUM(A3" },
    });
    expect(graph.stats.unparsedFormulas).toBe(1);
    expect(graph.nodeAt("S", 2, 1)!.unparsed).toBe(true);
    expect(addressAt(graph, "S", 0, 1)).toBe("S!B1:B2");
  });

  describe("collapsed runs carry every member cell's precedents", () => {
    it("expands relative references across a filled row", () => {
      // C1:F1 are one run; each reads the driver directly above it.
      const graph = graphOf({
        S: {
          C1: "=C5*2",
          D1: "=D5*2",
          E1: "=E5*2",
          F1: "=F5*2",
          C5: 1,
          D5: 2,
          E5: 3,
          F5: 4,
        },
      });
      expect(addressAt(graph, "S", 0, 2)).toBe("S!C1:F1");
      // Every driver must show the run as a dependent, not just the anchor's.
      for (const col of [2, 3, 4, 5]) {
        const driver = graph.nodeAt("S", 4, col)!;
        expect(graph.dependents(driver.id), `driver at col ${col}`).toHaveLength(1);
      }
    });

    it("expands relative references across a filled column", () => {
      const graph = graphOf({
        S: { A1: 1, A2: 2, A3: 3, B1: "=A1*2", B2: "=A2*2", B3: "=A3*2" },
      });
      expect(addressAt(graph, "S", 0, 1)).toBe("S!B1:B3");
      for (const row of [0, 1, 2]) {
        expect(graph.dependents(graph.nodeAt("S", row, 0)!.id)).toHaveLength(1);
      }
    });

    it("does not expand absolute references", () => {
      // Every cell reads the SAME pinned driver; nothing else should be linked.
      const graph = graphOf({
        S: { A1: 5, A2: 99, A3: 99, B1: "=$A$1*2", B2: "=$A$1*2", B3: "=$A$1*2" },
      });
      expect(addressAt(graph, "S", 0, 1)).toBe("S!B1:B3");
      expect(graph.dependents(graph.nodeAt("S", 0, 0)!.id)).toHaveLength(1);
      // A2/A3 are never referenced, so they get no node at all.
      expect(graph.nodeAt("S", 1, 0)).toBeUndefined();
      expect(graph.nodeAt("S", 2, 0)).toBeUndefined();
    });

    it("expands the relative half of a mixed anchored range", () => {
      // Classic running total: =SUM(B$1:B3) filled down.
      const graph = graphOf({
        S: { B1: 1, B2: 2, B3: 3, C1: "=SUM(B$1:B1)", C2: "=SUM(B$1:B2)", C3: "=SUM(B$1:B3)" },
      });
      expect(addressAt(graph, "S", 0, 2)).toBe("S!C1:C3");
      for (const row of [0, 1, 2]) {
        expect(graph.dependents(graph.nodeAt("S", row, 1)!.id)).toHaveLength(1);
      }
    });

    it("expands cross-sheet references across a run", () => {
      const graph = graphOf({
        Drivers: { B1: 0.1, C1: 0.2, D1: 0.3 },
        Model: { B2: "=Drivers!B1", C2: "=Drivers!C1", D2: "=Drivers!D1" },
      });
      expect(addressAt(graph, "Model", 1, 1)).toBe("Model!B2:D2");
      for (const col of [1, 2, 3]) {
        expect(graph.dependents(graph.nodeAt("Drivers", 0, col)!.id)).toHaveLength(1);
      }
    });
  });

  it("does not collapse cells that merely share formula text without the fill pattern", () => {
    // Both cells literally reference A1, so their R1C1 signatures differ.
    const graph = graphOf({ S: { A1: 1, B1: "=A1*2", B2: "=A1*2" } });
    expect(addressAt(graph, "S", 0, 1)).toBe("S!B1");
    expect(addressAt(graph, "S", 1, 1)).toBe("S!B2");
  });
});

describe("structured references", () => {
  function tableWorkbook(): Workbook {
    const workbook = workbookOf({
      Data: {
        A1: "Item",
        B1: "Qty",
        C1: "Price",
        D1: "Total",
        A2: "x",
        B2: 2,
        C2: 10,
        D2: "=Sales[@Qty]*Sales[@Price]",
        A3: "y",
        B3: 3,
        C3: 20,
        D3: "=Sales[@Qty]*Sales[@Price]",
        F1: "=SUM(Sales[Total])",
      },
    });
    workbook.tables.push({
      name: "Sales",
      sheet: "Data",
      headerRow: 0,
      startRow: 0,
      endRow: 2,
      startCol: 0,
      endCol: 3,
      hasTotals: false,
      columns: [
        { name: "Item", col: 0 },
        { name: "Qty", col: 1 },
        { name: "Price", col: 2 },
        { name: "Total", col: 3 },
      ],
    });
    return workbook;
  }

  it("resolves [@Col] across the rows a filled run covers", () => {
    const graph = DependencyGraph.build(tableWorkbook());
    // D2:D3 share a signature and collapse, so the run reads both rows' Qty
    // and Price cells.
    const totalColumn = graph.nodeAt("Data", 1, 3)!;
    expect(graph.nodeAddress(totalColumn)).toBe("Data!D2:D3");
    const precedents = graph
      .precedents(totalColumn.id)
      .map((id) => graph.nodeAddress(graph.node(id)!))
      .sort();
    expect(precedents).toEqual(["Data!B2", "Data!B3", "Data!C2", "Data!C3"]);
  });

  it("keeps [@Col] inside the table body for a single-row formula", () => {
    const workbook = tableWorkbook();
    // Break the fill so D3 is its own node, then check it reads only row 3.
    workbook.sheet("Data")!.set({ row: 2, col: 3, value: 0, formula: "=Sales[@Qty]*2" });
    const graph = DependencyGraph.build(workbook);
    const single = graph.nodeAt("Data", 2, 3)!;
    expect(graph.nodeAddress(single)).toBe("Data!D3");
    expect(
      graph.precedents(single.id).map((id) => graph.nodeAddress(graph.node(id)!))
    ).toEqual(["Data!B3"]);
  });

  it("resolves a whole-column structured ref to the data body", () => {
    const graph = DependencyGraph.build(tableWorkbook());
    const total = graph.nodeAt("Data", 0, 5)!;
    const precedents = graph.precedents(total.id).map((id) => graph.nodeAddress(graph.node(id)!));
    // D2:D3 collapse into one run node.
    expect(precedents).toEqual(["Data!D2:D3"]);
  });

  it("reports an unknown table as unresolved", () => {
    const graph = graphOf({ S: { A1: "=SUM(NoSuchTable[Amount])" } });
    expect(graph.stats.unresolvedRefs).toBe(1);
  });
});

describe("cycle detection", () => {
  it("finds a direct self-reference", () => {
    const graph = graphOf({ S: { A1: "=A1+1" } });
    expect(graph.stats.cycleCount).toBe(1);
    expect(graph.circularGroups[0]).toHaveLength(1);
  });

  it("finds a two-node cycle", () => {
    const graph = graphOf({ S: { A1: "=B1+1", B1: "=A1+1" } });
    expect(graph.stats.cycleCount).toBe(1);
    expect(graph.circularGroups[0]).toHaveLength(2);
  });

  it("finds a longer cycle across sheets", () => {
    const graph = graphOf({
      A: { A1: "=B!A1" },
      B: { A1: "=C!A1" },
      C: { A1: "=A!A1" },
    });
    expect(graph.stats.cycleCount).toBe(1);
    expect(graph.circularGroups[0]).toHaveLength(3);
  });

  it("reports no cycles for a clean model", () => {
    const graph = graphOf({
      S: { A1: 1, B1: "=A1", C1: "=B1", D1: "=C1+A1" },
    });
    expect(graph.stats.cycleCount).toBe(0);
  });

  it("does not flag a diamond dependency as circular", () => {
    const graph = graphOf({
      S: { A1: 1, B1: "=A1", C1: "=A1", D1: "=B1+C1" },
    });
    expect(graph.stats.cycleCount).toBe(0);
  });

  it("handles a deep chain without stack overflow", () => {
    const cells: Record<string, string | number> = { A1: 1 };
    for (let row = 2; row <= 5000; row++) cells[`A${row}`] = `=A${row - 1}+1`;
    const graph = graphOf({ S: cells });
    expect(graph.stats.cycleCount).toBe(0);
    // Each row differs from its neighbours only by fill, so they collapse.
    expect(graph.stats.runCount).toBeLessThan(5);
  });

  it("finds two independent cycles", () => {
    const graph = graphOf({
      S: { A1: "=A2", A2: "=A1", C1: "=C2", C2: "=C1" },
    });
    expect(graph.stats.cycleCount).toBe(2);
  });
});

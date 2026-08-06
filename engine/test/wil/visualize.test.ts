import { describe, expect, it } from "vitest";
import { DependencyGraph } from "../../src/graph/graph";
import { renderSvg, visualizeAround } from "../../src/wil/visualize";
import { workbookOf } from "../helpers/build";
import { threeStatementModel } from "../../src/corpus";

describe("dependency visualizer", () => {
  it("lays out a chain left to right, focus in the middle", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=A1*2", C1: "=B1+1", D1: "=C1*3" },
    });
    const graph = DependencyGraph.build(workbook);
    const visual = visualizeAround(graph, "S", 0, 2); // focus on C1

    const focus = visual.nodes.find((node) => node.kind === "focus")!;
    const input = visual.nodes.find((node) => node.label === "S!A1")!;
    const output = visual.nodes.find((node) => node.label === "S!D1")!;

    expect(input.layer).toBeLessThan(focus.layer);
    expect(output.layer).toBeGreaterThan(focus.layer);
  });

  it("draws a 200-cell filled column as ONE box", () => {
    // One shared driver so the picture is a run node plus its input, which is
    // what makes a real model's graph legible at all.
    const cells: Record<string, string | number> = { A1: 5 };
    for (let row = 1; row <= 200; row++) cells[`B${row}`] = "=$A$1*2";
    const graph = DependencyGraph.build(workbookOf({ S: cells }));
    const visual = visualizeAround(graph, "S", 0, 1);

    const run = visual.nodes.find((node) => node.label === "S!B1:B200");
    expect(run).toBeDefined();
    expect(run!.cellCount).toBe(200);
    // The whole picture is the run plus the driver it reads.
    expect(visual.nodes).toHaveLength(2);
  });

  it("caps the node count so the picture stays readable", () => {
    const { workbook } = threeStatementModel();
    const graph = DependencyGraph.build(workbook);
    const visual = visualizeAround(graph, "BS", 8, 1, { depth: 10, maxNodes: 12 });
    expect(visual.nodes.length).toBeLessThanOrEqual(12);
  });

  it("marks cross-sheet edges", () => {
    const workbook = workbookOf({
      Inputs: { B2: 0.05 },
      Model: { A1: "=Inputs!B2*100" },
    });
    const graph = DependencyGraph.build(workbook);
    const visual = visualizeAround(graph, "Model", 0, 0);
    expect(visual.edges.some((edge) => edge.crossSheet)).toBe(true);
  });

  it("says when opaque references make the picture incomplete", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: '=INDIRECT("A1")*2', C1: "=B1+1" },
    });
    const graph = DependencyGraph.build(workbook);
    const visual = visualizeAround(graph, "S", 0, 2);
    expect(visual.caveat).toContain("lower bound");
  });

  it("returns an empty graph for a cell with no node", () => {
    const graph = DependencyGraph.build(workbookOf({ S: { A1: 1 } }));
    const visual = visualizeAround(graph, "S", 50, 50);
    expect(visual.nodes).toEqual([]);
  });

  it("renders valid SVG", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1*2", C1: "=B1+1" } });
    const graph = DependencyGraph.build(workbook);
    const svg = renderSvg(visualizeAround(graph, "S", 0, 1));
    expect(svg.startsWith("<svg")).toBe(true);
    expect(svg.endsWith("</svg>")).toBe(true);
    expect(svg).toContain("<rect");
    expect(svg).toContain("S!B1");
  });

  it("escapes sheet names that would break the SVG", () => {
    const workbook = workbookOf({ "A&B<C>": { A1: 1, B1: "=A1*2" } });
    const graph = DependencyGraph.build(workbook);
    const svg = renderSvg(visualizeAround(graph, "A&B<C>", 0, 1));
    expect(svg).not.toMatch(/<text[^>]*>[^<]*&(?!amp;|lt;|gt;|quot;)/);
    expect(svg).toContain("&amp;");
  });

  it("renders a placeholder rather than empty SVG when there is nothing to show", () => {
    const svg = renderSvg({ nodes: [], edges: [], layerCount: 0, truncated: 0 });
    expect(svg).toContain("No dependencies");
  });
});

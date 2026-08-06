/**
 * Dependency visualizer (handoff §10, Phase 4).
 *
 * Produces a layered graph layout the task pane renders as SVG. Works on RUN
 * NODES, not cells — a 10k-row column is one box, which is the only way a
 * real model's graph is legible at all.
 *
 * Scoped to a focus cell and a depth, because a whole-workbook graph is a
 * hairball nobody can read. The layering is longest-path so upstream inputs
 * sit on the left and outputs on the right, matching how modellers draw it.
 */

import { DependencyGraph, GraphNode } from "../graph/graph";

export interface VisualNode {
  id: number;
  label: string;
  sheet: string;
  /** Column in the layered layout; 0 is furthest upstream. */
  layer: number;
  /** Position within the layer. */
  index: number;
  kind: "input" | "calculation" | "output" | "focus";
  cellCount: number;
  formula?: string;
  opaque: boolean;
  hasError: boolean;
}

export interface VisualEdge {
  from: number;
  to: number;
  /** True when the edge crosses sheets — worth drawing differently. */
  crossSheet: boolean;
}

export interface VisualGraph {
  nodes: VisualNode[];
  edges: VisualEdge[];
  layerCount: number;
  /** Nodes omitted because the depth or node cap was reached. */
  truncated: number;
  /** Stated when opaque nodes are in view: the picture is incomplete. */
  caveat?: string;
}

export interface VisualizeOptions {
  /** How far upstream and downstream to walk. */
  depth?: number;
  /** Hard cap on nodes, so the picture stays readable. */
  maxNodes?: number;
}

export function visualizeAround(
  graph: DependencyGraph,
  sheet: string,
  row: number,
  col: number,
  options: VisualizeOptions = {}
): VisualGraph {
  const depth = options.depth ?? 3;
  const maxNodes = options.maxNodes ?? 60;

  const focus = graph.nodeAt(sheet, row, col);
  if (!focus) {
    return { nodes: [], edges: [], layerCount: 0, truncated: 0 };
  }

  const upstream = graph.trace(focus.id, depth);
  const downstream = graph.impact(focus.id, depth);
  const all = [focus.id, ...upstream, ...downstream];

  const included = new Set(all.slice(0, maxNodes));
  const truncated = all.length - included.size;

  // Layer assignment: distance from the furthest upstream node in view.
  const layerOf = new Map<number, number>();
  const upstreamSet = new Set(upstream);

  // Walk upstream levels outward from the focus.
  let frontier = [focus.id];
  layerOf.set(focus.id, 0);
  for (let step = 1; step <= depth; step++) {
    const next: number[] = [];
    for (const id of frontier) {
      for (const precedent of graph.precedents(id)) {
        if (!included.has(precedent) || layerOf.has(precedent)) continue;
        layerOf.set(precedent, -step);
        next.push(precedent);
      }
    }
    frontier = next;
    if (frontier.length === 0) break;
  }

  frontier = [focus.id];
  for (let step = 1; step <= depth; step++) {
    const next: number[] = [];
    for (const id of frontier) {
      for (const dependent of graph.dependents(id)) {
        if (!included.has(dependent) || layerOf.has(dependent)) continue;
        layerOf.set(dependent, step);
        next.push(dependent);
      }
    }
    frontier = next;
    if (frontier.length === 0) break;
  }

  // Normalize so the leftmost layer is 0.
  const minLayer = Math.min(...layerOf.values());
  const perLayerCounts = new Map<number, number>();
  const nodes: VisualNode[] = [];

  for (const id of included) {
    const node = graph.node(id);
    if (!node) continue;
    const layer = (layerOf.get(id) ?? 0) - minLayer;
    const index = perLayerCounts.get(layer) ?? 0;
    perLayerCounts.set(layer, index + 1);

    nodes.push({
      id,
      label: graph.nodeAddress(node),
      sheet: node.sheet,
      layer,
      index,
      kind: classify(graph, node, id === focus.id, upstreamSet.has(id)),
      cellCount: node.cellCount,
      formula: node.formula,
      opaque: node.opaque,
      hasError: node.hasError,
    });
  }

  const edges: VisualEdge[] = [];
  for (const id of included) {
    for (const dependent of graph.dependents(id)) {
      if (!included.has(dependent)) continue;
      const from = graph.node(id);
      const to = graph.node(dependent);
      edges.push({
        from: id,
        to: dependent,
        crossSheet: (from?.sheet ?? "") !== (to?.sheet ?? ""),
      });
    }
  }

  const opaqueInView = nodes.filter((node) => node.opaque).length;
  return {
    nodes: nodes.sort((a, b) => a.layer - b.layer || a.index - b.index),
    edges,
    layerCount: perLayerCounts.size,
    truncated,
    caveat:
      opaqueInView > 0
        ? `${opaqueInView} block(s) in view use INDIRECT, OFFSET or external links, so some ` +
          `dependencies are not drawn — this picture is a lower bound.`
        : undefined,
  };
}

function classify(
  graph: DependencyGraph,
  node: GraphNode,
  isFocus: boolean,
  isUpstream: boolean
): VisualNode["kind"] {
  if (isFocus) return "focus";
  if (node.kind === "value") return "input";
  if (graph.dependents(node.id).length === 0) return "output";
  return isUpstream ? "calculation" : "calculation";
}

/**
 * Render to standalone SVG. Kept in the engine (not the React layer) so the
 * layout is unit-testable and identical wherever it is drawn.
 */
export function renderSvg(visual: VisualGraph, width = 720): string {
  if (visual.nodes.length === 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="60"><text x="8" y="32" font-size="12">No dependencies to show.</text></svg>`;
  }

  const boxWidth = 130;
  const boxHeight = 34;
  const gapX = 60;
  const gapY = 16;
  const perLayer = new Map<number, number>();
  for (const node of visual.nodes) {
    perLayer.set(node.layer, (perLayer.get(node.layer) ?? 0) + 1);
  }
  const tallestLayer = Math.max(...perLayer.values());
  const height = Math.max(80, tallestLayer * (boxHeight + gapY) + 40);
  const totalWidth = Math.max(width, visual.layerCount * (boxWidth + gapX));

  const position = (node: VisualNode): { x: number; y: number } => ({
    x: 20 + node.layer * (boxWidth + gapX),
    y: 20 + node.index * (boxHeight + gapY),
  });

  const colors: Record<VisualNode["kind"], string> = {
    input: "#DCEEFB",
    calculation: "#F3F2F1",
    output: "#DFF6DD",
    focus: "#FFF1C2",
  };

  const byId = new Map(visual.nodes.map((node) => [node.id, node]));
  const parts: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${totalWidth}" height="${height}" font-family="Segoe UI, sans-serif">`,
  ];

  for (const edge of visual.edges) {
    const from = byId.get(edge.from);
    const to = byId.get(edge.to);
    if (!from || !to) continue;
    const start = position(from);
    const end = position(to);
    parts.push(
      `<line x1="${start.x + boxWidth}" y1="${start.y + boxHeight / 2}" ` +
        `x2="${end.x}" y2="${end.y + boxHeight / 2}" stroke="#8A8886" stroke-width="1"` +
        (edge.crossSheet ? ` stroke-dasharray="4 3"` : "") +
        ` />`
    );
  }

  for (const node of visual.nodes) {
    const { x, y } = position(node);
    const stroke = node.hasError ? "#A4262C" : node.opaque ? "#986F0B" : "#C8C6C4";
    parts.push(
      `<rect x="${x}" y="${y}" width="${boxWidth}" height="${boxHeight}" rx="4" ` +
        `fill="${colors[node.kind]}" stroke="${stroke}" stroke-width="${node.kind === "focus" ? 2 : 1}" />`
    );
    const label = node.label.length > 20 ? `${node.label.slice(0, 19)}…` : node.label;
    parts.push(
      `<text x="${x + 8}" y="${y + 15}" font-size="11" fill="#201F1E">${escapeXml(label)}</text>`
    );
    parts.push(
      `<text x="${x + 8}" y="${y + 28}" font-size="9" fill="#605E5C">` +
        `${node.cellCount} cell${node.cellCount === 1 ? "" : "s"}` +
        `${node.opaque ? " · opaque" : ""}</text>`
    );
  }

  parts.push("</svg>");
  return parts.join("");
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

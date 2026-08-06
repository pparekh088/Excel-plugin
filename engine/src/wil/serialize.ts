/**
 * WIL serializer (handoff §5.5) — the only workbook representation the LLM
 * ever sees (INV-6). Never a raw grid: a hierarchical outline of sheets ->
 * regions -> named inputs/outputs -> key dependency chains, plus workbook
 * statistics, fitted to a token budget (default 6k).
 *
 * Budgeting is greedy by importance: stats and sheet headers first, then
 * regions ranked by how much they explain (inputs and outputs before filler),
 * then key chains. Whatever is dropped is reported in `truncated` so the
 * agent knows to request detail with targeted range.read calls rather than
 * assuming it saw everything.
 */

import { DependencyGraph, GraphNode } from "../graph/graph";
import { Workbook, a1Range } from "../model/workbook";
import { Region, SemanticMap, buildSemanticMap, regionAddress } from "./semantic";

export interface WilOptions {
  /** Approximate token budget for the rendered text. */
  tokenBudget?: number;
  /** Max regions listed per sheet before summarizing the rest. */
  maxRegionsPerSheet?: number;
  /** Max key chains to trace. */
  maxChains?: number;
}

export interface WilStats {
  sheets: number;
  cells: number;
  formulaCells: number;
  runNodes: number;
  edges: number;
  errorCells: number;
  volatileNodes: number;
  opaqueNodes: number;
  unresolvedRefs: number;
  unparsedFormulas: number;
  circularGroups: number;
  externalRefNodes: number;
  tables: number;
  definedNames: number;
  charts: number;
  pivots: number;
  olapPivots: number;
}

export interface Wil {
  text: string;
  stats: WilStats;
  semanticMap: SemanticMap;
  /** True when the budget forced content to be dropped. */
  truncated: boolean;
  approxTokens: number;
}

/**
 * Token estimate. Deliberately conservative (~3.6 chars/token rather than 4):
 * overshooting the budget is a real failure, undershooting is not.
 */
export function estimateTokens(text: string): number {
  return Math.ceil(text.length / 3.6);
}

function statsOf(workbook: Workbook, graph: DependencyGraph): WilStats {
  return {
    sheets: workbook.sheets.length,
    cells: workbook.cellCount,
    formulaCells: graph.stats.formulaCells,
    runNodes: graph.stats.runCount,
    edges: graph.stats.edgeCount,
    errorCells: graph.stats.errorCells,
    volatileNodes: graph.stats.volatileNodes,
    opaqueNodes: graph.stats.opaqueNodes,
    unresolvedRefs: graph.stats.unresolvedRefs,
    unparsedFormulas: graph.stats.unparsedFormulas,
    circularGroups: graph.stats.cycleCount,
    externalRefNodes: graph.stats.externalRefNodes,
    tables: workbook.tables.length,
    definedNames: workbook.names.length,
    charts: workbook.charts.length,
    pivots: workbook.pivots.length,
    olapPivots: workbook.pivots.filter((p) => p.olap).length,
  };
}

/** Regions ranked by explanatory value, so truncation drops filler first. */
function regionPriority(region: Region): number {
  switch (region.kind) {
    case "output":
      return 0;
    case "input":
      return 1;
    case "timeAxis":
      return 2;
    case "calculation":
      return 3;
    case "label":
      return 5;
    case "constant":
      return 6;
    case "unknown":
      return 4;
  }
}

function describeRegion(region: Region): string {
  const address = `${a1Range(region.startRow, region.startCol, region.endRow, region.endCol)}`;
  const parts = [`  ${address} ${region.kind}`];
  if (region.formula) parts.push(`formula=${region.formula}`);
  if (region.labels && region.labels.length > 0) {
    parts.push(`labels=[${region.labels.slice(0, 4).join(", ")}]`);
  }
  parts.push(`(${region.cellCount} cell${region.cellCount === 1 ? "" : "s"})`);
  return parts.join(" ");
}

/**
 * Key chains: the longest dependency paths ending at outputs. These are what
 * a modeller would trace by hand, and what the agent needs to reason about
 * before editing anything.
 */
function keyChains(graph: DependencyGraph, limit: number): string[] {
  const outputs = graph.nodes.filter(
    (node) => node.kind === "formula" && graph.dependents(node.id).length === 0
  );

  const scored = outputs.map((output) => {
    const upstream = graph.trace(output.id);
    return { output, depth: upstream.length };
  });
  scored.sort((a, b) => b.depth - a.depth);

  const chains: string[] = [];
  for (const { output, depth } of scored.slice(0, limit)) {
    if (depth === 0) continue;
    chains.push(`  ${describeChain(graph, output)} (${depth} upstream nodes)`);
  }
  return chains;
}

/** Walk one representative path upward, preferring the deepest branch. */
function describeChain(graph: DependencyGraph, output: GraphNode): string {
  const path: string[] = [graph.nodeAddress(output)];
  const seen = new Set<number>([output.id]);
  let current = output;
  for (let step = 0; step < 8; step++) {
    const precedents = graph.precedents(current.id).filter((id) => !seen.has(id));
    if (precedents.length === 0) break;
    let best = precedents[0]!;
    let bestDepth = -1;
    for (const id of precedents.slice(0, 12)) {
      const depth = graph.trace(id, 6).length;
      if (depth > bestDepth) {
        bestDepth = depth;
        best = id;
      }
    }
    seen.add(best);
    const node = graph.node(best)!;
    path.push(graph.nodeAddress(node));
    current = node;
  }
  return path.join(" <- ");
}

export function buildWil(
  workbook: Workbook,
  graph: DependencyGraph,
  options: WilOptions = {}
): Wil {
  const budget = options.tokenBudget ?? 6000;
  const maxRegions = options.maxRegionsPerSheet ?? 40;
  const maxChains = options.maxChains ?? 8;
  const semanticMap = buildSemanticMap(workbook, graph);
  const stats = statsOf(workbook, graph);

  const header: string[] = [];
  header.push(`WORKBOOK ${workbook.name}`);
  header.push(
    `stats: ${stats.sheets} sheets, ${stats.cells} cells, ${stats.formulaCells} formulas ` +
      `collapsed into ${stats.runNodes} run nodes, ${stats.edges} edges`
  );
  header.push(
    `structure: ${stats.tables} tables, ${stats.definedNames} names, ` +
      `${stats.charts} charts, ${stats.pivots} pivots` +
      (stats.olapPivots > 0 ? ` (${stats.olapPivots} OLAP — not inspectable via API)` : "")
  );

  // Coverage caveats belong up front: the agent must not assume completeness.
  const caveats: string[] = [];
  if (stats.errorCells > 0) caveats.push(`${stats.errorCells} error cells`);
  if (stats.circularGroups > 0) caveats.push(`${stats.circularGroups} circular groups`);
  if (stats.opaqueNodes > 0) {
    caveats.push(`${stats.opaqueNodes} opaque nodes (INDIRECT/OFFSET/computed refs)`);
  }
  if (stats.unresolvedRefs > 0) caveats.push(`${stats.unresolvedRefs} unresolved references`);
  if (stats.unparsedFormulas > 0) caveats.push(`${stats.unparsedFormulas} unparsed formulas`);
  if (stats.externalRefNodes > 0) {
    caveats.push(`${stats.externalRefNodes} nodes referencing external workbooks`);
  }
  if (stats.volatileNodes > 0) caveats.push(`${stats.volatileNodes} volatile nodes`);
  header.push(
    caveats.length > 0
      ? `COVERAGE CAVEATS: ${caveats.join("; ")}. Dependency coverage is NOT complete ` +
          `where opaque or unresolved references exist — verify with targeted reads.`
      : `coverage: full (no opaque, unresolved, or unparsed formulas)`
  );

  const sections: string[] = [];
  let truncated = false;

  for (const sheet of workbook.sheets) {
    const sheetRegions = semanticMap.regions.filter(
      (region) => region.sheet.toUpperCase() === sheet.name.toUpperCase()
    );
    const bounds = sheet.bounds();
    const extent = bounds
      ? a1Range(bounds.startRow, bounds.startCol, bounds.endRow, bounds.endCol)
      : "(empty)";
    const lines: string[] = [
      `SHEET ${sheet.name}${sheet.visible ? "" : " [hidden]"}` +
        `${sheet.protectedSheet ? " [protected]" : ""} extent=${extent} cells=${sheet.cells.size}`,
    ];

    const ranked = [...sheetRegions].sort(
      (a, b) => regionPriority(a) - regionPriority(b) || b.cellCount - a.cellCount
    );
    for (const region of ranked.slice(0, maxRegions)) {
      lines.push(describeRegion(region));
    }
    if (ranked.length > maxRegions) {
      truncated = true;
      lines.push(`  ... ${ranked.length - maxRegions} further regions omitted`);
    }
    sections.push(lines.join("\n"));
  }

  if (workbook.tables.length > 0) {
    const lines = ["TABLES"];
    for (const table of workbook.tables) {
      lines.push(
        `  ${table.name} on ${table.sheet}!` +
          `${a1Range(table.startRow, table.startCol, table.endRow, table.endCol)} ` +
          `columns=[${table.columns.map((c) => c.name).join(", ")}]`
      );
    }
    sections.push(lines.join("\n"));
  }

  if (workbook.names.length > 0) {
    const lines = ["DEFINED NAMES"];
    for (const name of workbook.names.slice(0, 60)) {
      lines.push(`  ${name.name}${name.scope ? ` [${name.scope}]` : ""} -> ${name.refersTo}`);
    }
    if (workbook.names.length > 60) {
      truncated = true;
      lines.push(`  ... ${workbook.names.length - 60} more`);
    }
    sections.push(lines.join("\n"));
  }

  const chains = keyChains(graph, maxChains);
  if (chains.length > 0) {
    sections.push(["KEY CHAINS (output <- ... <- input)", ...chains].join("\n"));
  }

  // Assemble greedily against the budget.
  const parts = [header.join("\n")];
  let used = estimateTokens(parts[0]!);
  for (const section of sections) {
    const cost = estimateTokens(section) + 1;
    if (used + cost > budget) {
      truncated = true;
      continue;
    }
    parts.push(section);
    used += cost;
  }
  if (truncated) {
    const notice =
      "NOTE: this summary was truncated to fit the context budget. " +
      "Use range.read / graph.trace for anything not shown.";
    parts.push(notice);
    used += estimateTokens(notice);
  }

  const text = parts.join("\n\n");
  return { text, stats, semanticMap, truncated, approxTokens: estimateTokens(text) };
}

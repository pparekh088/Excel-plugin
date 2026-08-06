/**
 * Dependency graph over range-run nodes (INV-4).
 *
 * Built entirely from our own parse of the formulas — never getPrecedents().
 * Nodes are formula runs (contiguous cells sharing an R1C1 signature) and the
 * constant cells those formulas reference. Edges point precedent -> dependent
 * so impact analysis is a forward traversal.
 *
 * Coverage honesty: formulas containing INDIRECT/OFFSET or computed range
 * endpoints are marked opaque, and their unresolvable references are counted.
 * `stats.opaqueNodes` and `stats.unresolvedRefs` are surfaced in the WIL and
 * audit output so we never present the graph as complete when it is not.
 */

import { Cell, Sheet, Workbook, a1Range, fullAddress } from "../model/workbook";
import { ParseResult } from "../parser/ast";
import { isFormulaText, parseFormula } from "../parser/parser";
import { normalizeR1C1 } from "../parser/r1c1";
import { ExtractedRefs, extractRefs } from "../parser/refs";
import { CellIndex } from "./cellIndex";
import { ResolvedArea, resolveExtracted } from "./resolve";
import { Run, findRuns } from "./runs";

export type NodeKind = "formula" | "value";

export interface GraphNode {
  id: number;
  kind: NodeKind;
  sheet: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  cellCount: number;
  /** R1C1 signature (formula nodes only). */
  signature?: string;
  /** Representative formula text of the run's anchor cell. */
  formula?: string;
  /** Value of the anchor cell (value nodes, and last-computed for formulas). */
  value?: Cell["value"];
  opaque: boolean;
  volatile: boolean;
  hasError: boolean;
  functions: string[];
  /** Parse failed — formula text preserved but not understood. */
  unparsed: boolean;
  /**
   * The run reads cells inside its own extent — a cascading fill such as
   * `=B1+1` filled right, where each cell reads its neighbour. Real and
   * acyclic per cell, so it is recorded here rather than reported as a
   * circular reference (cycle detection verifies these at cell level).
   */
  readsOwnRange: boolean;
}

export interface GraphStats {
  nodeCount: number;
  edgeCount: number;
  formulaCells: number;
  valueCells: number;
  runCount: number;
  opaqueNodes: number;
  volatileNodes: number;
  unparsedFormulas: number;
  unresolvedRefs: number;
  errorCells: number;
  externalRefNodes: number;
  cycleCount: number;
}

export interface CellRef {
  sheet: string;
  row: number;
  col: number;
}

const EXCEL_ERRORS = new Set([
  "#REF!",
  "#VALUE!",
  "#DIV/0!",
  "#NAME?",
  "#N/A",
  "#NULL!",
  "#NUM!",
  "#SPILL!",
  "#CALC!",
  "#GETTING_DATA",
  "#FIELD!",
  "#BLOCKED!",
  "#CONNECT!",
  "#BUSY!",
]);

export function isErrorValue(value: unknown): value is string {
  return typeof value === "string" && EXCEL_ERRORS.has(value.toUpperCase());
}

/** Per-formula-node analysis kept for the audit engine. */
export interface NodeAnalysis {
  parse: ParseResult;
  refs: ExtractedRefs;
  resolved: ResolvedArea[];
  unresolved: string[];
}

export class DependencyGraph {
  readonly nodes: GraphNode[] = [];
  /** precedent -> dependents. */
  private readonly outgoing = new Map<number, Set<number>>();
  /** dependent -> precedents. */
  private readonly incoming = new Map<number, Set<number>>();
  /** sheet -> "row,col" -> node id. */
  private readonly cellToNode = new Map<string, Map<string, number>>();
  private readonly analyses = new Map<number, NodeAnalysis>();
  private readonly indexes = new Map<string, CellIndex>();
  /** Run nodes whose expanded references overlap themselves — verified later. */
  private readonly selfEdgeCandidates = new Set<number>();
  private cycles: number[][] = [];

  readonly stats: GraphStats = {
    nodeCount: 0,
    edgeCount: 0,
    formulaCells: 0,
    valueCells: 0,
    runCount: 0,
    opaqueNodes: 0,
    volatileNodes: 0,
    unparsedFormulas: 0,
    unresolvedRefs: 0,
    errorCells: 0,
    externalRefNodes: 0,
    cycleCount: 0,
  };

  private constructor(readonly workbook: Workbook) {}

  // ------------------------------------------------------------------ build

  static build(workbook: Workbook): DependencyGraph {
    const graph = new DependencyGraph(workbook);
    for (const sheet of workbook.sheets) {
      graph.indexes.set(sheet.name.toUpperCase(), new CellIndex(sheet));
      graph.cellToNode.set(sheet.name.toUpperCase(), new Map());
    }
    graph.buildFormulaNodes();
    graph.buildEdges();
    graph.detectCycles();
    graph.stats.nodeCount = graph.nodes.length;
    return graph;
  }

  private buildFormulaNodes(): void {
    for (const sheet of this.workbook.sheets) {
      const signatures: Array<{ row: number; col: number; signature: string }> = [];
      const parsedByCell = new Map<string, { parse: ParseResult; refs: ExtractedRefs }>();

      for (const cell of sheet.cells.values()) {
        if (isErrorValue(cell.value)) this.stats.errorCells++;
        if (!isFormulaText(cell.formula)) continue;
        this.stats.formulaCells++;
        const parse = parseFormula(cell.formula);
        const refs = extractRefs(parse.ast);
        parsedByCell.set(`${cell.row},${cell.col}`, { parse, refs });
        // Unparsed formulas get their own signature so they never merge into
        // a neighbouring run and mask an inconsistency.
        const signature = parse.ok
          ? normalizeR1C1(parse.ast, cell.row, cell.col)
          : `<unparsed:${cell.row},${cell.col}>`;
        if (!parse.ok) this.stats.unparsedFormulas++;
        signatures.push({ row: cell.row, col: cell.col, signature });
      }

      const runs = findRuns(signatures);
      this.stats.runCount += runs.length;
      for (const run of runs) {
        this.addFormulaNode(sheet, run, parsedByCell);
      }
    }
    this.stats.valueCells = this.workbook.cellCount - this.stats.formulaCells;
  }

  private addFormulaNode(
    sheet: Sheet,
    run: Run,
    parsedByCell: Map<string, { parse: ParseResult; refs: ExtractedRefs }>
  ): void {
    const anchor = sheet.get(run.startRow, run.startCol);
    const parsed = parsedByCell.get(`${run.startRow},${run.startCol}`);
    if (!anchor || !parsed) return;

    // Resolve against the run's full extent, not just its anchor: every cell
    // in a filled run reads its own shifted references (see resolve.ts).
    const context = {
      workbook: this.workbook,
      hostSheet: sheet.name,
      hostRow: run.startRow,
      hostRowEnd: run.endRow,
      spanRows: run.endRow - run.startRow,
      spanCols: run.endCol - run.startCol,
    };
    const resolution = resolveExtracted(parsed.refs, context);
    this.stats.unresolvedRefs += resolution.unresolved.length;

    let hasError = false;
    for (let row = run.startRow; row <= run.endRow; row++) {
      for (let col = run.startCol; col <= run.endCol; col++) {
        if (isErrorValue(sheet.get(row, col)?.value)) hasError = true;
      }
    }

    const node: GraphNode = {
      id: this.nodes.length,
      kind: "formula",
      sheet: sheet.name,
      startRow: run.startRow,
      startCol: run.startCol,
      endRow: run.endRow,
      endCol: run.endCol,
      cellCount: run.cellCount,
      signature: run.signature,
      formula: anchor.formula,
      value: anchor.value,
      opaque: parsed.refs.opaque || resolution.unresolved.length > 0,
      volatile: parsed.refs.volatile,
      hasError,
      functions: parsed.refs.functions,
      unparsed: !parsed.parse.ok,
      readsOwnRange: false,
    };
    this.registerNode(node);
    this.analyses.set(node.id, {
      parse: parsed.parse,
      refs: parsed.refs,
      resolved: resolution.areas,
      unresolved: resolution.unresolved,
    });
    if (node.opaque) this.stats.opaqueNodes++;
    if (node.volatile) this.stats.volatileNodes++;
    if (parsed.refs.external) this.stats.externalRefNodes++;
  }

  private registerNode(node: GraphNode): void {
    this.nodes.push(node);
    const map = this.cellToNode.get(node.sheet.toUpperCase());
    if (!map) return;
    for (let row = node.startRow; row <= node.endRow; row++) {
      for (let col = node.startCol; col <= node.endCol; col++) {
        map.set(`${row},${col}`, node.id);
      }
    }
  }

  /** Create (or reuse) a single-cell node for a referenced constant. */
  private valueNodeAt(sheet: Sheet, row: number, col: number): number | null {
    const map = this.cellToNode.get(sheet.name.toUpperCase());
    if (!map) return null;
    const existing = map.get(`${row},${col}`);
    if (existing !== undefined) return existing;

    const cell = sheet.get(row, col);
    if (!cell) return null;
    const node: GraphNode = {
      id: this.nodes.length,
      kind: "value",
      sheet: sheet.name,
      startRow: row,
      startCol: col,
      endRow: row,
      endCol: col,
      cellCount: 1,
      value: cell.value,
      opaque: false,
      volatile: false,
      hasError: isErrorValue(cell.value),
      functions: [],
      unparsed: false,
      readsOwnRange: false,
    };
    this.registerNode(node);
    return node.id;
  }

  private buildEdges(): void {
    // Snapshot: value nodes get appended during this loop.
    const formulaNodes = this.nodes.filter((node) => node.kind === "formula");
    for (const node of formulaNodes) {
      const analysis = this.analyses.get(node.id);
      if (!analysis) continue;
      for (const area of analysis.resolved) {
        const sheet = this.workbook.sheet(area.sheet);
        if (!sheet) continue;
        const index = this.indexes.get(sheet.name.toUpperCase());
        if (!index) continue;

        for (const [row, col] of index.cellsIn(
          area.startRow,
          area.endRow,
          area.startCol,
          area.endCol
        )) {
          const precedent =
            this.cellToNode.get(sheet.name.toUpperCase())?.get(`${row},${col}`) ??
            this.valueNodeAt(sheet, row, col);
          if (precedent === null || precedent === undefined) continue;
          if (precedent === node.id) {
            // A run whose expanded references overlap itself is usually a
            // cascading fill (=A1 filled right reads its own left neighbour),
            // which is perfectly acyclic per cell. Defer to cell-level
            // verification rather than reporting a phantom circular reference.
            this.selfEdgeCandidates.add(node.id);
            node.readsOwnRange = true;
            continue;
          }
          this.addEdge(precedent, node.id);
        }
      }
    }
  }

  private addEdge(from: number, to: number): void {
    let out = this.outgoing.get(from);
    if (!out) this.outgoing.set(from, (out = new Set()));
    let inc = this.incoming.get(to);
    if (!inc) this.incoming.set(to, (inc = new Set()));
    if (!out.has(to)) this.stats.edgeCount++;
    out.add(to);
    inc.add(from);
  }

  // ------------------------------------------------------------- accessors

  node(id: number): GraphNode | undefined {
    return this.nodes[id];
  }

  analysis(id: number): NodeAnalysis | undefined {
    return this.analyses.get(id);
  }

  nodeAt(sheet: string, row: number, col: number): GraphNode | undefined {
    const id = this.cellToNode.get(sheet.toUpperCase())?.get(`${row},${col}`);
    return id === undefined ? undefined : this.nodes[id];
  }

  /** Direct dependents (cells whose formulas read this node). */
  dependents(id: number): number[] {
    return [...(this.outgoing.get(id) ?? [])];
  }

  /** Direct precedents (what this node's formula reads). */
  precedents(id: number): number[] {
    return [...(this.incoming.get(id) ?? [])];
  }

  /**
   * Transitive traversal with an optional depth limit. Iterative — model
   * chains can be thousands deep and must not blow the stack.
   */
  private traverse(startIds: number[], direction: "up" | "down", maxDepth: number): number[] {
    const edges = direction === "down" ? this.outgoing : this.incoming;
    const seen = new Set<number>(startIds);
    const result: number[] = [];
    let frontier = [...startIds];
    let depth = 0;

    while (frontier.length > 0 && depth < maxDepth) {
      const next: number[] = [];
      for (const id of frontier) {
        for (const neighbour of edges.get(id) ?? []) {
          if (seen.has(neighbour)) continue;
          seen.add(neighbour);
          result.push(neighbour);
          next.push(neighbour);
        }
      }
      frontier = next;
      depth++;
    }
    return result;
  }

  /** Everything downstream of a node — the blast radius of changing it. */
  impact(id: number, maxDepth = Infinity): number[] {
    return this.traverse([id], "down", maxDepth);
  }

  /** Everything upstream — what feeds a node. */
  trace(id: number, maxDepth = Infinity): number[] {
    return this.traverse([id], "up", maxDepth);
  }

  /** Blast radius of an arbitrary rectangle (used by change-set previews). */
  impactOfArea(
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): number[] {
    const seeds = new Set<number>();
    const map = this.cellToNode.get(sheet.toUpperCase());
    if (map) {
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const id = map.get(`${row},${col}`);
          if (id !== undefined) seeds.add(id);
        }
      }
    }
    return this.traverse([...seeds], "down", Infinity);
  }

  // -------------------------------------------------------- cycle detection

  /**
   * Tarjan's SCC, iterative. Any SCC with more than one node — or a node with
   * a self-edge — is a circular reference (AUD-005).
   */
  private detectCycles(): void {
    const index = new Map<number, number>();
    const low = new Map<number, number>();
    const onStack = new Set<number>();
    const stack: number[] = [];
    let counter = 0;
    const candidates: number[][] = [];

    for (const root of this.nodes) {
      if (index.has(root.id)) continue;
      // Explicit work stack: [node, next-neighbour-position]
      const work: Array<[number, number]> = [[root.id, 0]];
      index.set(root.id, counter);
      low.set(root.id, counter);
      counter++;
      stack.push(root.id);
      onStack.add(root.id);

      while (work.length > 0) {
        const frame = work[work.length - 1]!;
        const [current, position] = frame;
        const neighbours = [...(this.outgoing.get(current) ?? [])];

        if (position < neighbours.length) {
          frame[1]++;
          const neighbour = neighbours[position]!;
          if (!index.has(neighbour)) {
            index.set(neighbour, counter);
            low.set(neighbour, counter);
            counter++;
            stack.push(neighbour);
            onStack.add(neighbour);
            work.push([neighbour, 0]);
          } else if (onStack.has(neighbour)) {
            low.set(current, Math.min(low.get(current)!, index.get(neighbour)!));
          }
          continue;
        }

        work.pop();
        if (work.length > 0) {
          const parent = work[work.length - 1]![0];
          low.set(parent, Math.min(low.get(parent)!, low.get(current)!));
        }
        if (low.get(current) === index.get(current)) {
          const component: number[] = [];
          for (;;) {
            const popped = stack.pop()!;
            onStack.delete(popped);
            component.push(popped);
            if (popped === current) break;
          }
          if (component.length > 1) {
            candidates.push(component);
          } else if (this.selfEdgeCandidates.has(current)) {
            candidates.push([current]);
          }
        }
      }
    }

    // Node-level SCCs over collapsed runs over-approximate: verify each
    // candidate against the actual per-cell dependencies before reporting a
    // circular reference to the user.
    this.cycles = candidates.filter((component) => this.hasCellLevelCycle(component));
    this.stats.cycleCount = this.cycles.length;
  }

  /**
   * Exact check inside a candidate component: expand its nodes to cells, wire
   * up each cell's own precedents (no run expansion), and look for a real
   * cycle. Bounded by the component's size, so it stays cheap.
   */
  private hasCellLevelCycle(component: number[]): boolean {
    const inComponent = new Set(component);
    const cells: string[] = [];
    const cellPrecedents = new Map<string, string[]>();

    for (const id of component) {
      const node = this.nodes[id];
      if (!node || node.kind !== "formula") continue;
      const sheet = this.workbook.sheet(node.sheet);
      if (!sheet) continue;

      for (let row = node.startRow; row <= node.endRow; row++) {
        for (let col = node.startCol; col <= node.endCol; col++) {
          const cell = sheet.get(row, col);
          if (!cell?.formula) continue;
          const key = `${node.sheet.toUpperCase()}!${row},${col}`;
          cells.push(key);

          const parse = parseFormula(cell.formula);
          const resolution = resolveExtracted(extractRefs(parse.ast), {
            workbook: this.workbook,
            hostSheet: node.sheet,
            hostRow: row,
          });
          const precedents: string[] = [];
          for (const area of resolution.areas) {
            const target = this.workbook.sheet(area.sheet);
            const index = target
              ? this.indexes.get(target.name.toUpperCase())
              : undefined;
            if (!target || !index) continue;
            for (const [pRow, pCol] of index.cellsIn(
              area.startRow,
              area.endRow,
              area.startCol,
              area.endCol
            )) {
              const owner = this.cellToNode
                .get(target.name.toUpperCase())
                ?.get(`${pRow},${pCol}`);
              // Only edges that stay inside the candidate component matter.
              if (owner === undefined || !inComponent.has(owner)) continue;
              precedents.push(`${target.name.toUpperCase()}!${pRow},${pCol}`);
            }
          }
          cellPrecedents.set(key, precedents);
        }
      }
    }

    // Iterative DFS with colouring: grey = on the current path.
    const state = new Map<string, 0 | 1 | 2>();
    for (const start of cells) {
      if (state.get(start)) continue;
      const stack: Array<[string, number]> = [[start, 0]];
      state.set(start, 1);
      while (stack.length > 0) {
        const frame = stack[stack.length - 1]!;
        const neighbours = cellPrecedents.get(frame[0]) ?? [];
        if (frame[1] < neighbours.length) {
          const next = neighbours[frame[1]++]!;
          const colour = state.get(next);
          if (colour === 1) return true; // back edge -> real cycle
          if (colour === undefined) {
            state.set(next, 1);
            stack.push([next, 0]);
          }
          continue;
        }
        state.set(frame[0], 2);
        stack.pop();
      }
    }
    return false;
  }

  get circularGroups(): readonly number[][] {
    return this.cycles;
  }

  // ------------------------------------------------------------- reporting

  nodeAddress(node: GraphNode): string {
    return `${node.sheet}!${a1Range(node.startRow, node.startCol, node.endRow, node.endCol)}`;
  }

  cellAddress(ref: CellRef): string {
    return fullAddress(ref.sheet, ref.row, ref.col);
  }

  /** Cells covered by a node, expanded — used for highlight and snapshots. */
  *nodeCells(node: GraphNode): Generator<CellRef> {
    for (let row = node.startRow; row <= node.endRow; row++) {
      for (let col = node.startCol; col <= node.endCol; col++) {
        yield { sheet: node.sheet, row, col };
      }
    }
  }
}

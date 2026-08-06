/**
 * Deterministic audit rules AUD-001..AUD-011 (handoff §6).
 *
 * Every rule takes the graph + semantic map and returns findings. No LLM, no
 * network, no I/O. Precision is the priority: the ≥95% gate means a rule that
 * cannot be sure should either not fire or fire with lower `confidence`, and
 * false positives on clean models are counted directly against us.
 */

import { DependencyGraph, GraphNode, isErrorValue } from "../graph/graph";
import { Sheet, Workbook, a1, fullAddress } from "../model/workbook";
import { Node, walk } from "../parser/ast";
import { VOLATILE_FUNCTIONS } from "../parser/refs";
import { SemanticMap, regionOfNode } from "../wil/semantic";
import { Finding, RuleMeta } from "./types";

export interface RuleContext {
  workbook: Workbook;
  graph: DependencyGraph;
  semanticMap: SemanticMap;
}

export interface Rule extends RuleMeta {
  run(context: RuleContext): Finding[];
}

// ------------------------------------------------------------------ helpers

function blastRadius(graph: DependencyGraph, node: GraphNode): number {
  return graph.impact(node.id).reduce((sum, id) => sum + (graph.node(id)?.cellCount ?? 0), 0);
}

function traceOf(graph: DependencyGraph, node: GraphNode, limit = 40): Finding["trace"] {
  const cells = [...graph.nodeCells(node)].slice(0, limit);
  for (const id of graph.impact(node.id, 2)) {
    const dependent = graph.node(id);
    if (!dependent) continue;
    for (const cell of graph.nodeCells(dependent)) {
      if (cells.length >= limit) break;
      cells.push(cell);
    }
  }
  return cells;
}

function finding(
  partial: Omit<Finding, "address" | "sheet" | "row" | "col"> & {
    sheet: string;
    row: number;
    col: number;
  }
): Finding {
  return {
    ...partial,
    address: fullAddress(partial.sheet, partial.row, partial.col),
  };
}

/** Numeric literals in a formula, excluding ones that are structurally fine. */
function magicNumbers(ast: Node): number[] {
  const found: number[] = [];
  const benign = new Set([0, 1, -1, 2, 12, 100, 365, 366, 4, 24, 60, 1000, 10000, 1_000_000]);

  const collect = (node: Node, insideBenignContext: boolean): void => {
    switch (node.kind) {
      case "number":
        if (!insideBenignContext && !benign.has(node.value)) found.push(node.value);
        return;
      case "unary":
        // -0.05 reads as one literal.
        if (node.operand.kind === "number") {
          const value = node.op === "-" ? -node.operand.value : node.operand.value;
          if (!insideBenignContext && !benign.has(Math.abs(value))) found.push(value);
          return;
        }
        collect(node.operand, insideBenignContext);
        return;
      case "func": {
        // Index/position/rounding arguments are structural, not assumptions.
        const structural = new Set([
          "INDEX", "MATCH", "VLOOKUP", "HLOOKUP", "OFFSET", "ROUND", "ROUNDUP",
          "ROUNDDOWN", "LEFT", "RIGHT", "MID", "SUBTOTAL", "CHOOSE", "LARGE",
          "SMALL", "RANK", "POWER", "SEQUENCE", "TEXT", "REPT", "COLUMN", "ROW",
        ]);
        const isStructural = structural.has(node.name);
        node.args.forEach((arg, index) => {
          // First arg of ROUND-family is the value; later args are digits.
          const structuralArg = isStructural && (node.name.startsWith("ROUND") ? index > 0 : true);
          collect(arg, insideBenignContext || structuralArg);
        });
        return;
      }
      case "binary":
        collect(node.left, insideBenignContext);
        // The exponent of ^ is a period index in discount factors and
        // compounding (=1/(1+WACC)^3), not an assumption.
        collect(node.right, insideBenignContext || node.op === "^");
        return;
      case "group":
        collect(node.expr, insideBenignContext);
        return;
      case "percent":
        collect(node.operand, true); // 5% is self-documenting
        return;
      case "array":
        for (const row of node.rows) for (const item of row) collect(item, insideBenignContext);
        return;
      case "callExpr":
        collect(node.callee, insideBenignContext);
        for (const arg of node.args) collect(arg, insideBenignContext);
        return;
      case "implicitIntersection":
      case "spill":
        collect(node.operand, insideBenignContext);
        return;
      default:
        return;
    }
  };
  collect(ast, false);
  return found;
}

// -------------------------------------------------------------- AUD-001

/**
 * A cell whose formula breaks the R1C1 pattern of the run it sits inside.
 * Detected structurally: a small run wedged between much larger runs that
 * share a signature with each other, along the same row or column band.
 */
export const AUD_001: Rule = {
  id: "AUD-001",
  title: "Formula inconsistency",
  severity: "high",
  description: "A cell breaks the fill pattern of the formulas around it.",
  run({ graph }) {
    const findings: Finding[] = [];

    for (const suspect of graph.nodes) {
      if (suspect.kind !== "formula") continue;
      // Only small runs can be "the odd one out".
      if (suspect.cellCount > 2) continue;

      /**
       * Look at the CELLS flanking the suspect rather than at neighbouring
       * nodes with matching bounds. When other defects fragment a block into
       * irregular rectangles, the flanking cells still belong to nodes that
       * share the block's signature — bounds-matching would miss it.
       */
      const flanking = (
        rowStep: number,
        colStep: number
      ): GraphNode | undefined => {
        const row = rowStep < 0 ? suspect.startRow - 1 : rowStep > 0 ? suspect.endRow + 1 : suspect.startRow;
        const col = colStep < 0 ? suspect.startCol - 1 : colStep > 0 ? suspect.endCol + 1 : suspect.startCol;
        const node = graph.nodeAt(suspect.sheet, row, col);
        return node?.kind === "formula" ? node : undefined;
      };

      /**
       * Does the suspect's own formula recur along the perpendicular axis?
       * If so it belongs to a fill pattern running the other way and the
       * "anomaly" is coincidental. In a financial model, vertically adjacent
       * rows are different line items that can share an R1C1 signature by
       * chance (gross profit and EBITDA are both "sum of the two rows above"),
       * so without this check every fragmented row produces false positives.
       */
      const recursAlong = (axis: "row" | "column"): boolean =>
        graph.nodes.some((other) => {
          if (other === suspect || other.kind !== "formula") return false;
          if (other.sheet !== suspect.sheet) return false;
          if (other.signature !== suspect.signature) return false;
          return axis === "row"
            ? other.startCol <= suspect.endCol && other.endCol >= suspect.startCol
            : other.startRow <= suspect.endRow && other.endRow >= suspect.startRow;
        });

      const axes = [
        { axis: "row" as const, before: flanking(0, -1), after: flanking(0, 1) },
        { axis: "column" as const, before: flanking(-1, 0), after: flanking(1, 0) },
      ];

      for (const { axis, before, after } of axes) {
        // The pattern must be established on BOTH sides — a run that merely
        // ends is not an inconsistency, it is the edge of a block.
        if (!before || !after) continue;
        if (before.signature !== after.signature) continue;
        if (before.signature === suspect.signature) continue;
        // Checking down a column? Then the suspect must not recur across its row.
        if (recursAlong(axis === "column" ? "column" : "row")) continue;

        const neighbours = before === after ? [before] : [before, after];
        const surroundingCells = neighbours.reduce((sum, n) => sum + n.cellCount, 0);
        if (surroundingCells < 3) continue;

          findings.push(
            finding({
              rule: AUD_001.id,
              title: AUD_001.title,
              severity: AUD_001.severity,
              confidence: surroundingCells >= 6 ? "certain" : "likely",
              sheet: suspect.sheet,
              row: suspect.startRow,
              col: suspect.startCol,
              explanation:
                `${a1(suspect.startRow, suspect.startCol)} uses a different formula from the ` +
                `${surroundingCells} cells around it in the same ${axis}. The neighbours all ` +
                `follow one pattern; this cell does not, which usually means it was edited by ` +
                `hand after the block was filled.`,
              evidence: suspect.formula,
              trace: traceOf(graph, suspect),
              blastRadius: blastRadius(graph, suspect),
              autoFix: neighbours[0]!.formula
                ? {
                    description: `Restore the surrounding pattern (copy the formula from ${a1(
                      neighbours[0]!.startRow,
                      neighbours[0]!.startCol
                    )})`,
                    // Overwriting an existing formula is never a safe auto-fix.
                    risk: "high",
                    edits: [
                      {
                        sheet: suspect.sheet,
                        row: suspect.startRow,
                        col: suspect.startCol,
                        formula: neighbours[0]!.formula,
                      },
                    ],
                  }
                : undefined,
          })
        );
        break; // one finding per suspect
      }
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-002

export const AUD_002: Rule = {
  id: "AUD-002",
  title: "Hardcoded value inside a formula",
  severity: "medium",
  description: "A magic number is embedded in a formula instead of referencing an input cell.",
  run({ graph }) {
    const findings: Finding[] = [];
    for (const node of graph.nodes) {
      if (node.kind !== "formula") continue;
      const analysis = graph.analysis(node.id);
      if (!analysis || !analysis.parse.ok) continue;

      const numbers = magicNumbers(analysis.parse.ast);
      if (numbers.length === 0) continue;
      // A formula that is nothing but a literal (=41250) is a plugged value,
      // which AUD-003 covers. Anything with actual computation in it —
      // =1/(1+0.095)^3 — belongs here even with no cell references.
      if (isBareLiteral(analysis.parse.ast)) continue;

      findings.push(
        finding({
          rule: AUD_002.id,
          title: AUD_002.title,
          severity: AUD_002.severity,
          confidence: numbers.length > 1 ? "certain" : "likely",
          sheet: node.sheet,
          row: node.startRow,
          col: node.startCol,
          explanation:
            `${a1(node.startRow, node.startCol)} embeds the literal ` +
            `${numbers.map((n) => String(n)).join(", ")} directly in its formula. ` +
            `Assumptions buried in formulas are invisible to anyone reviewing the model and ` +
            `cannot be changed in one place — move the value to a labelled input cell.`,
          evidence: node.formula,
          trace: traceOf(graph, node),
          blastRadius: blastRadius(graph, node),
        })
      );
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-003

export const AUD_003: Rule = {
  id: "AUD-003",
  title: "Hardcoded value in a calculation chain",
  severity: "high",
  description:
    "A constant sits inside a calculation block and feeds downstream formulas, " +
    "outside any designated input region.",
  run({ graph, semanticMap, workbook }) {
    const findings: Finding[] = [];
    for (const node of graph.nodes) {
      if (node.kind !== "value") continue;
      if (typeof node.value !== "number") continue;
      const dependents = graph.dependents(node.id);
      if (dependents.length === 0) continue;

      const sheet = workbook.sheet(node.sheet);
      if (!sheet) continue;

      /**
       * A constant is a PLUG when formulas flank it on both sides of an axis:
       * something that used to be computed was typed over. A constant with
       * formulas only after it is a SEED — the year-0 units or opening
       * balance every forecast row starts from — which is correct modelling
       * and must not be flagged.
       */
      const isFormula = (row: number, col: number): boolean =>
        sheet.get(row, col)?.formula !== undefined;
      const flankedHorizontally =
        isFormula(node.startRow, node.startCol - 1) && isFormula(node.startRow, node.startCol + 1);
      const flankedVertically =
        isFormula(node.startRow - 1, node.startCol) && isFormula(node.startRow + 1, node.startCol);
      if (!flankedHorizontally && !flankedVertically) continue;

      const neighbours = formulaNeighbourCount(sheet, node.startRow, node.startCol);
      const region = regionOfNode(semanticMap, node);
      if (region?.kind === "input" && region.cellCount > 2) continue;

      findings.push(
        finding({
          rule: AUD_003.id,
          title: AUD_003.title,
          severity: AUD_003.severity,
          confidence: neighbours >= 3 ? "certain" : "likely",
          sheet: node.sheet,
          row: node.startRow,
          col: node.startCol,
          explanation:
            `${a1(node.startRow, node.startCol)} holds the constant ${node.value} but sits ` +
            `among ${neighbours} formula cells and feeds ${dependents.length} downstream ` +
            `node(s). A plugged number inside a calculation block silently overrides the ` +
            `model's logic and will not update when assumptions change.`,
          evidence: String(node.value),
          trace: traceOf(graph, node),
          blastRadius: blastRadius(graph, node),
        })
      );
    }
    return findings;
  },
};

/** `=42`, `=-42`, `=(42)` — a value typed as a formula, with no computation. */
function isBareLiteral(node: Node): boolean {
  switch (node.kind) {
    case "number":
      return true;
    case "unary":
      return isBareLiteral(node.operand);
    case "group":
      return isBareLiteral(node.expr);
    default:
      return false;
  }
}

/** Formula cells directly adjacent (4-way) to a position. */
function formulaNeighbourCount(sheet: Sheet, row: number, col: number): number {
  const offsets: Array<[number, number]> = [
    [-1, 0],
    [1, 0],
    [0, -1],
    [0, 1],
  ];
  let count = 0;
  for (const [dr, dc] of offsets) {
    if (sheet.get(row + dr, col + dc)?.formula !== undefined) count++;
  }
  return count;
}

// -------------------------------------------------------------- AUD-004

export const AUD_004: Rule = {
  id: "AUD-004",
  title: "Error cell",
  severity: "critical",
  description: "A cell evaluates to an Excel error, with its downstream blast radius.",
  run({ graph, workbook }) {
    // Collect every error cell first, then report only ROOTS. One deleted
    // precedent can poison dozens of cells; listing each one separately would
    // bury every other issue in the report under identical criticals. The
    // root's explanation carries the propagation count instead.
    const errorCells: Array<{ sheet: string; row: number; col: number; value: string }> = [];
    for (const sheet of workbook.sheets) {
      for (const cell of sheet.cells.values()) {
        if (isErrorValue(cell.value)) {
          errorCells.push({
            sheet: sheet.name,
            row: cell.row,
            col: cell.col,
            value: String(cell.value),
          });
        }
      }
    }
    if (errorCells.length === 0) return [];

    const errorKeys = new Set(
      errorCells.map((cell) => `${cell.sheet.toUpperCase()}!${cell.row},${cell.col}`)
    );
    const isErrorCell = (sheet: string, row: number, col: number): boolean =>
      errorKeys.has(`${sheet.toUpperCase()}!${row},${col}`);

    /** An error is propagated when any precedent is itself an error cell. */
    const isPropagated = (cell: { sheet: string; row: number; col: number }): boolean => {
      const node = graph.nodeAt(cell.sheet, cell.row, cell.col);
      if (!node) return false;
      return graph.precedents(node.id).some((id) => {
        const precedent = graph.node(id);
        if (!precedent) return false;
        for (const ref of graph.nodeCells(precedent)) {
          if (isErrorCell(ref.sheet, ref.row, ref.col)) return true;
        }
        return false;
      });
    };

    const roots = errorCells.filter((cell) => !isPropagated(cell));
    // Every error inside a cycle looks propagated; fall back to reporting all
    // of them rather than silently reporting nothing.
    const reportable = roots.length > 0 ? roots : errorCells;

    const findings: Finding[] = [];
    for (const cell of reportable) {
      const node = graph.nodeAt(cell.sheet, cell.row, cell.col);
      const radius = node ? blastRadius(graph, node) : 0;
      const downstreamErrors = node
        ? graph
            .impact(node.id)
            .flatMap((id) => {
              const dependent = graph.node(id);
              return dependent ? [...graph.nodeCells(dependent)] : [];
            })
            .filter((ref) => isErrorCell(ref.sheet, ref.row, ref.col)).length
        : 0;

      findings.push(
        finding({
          rule: AUD_004.id,
          title: AUD_004.title,
          severity: AUD_004.severity,
          confidence: "certain",
          sheet: cell.sheet,
          row: cell.row,
          col: cell.col,
          explanation:
            `${a1(cell.row, cell.col)} evaluates to ${cell.value}` +
            (downstreamErrors > 0
              ? `, and the error has already spread to ${downstreamErrors} other cell(s) ` +
                `downstream of it. Fixing this one cell should clear them all.`
              : radius > 0
                ? `, and ${radius} downstream cell(s) read it — the error will propagate through ` +
                  `every one of them.`
                : `. Nothing depends on it yet, so fixing it now is cheap.`),
          evidence:
            workbook.sheet(cell.sheet)?.get(cell.row, cell.col)?.formula ?? cell.value,
          trace: node
            ? traceOf(graph, node)
            : [{ sheet: cell.sheet, row: cell.row, col: cell.col }],
          blastRadius: radius,
        })
      );
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-005

export const AUD_005: Rule = {
  id: "AUD-005",
  title: "Circular reference",
  severity: "critical",
  description: "A group of cells depends on itself, verified at cell level.",
  run({ graph }) {
    const findings: Finding[] = [];
    for (const group of graph.circularGroups) {
      const nodes = group.map((id) => graph.node(id)!).filter(Boolean);
      if (nodes.length === 0) continue;
      const anchor = nodes[0]!;
      const addresses = nodes.map((node) => graph.nodeAddress(node));
      findings.push(
        finding({
          rule: AUD_005.id,
          title: AUD_005.title,
          severity: AUD_005.severity,
          confidence: "certain",
          sheet: anchor.sheet,
          row: anchor.startRow,
          col: anchor.startCol,
          explanation:
            `${addresses.join(" -> ")} form a dependency loop: each one ultimately depends on ` +
            `itself. Excel cannot resolve this without iterative calculation, so the values ` +
            `shown are not trustworthy.`,
          evidence: anchor.formula,
          trace: nodes.flatMap((node) => [...graph.nodeCells(node)]).slice(0, 40),
          blastRadius: nodes.reduce((sum, node) => sum + blastRadius(graph, node), 0),
        })
      );
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-006

export const AUD_006: Rule = {
  id: "AUD-006",
  title: "Formula references an empty cell",
  severity: "medium",
  description: "A formula reads a cell that holds nothing, which Excel silently treats as zero.",
  run({ graph, workbook }) {
    const findings: Finding[] = [];
    for (const node of graph.nodes) {
      if (node.kind !== "formula") continue;
      const analysis = graph.analysis(node.id);
      if (!analysis) continue;

      const empties: string[] = [];
      for (const area of analysis.resolved) {
        // Only single-cell references: an empty cell inside a SUM range is
        // normal and flagging it would drown the report in noise.
        if (area.startRow === null || area.startCol === null) continue;
        if (area.startRow !== area.endRow || area.startCol !== area.endCol) continue;
        const target = workbook.sheet(area.sheet);
        if (!target) continue;
        const cell = target.get(area.startRow, area.startCol);
        if (cell === undefined || cell.value === null || cell.value === "") {
          empties.push(fullAddress(area.sheet, area.startRow, area.startCol));
        }
      }
      if (empties.length === 0) continue;

      findings.push(
        finding({
          rule: AUD_006.id,
          title: AUD_006.title,
          severity: AUD_006.severity,
          confidence: "likely",
          sheet: node.sheet,
          row: node.startRow,
          col: node.startCol,
          explanation:
            `${a1(node.startRow, node.startCol)} reads ${empties.join(", ")}, which ` +
            `${empties.length === 1 ? "is" : "are"} empty. Excel treats a blank as zero, so the ` +
            `result looks valid while quietly assuming a value nobody supplied.`,
          evidence: node.formula,
          trace: traceOf(graph, node),
          blastRadius: blastRadius(graph, node),
        })
      );
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-007

export const AUD_007: Rule = {
  id: "AUD-007",
  title: "Broken calculation chain",
  severity: "high",
  description: "An output cell is no longer connected to any input.",
  run({ graph, semanticMap }) {
    const findings: Finding[] = [];
    for (const region of semanticMap.outputs) {
      const node = graph.nodeAt(region.sheet, region.startRow, region.startCol);
      if (!node || node.kind !== "formula") continue;

      const upstream = graph.trace(node.id);
      const reachesInput = upstream.some((id) => graph.node(id)?.kind === "value");
      if (reachesInput || upstream.length === 0) continue;

      findings.push(
        finding({
          rule: AUD_007.id,
          title: AUD_007.title,
          severity: AUD_007.severity,
          confidence: "possible",
          sheet: node.sheet,
          row: node.startRow,
          col: node.startCol,
          explanation:
            `${a1(node.startRow, node.startCol)} looks like an output but its dependency chain ` +
            `never reaches an input cell — it is computed entirely from other formulas with no ` +
            `underlying assumption. Check that the chain was not severed.`,
          evidence: node.formula,
          trace: traceOf(graph, node),
          blastRadius: blastRadius(graph, node),
        })
      );
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-009

export const AUD_009: Rule = {
  id: "AUD-009",
  title: "Volatile function overuse",
  severity: "low",
  description:
    "Volatile functions (NOW, RAND, OFFSET, INDIRECT) recalculate on every change and slow the model.",
  run({ graph }) {
    const volatileNodes = graph.nodes.filter((node) => node.volatile);
    if (volatileNodes.length === 0) return [];

    // One finding per location, but severity rises with the total count.
    const total = volatileNodes.reduce((sum, node) => sum + node.cellCount, 0);
    return volatileNodes.map((node) => {
      const used = node.functions.filter((fn) => VOLATILE_FUNCTIONS.has(fn));
      return finding({
        rule: AUD_009.id,
        title: AUD_009.title,
        severity: total > 50 ? "medium" : "low",
        confidence: "certain",
        sheet: node.sheet,
        row: node.startRow,
        col: node.startCol,
        explanation:
          `${graph.nodeAddress(node)} uses ${used.join(", ")}. Volatile functions recalculate ` +
          `on every edit anywhere in the workbook; this model has ${total} such cell(s). ` +
          `INDIRECT and OFFSET additionally hide their references from dependency analysis.`,
        evidence: node.formula,
        trace: [...graph.nodeCells(node)].slice(0, 40),
        blastRadius: blastRadius(graph, node),
      });
    });
  },
};

// -------------------------------------------------------------- AUD-010

export const AUD_010: Rule = {
  id: "AUD-010",
  title: "Inconsistent number formats",
  severity: "low",
  description: "Cells in one semantic region are formatted differently from their neighbours.",
  run({ workbook, semanticMap }) {
    const findings: Finding[] = [];
    for (const region of semanticMap.regions) {
      if (region.cellCount < 4) continue;
      if (region.kind === "label" || region.kind === "timeAxis") continue;
      const sheet = workbook.sheet(region.sheet);
      if (!sheet) continue;

      const counts = new Map<string, number>();
      const positions = new Map<string, Array<[number, number]>>();
      for (let row = region.startRow; row <= region.endRow; row++) {
        for (let col = region.startCol; col <= region.endCol; col++) {
          const format = sheet.get(row, col)?.numberFormat;
          if (format === undefined) continue;
          counts.set(format, (counts.get(format) ?? 0) + 1);
          let list = positions.get(format);
          if (!list) positions.set(format, (list = []));
          list.push([row, col]);
        }
      }
      if (counts.size < 2) continue;

      const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1]);
      const [dominant, dominantCount] = sorted[0]!;
      // Only flag when the majority is clear and the outliers are few.
      const outliers = sorted.slice(1);
      const outlierCount = outliers.reduce((sum, [, count]) => sum + count, 0);
      if (dominantCount < 3 || outlierCount > dominantCount / 2) continue;

      for (const [format] of outliers) {
        const where = positions.get(format)?.[0];
        if (!where) continue;
        findings.push(
          finding({
            rule: AUD_010.id,
            title: AUD_010.title,
            severity: AUD_010.severity,
            confidence: "possible",
            sheet: region.sheet,
            row: where[0],
            col: where[1],
            explanation:
              `${a1(where[0], where[1])} is formatted "${format}" while ${dominantCount} other ` +
              `cells in the same ${region.kind} block use "${dominant}". Mixed formats in one ` +
              `block usually mean a cell was pasted from elsewhere.`,
            evidence: format,
            trace: [{ sheet: region.sheet, row: where[0], col: where[1] }],
            blastRadius: 0,
            autoFix: {
              description: `Apply the block's dominant format "${dominant}"`,
              risk: "low",
              edits: [{ sheet: region.sheet, row: where[0], col: where[1], value: dominant }],
            },
          })
        );
      }
    }
    return findings;
  },
};

// -------------------------------------------------------------- AUD-011

export const AUD_011: Rule = {
  id: "AUD-011",
  title: "External or opaque reference",
  severity: "medium",
  description:
    "References into other workbooks, or through INDIRECT/OFFSET, which dependency analysis cannot follow.",
  run({ graph }) {
    const findings: Finding[] = [];
    for (const node of graph.nodes) {
      if (node.kind !== "formula") continue;
      const analysis = graph.analysis(node.id);
      if (!analysis) continue;
      const external = analysis.refs.external;
      const opaque = analysis.refs.opaque;
      if (!external && !opaque && analysis.unresolved.length === 0) continue;

      const reasons: string[] = [];
      if (external) reasons.push("links to another workbook");
      if (opaque) reasons.push("computes its references at runtime (INDIRECT/OFFSET)");
      if (analysis.unresolved.length > 0) {
        reasons.push(`references we could not resolve (${analysis.unresolved.join(", ")})`);
      }

      findings.push(
        finding({
          rule: AUD_011.id,
          title: AUD_011.title,
          severity: external ? "medium" : "low",
          confidence: "certain",
          sheet: node.sheet,
          row: node.startRow,
          col: node.startCol,
          explanation:
            `${a1(node.startRow, node.startCol)} ${reasons.join(" and ")}. Dependency analysis ` +
            `cannot see through this, so any impact assessment that includes this cell is ` +
            `incomplete — it is listed here so the gap is explicit rather than hidden.`,
          evidence: node.formula,
          trace: [...graph.nodeCells(node)].slice(0, 40),
          blastRadius: blastRadius(graph, node),
        })
      );
    }
    return findings;
  },
};

export const DETERMINISTIC_RULES: Rule[] = [
  AUD_001,
  AUD_002,
  AUD_003,
  AUD_004,
  AUD_005,
  AUD_006,
  AUD_007,
  AUD_009,
  AUD_010,
  AUD_011,
];

export { walk };

/**
 * AI-derived cell inventory (handoff §8: "Every AI.* cell tagged in WIL so the
 * audit engine can inventory AI-derived values — auditors will demand this").
 *
 * A number produced by a language model and a number produced by a formula
 * look identical once they land in a cell. Anyone signing off on a model needs
 * to know which is which, so this is a first-class part of the audit output
 * rather than a nice-to-have.
 */

import { DependencyGraph } from "../graph/graph";
import { Workbook, a1, fullAddress } from "../model/workbook";
import { Finding } from "./types";
import { Rule } from "./rules";

/** Custom functions that produce model-generated values. */
const GENERATIVE_FUNCTIONS = new Set(["AI.ASK", "AI.EXTRACT", "AI.CLASSIFY", "AI.MATCH"]);
/** AI-namespaced but statistically computed — not a generated number. */
const COMPUTED_FUNCTIONS = new Set(["AI.FORECAST"]);

export interface AiCellInventory {
  /** Cells whose value came from a language model. */
  generative: Array<{ address: string; formula: string; fn: string; dependents: number }>;
  /** Cells using AI.* functions that are statistically computed, not generated. */
  computed: Array<{ address: string; formula: string; fn: string }>;
  /** Cells currently showing the budget error. */
  budgetExhausted: string[];
  totalAiCells: number;
  /** Cells downstream of any generative AI cell — the reach of the inference. */
  downstreamCells: number;
}

function aiFunctionIn(formula: string | undefined): string | null {
  if (!formula) return null;
  const match = /\b(AI\.[A-Z]+)\s*\(/i.exec(formula);
  return match ? match[1]!.toUpperCase() : null;
}

export function inventoryAiCells(workbook: Workbook, graph: DependencyGraph): AiCellInventory {
  const inventory: AiCellInventory = {
    generative: [],
    computed: [],
    budgetExhausted: [],
    totalAiCells: 0,
    downstreamCells: 0,
  };

  const generativeNodeIds = new Set<number>();

  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) {
      if (cell.value === "#AI_BUDGET!") {
        inventory.budgetExhausted.push(fullAddress(sheet.name, cell.row, cell.col));
      }
      const fn = aiFunctionIn(cell.formula);
      if (!fn) continue;
      inventory.totalAiCells++;

      const address = fullAddress(sheet.name, cell.row, cell.col);
      const node = graph.nodeAt(sheet.name, cell.row, cell.col);

      if (COMPUTED_FUNCTIONS.has(fn)) {
        inventory.computed.push({ address, formula: cell.formula!, fn });
        continue;
      }
      if (GENERATIVE_FUNCTIONS.has(fn) || fn.startsWith("AI.")) {
        inventory.generative.push({
          address,
          formula: cell.formula!,
          fn,
          dependents: node ? graph.dependents(node.id).length : 0,
        });
        if (node) generativeNodeIds.add(node.id);
      }
    }
  }

  const downstream = new Set<number>();
  for (const id of generativeNodeIds) {
    for (const dependentId of graph.impact(id)) downstream.add(dependentId);
  }
  for (const id of downstream) {
    inventory.downstreamCells += graph.node(id)?.cellCount ?? 0;
  }

  return inventory;
}

/**
 * AUD-013 — inventory AI-derived values.
 *
 * Informational by default: using AI.* is not a defect. It becomes a real
 * finding when generated values feed downstream calculations, because at that
 * point a model's guess is propagating through the model's arithmetic.
 */
export const AUD_013: Rule = {
  id: "AUD-013",
  title: "AI-derived value",
  severity: "info",
  description:
    "A cell's value came from a language model rather than a formula. Listed so reviewers " +
    "can see which figures are inferred.",
  run({ workbook, graph }) {
    const inventory = inventoryAiCells(workbook, graph);
    const findings: Finding[] = [];

    for (const entry of inventory.generative) {
      const feedsCalculations = entry.dependents > 0;
      findings.push({
        rule: AUD_013.id,
        title: AUD_013.title,
        severity: feedsCalculations ? "medium" : "info",
        confidence: "certain",
        address: entry.address,
        sheet: entry.address.slice(0, entry.address.lastIndexOf("!")).replace(/^'|'$/g, ""),
        row: 0,
        col: 0,
        explanation:
          `${entry.address} holds a value produced by ${entry.fn}, not by a formula` +
          (feedsCalculations
            ? `, and ${entry.dependents} downstream node(s) read it — an inferred value is ` +
              `propagating into calculated results. Check it before relying on those numbers.`
            : `. Nothing depends on it, so it is presentational rather than structural.`),
        evidence: entry.formula,
        trace: [],
        blastRadius: entry.dependents,
      });
    }

    for (const address of inventory.budgetExhausted) {
      findings.push({
        rule: AUD_013.id,
        title: "AI budget exhausted",
        severity: "high",
        confidence: "certain",
        address,
        sheet: address.slice(0, address.lastIndexOf("!")).replace(/^'|'$/g, ""),
        row: 0,
        col: 0,
        explanation:
          `${address} shows #AI_BUDGET!: the per-recalculation AI call budget ran out before ` +
          `this cell was evaluated, so it holds no answer at all. Recalculate to start a new ` +
          `cycle, or reduce the number of distinct values being classified.`,
        evidence: "#AI_BUDGET!",
        trace: [],
        blastRadius: 0,
      });
    }

    return findings;
  },
};

/** Fill in the real row/col on findings produced above. */
export function resolveAiFindingPositions(
  findings: Finding[],
  workbook: Workbook
): Finding[] {
  return findings.map((finding) => {
    if (finding.rule !== AUD_013.id) return finding;
    const bang = finding.address.lastIndexOf("!");
    const sheetName = finding.address.slice(0, bang).replace(/^'|'$/g, "");
    const ref = finding.address.slice(bang + 1);
    const match = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!match) return finding;
    let col = 0;
    for (const ch of match[1]!) col = col * 26 + (ch.charCodeAt(0) - 64);
    col -= 1;
    const row = Number(match[2]) - 1;
    void workbook;
    return {
      ...finding,
      sheet: sheetName,
      row,
      col,
      trace: [{ sheet: sheetName, row, col }],
    };
  });
}

export { a1 };

/**
 * AUD-008 — balance assertions.
 *
 * Unlike the other rules this one is about *values*, so it runs post-recalc.
 * Assertions come from the user, a template, or auto-detection: a labelled
 * "check" row whose cells should be zero is the near-universal convention in
 * three-statement models, and finding those automatically is what makes the
 * rule useful without configuration.
 */

import { DependencyGraph } from "../graph/graph";
import { Workbook, a1, fullAddress, parseA1Range } from "../model/workbook";
import { BalanceAssertion, Finding } from "./types";

const CHECK_LABEL = /\b(check|tie[- ]?out|balance(s)?|must be (zero|0)|diff(erence)?)\b/i;

/**
 * Auto-detect check rows: a text label matching CHECK_LABEL with a row of
 * formula cells beside it that all evaluate to (near) zero or are expected to.
 */
export function detectBalanceAssertions(workbook: Workbook): BalanceAssertion[] {
  const assertions: BalanceAssertion[] = [];
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells.values()) {
      if (typeof cell.value !== "string" || cell.formula !== undefined) continue;
      if (!CHECK_LABEL.test(cell.value)) continue;

      // Scan right from the label for formula cells on the same row.
      for (let col = cell.col + 1; col <= cell.col + 40; col++) {
        const candidate = sheet.get(cell.row, col);
        if (!candidate) continue;
        if (candidate.formula === undefined) continue;
        assertions.push({
          name: `${cell.value.trim()} (${sheet.name})`,
          sheet: sheet.name,
          a1: a1(cell.row, col),
          expected: 0,
          tolerance: 1e-6,
        });
      }
    }
  }
  return assertions;
}

export function checkBalanceAssertions(
  workbook: Workbook,
  graph: DependencyGraph,
  assertions: BalanceAssertion[]
): Finding[] {
  const findings: Finding[] = [];
  for (const assertion of assertions) {
    const parsed = parseA1Range(assertion.a1);
    if (!parsed) continue;
    const sheet = workbook.sheet(assertion.sheet);
    const cell = sheet?.get(parsed.startRow, parsed.startCol);
    if (!sheet || !cell) continue;

    const expected = assertion.expected ?? 0;
    const tolerance = assertion.tolerance ?? 1e-6;
    const value = cell.value;

    if (typeof value !== "number") {
      // A non-numeric check cell is itself a problem worth reporting.
      findings.push({
        rule: "AUD-008",
        title: "Balance assertion could not be evaluated",
        severity: "high",
        confidence: "certain",
        address: fullAddress(sheet.name, parsed.startRow, parsed.startCol),
        sheet: sheet.name,
        row: parsed.startRow,
        col: parsed.startCol,
        explanation:
          `${assertion.name}: the check cell ${assertion.a1} holds ${JSON.stringify(value)} ` +
          `instead of a number, so the assertion cannot be evaluated.`,
        evidence: cell.formula ?? String(value),
        trace: [{ sheet: sheet.name, row: parsed.startRow, col: parsed.startCol }],
        blastRadius: 0,
      });
      continue;
    }

    const delta = value - expected;
    if (Math.abs(delta) <= tolerance) continue;

    const node = graph.nodeAt(sheet.name, parsed.startRow, parsed.startCol);
    findings.push({
      rule: "AUD-008",
      title: "Balance assertion failed",
      severity: "critical",
      confidence: "certain",
      address: fullAddress(sheet.name, parsed.startRow, parsed.startCol),
      sheet: sheet.name,
      row: parsed.startRow,
      col: parsed.startCol,
      explanation:
        `${assertion.name}: ${assertion.a1} is ${value} but should be ${expected} ` +
        `(off by ${delta}). A failing tie-out means the statements disagree with each other, ` +
        `so every figure downstream of this point is suspect.`,
      evidence: cell.formula ?? String(value),
      trace: node
        ? [...graph.nodeCells(node), ...graph.trace(node.id, 2).flatMap((id) => {
            const upstream = graph.node(id);
            return upstream ? [...graph.nodeCells(upstream)] : [];
          })].slice(0, 40)
        : [{ sheet: sheet.name, row: parsed.startRow, col: parsed.startCol }],
      blastRadius: node
        ? graph.impact(node.id).reduce((sum, id) => sum + (graph.node(id)?.cellCount ?? 0), 0)
        : 0,
    });
  }
  return findings;
}

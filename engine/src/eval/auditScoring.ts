/**
 * Audit-detection scoring.
 *
 * Naive address matching is wrong for an audit engine, in three ways that all
 * showed up on the first eval run:
 *
 *  1. One injected defect legitimately trips several rules. Replacing a
 *     formula with a magic number (AUD-002) also breaks the fill pattern of
 *     the row it sits in (AUD-001). Both findings are correct.
 *  2. Errors propagate. Injecting one #REF! makes every downstream cell an
 *     error cell, and the engine SHOULD report them all — a report that hid
 *     the blast radius would be worse.
 *  3. Findings anchor at run or group granularity: a cycle finding anchors at
 *     one member, and a volatile-function finding at a collapsed run's first
 *     cell.
 *
 * So a finding counts against precision only when it is unrelated to any
 * injected defect: not at an injection site, not downstream of one, and not
 * covering one. On CLEAN workbooks there are no injections, so every finding
 * is a false positive — which is where the precision gate really bites.
 */

import { DependencyGraph } from "../graph/graph";
import { InjectedDefect } from "../corpus/models";
import { Workbook, fullAddress, parseA1Range } from "../model/workbook";
import { Finding } from "../audit/types";

export interface AuditScore {
  matched: number;
  expectedTotal: number;
  falsePositives: number;
  relatedFindings: number;
  findingsTotal: number;
  precision: number;
  recall: number;
  f1: number;
  falsePositiveDetail: string[];
  missedDetail: string[];
}

function normalize(address: string): string {
  const bang = address.lastIndexOf("!");
  if (bang < 0) return address.toUpperCase();
  const sheet = address.slice(0, bang).replace(/^'|'$/g, "");
  return `${sheet.toUpperCase()}!${address.slice(bang + 1).toUpperCase()}`;
}

/** Every rule this finding represents, including ones folded in by dedupe. */
function rulesOf(finding: Finding): Set<string> {
  return new Set([finding.rule, ...(finding.alsoFlaggedBy ?? []).map((item) => item.rule)]);
}

/** Addresses a finding legitimately stands for: its anchor plus its node. */
function findingAddresses(finding: Finding, graph: DependencyGraph): Set<string> {
  const out = new Set<string>([normalize(finding.address)]);
  const node = graph.nodeAt(finding.sheet, finding.row, finding.col);
  if (node) {
    for (const cell of graph.nodeCells(node)) {
      out.add(normalize(fullAddress(cell.sheet, cell.row, cell.col)));
    }
  }
  return out;
}

/** Injection site plus everything downstream of it — legitimate collateral. */
function defectFootprint(
  defect: InjectedDefect,
  workbook: Workbook,
  graph: DependencyGraph
): Set<string> {
  const out = new Set<string>();
  const addresses = [
    defect.address,
    ...(defect.alsoAcceptable ?? []),
    ...(defect.collateral ?? []),
  ];

  for (const address of addresses) {
    out.add(normalize(address));
    const bang = address.lastIndexOf("!");
    if (bang < 0) continue;
    const sheetName = address.slice(0, bang).replace(/^'|'$/g, "");
    const parsed = parseA1Range(address.slice(bang + 1));
    if (!parsed) continue;
    const sheet = workbook.sheet(sheetName);
    if (!sheet) continue;

    const node = graph.nodeAt(sheet.name, parsed.startRow, parsed.startCol);
    if (!node) continue;
    // The node itself (a run stands for all its cells) ...
    for (const cell of graph.nodeCells(node)) {
      out.add(normalize(fullAddress(cell.sheet, cell.row, cell.col)));
    }
    // ... and everything downstream, which is where errors and broken
    // assertions legitimately surface.
    for (const id of graph.impact(node.id)) {
      const downstream = graph.node(id);
      if (!downstream) continue;
      for (const cell of graph.nodeCells(downstream)) {
        out.add(normalize(fullAddress(cell.sheet, cell.row, cell.col)));
      }
    }
  }
  return out;
}

export function scoreAudit(
  findings: Finding[],
  defects: InjectedDefect[],
  workbook: Workbook,
  graph: DependencyGraph
): AuditScore {
  const footprints = defects.map((defect) => ({
    defect,
    acceptable: new Set(
      [defect.address, ...(defect.alsoAcceptable ?? [])].map(normalize)
    ),
    footprint: defectFootprint(defect, workbook, graph),
  }));

  const missedDetail: string[] = [];
  let matched = 0;
  for (const entry of footprints) {
    // A defect is found when a finding of the right rule lands on one of its
    // acceptable addresses (or a run covering one).
    const hit = findings.some((finding) => {
      // Findings are deduplicated by address, so a rule may have been folded
      // into a more severe finding at the same cell — still a detection.
      if (!rulesOf(finding).has(entry.defect.rule)) return false;
      const addresses = findingAddresses(finding, graph);
      for (const address of addresses) {
        if (entry.acceptable.has(address)) return true;
      }
      return false;
    });
    if (hit) matched++;
    else missedDetail.push(`${entry.defect.rule}@${entry.defect.address}`);
  }

  const falsePositiveDetail: string[] = [];
  let related = 0;
  for (const finding of findings) {
    const addresses = findingAddresses(finding, graph);
    const isRelated = footprints.some((entry) => {
      for (const address of addresses) {
        if (entry.footprint.has(address)) return true;
      }
      return false;
    });
    if (isRelated) related++;
    else falsePositiveDetail.push(`${finding.rule}@${finding.address}`);
  }

  const falsePositives = falsePositiveDetail.length;
  const precision =
    findings.length === 0 ? 1 : (findings.length - falsePositives) / findings.length;
  const recall = defects.length === 0 ? 1 : matched / defects.length;
  const f1 = precision + recall === 0 ? 0 : (2 * precision * recall) / (precision + recall);

  return {
    matched,
    expectedTotal: defects.length,
    falsePositives,
    relatedFindings: related,
    findingsTotal: findings.length,
    precision,
    recall,
    f1,
    falsePositiveDetail,
    missedDetail,
  };
}

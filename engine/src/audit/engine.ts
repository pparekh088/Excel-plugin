/**
 * Audit engine (handoff §6) — runs every deterministic rule, ranks findings,
 * computes a health score, and reports coverage honestly.
 *
 * ZERO LLM calls on the deterministic path. `report.stats.llmCallsUsed` is
 * always 0 unless judgment rules were explicitly opted into, which is what
 * lets the whole thing be demoed with the network to any model provider off.
 */

import { DependencyGraph } from "../graph/graph";
import { Workbook } from "../model/workbook";
import { SemanticMap, buildSemanticMap } from "../wil/semantic";
import { checkBalanceAssertions, detectBalanceAssertions } from "./balance";
import { DETERMINISTIC_RULES, Rule, RuleContext } from "./rules";
import {
  AuditReport,
  AuditScope,
  BalanceAssertion,
  Finding,
  HealthScore,
  SEVERITY_WEIGHT,
  Severity,
} from "./types";

export interface AuditOptions extends AuditScope {
  /** User- or template-supplied assertions, added to auto-detected ones. */
  assertions?: BalanceAssertion[];
  /** Skip auto-detection of "check" rows. */
  skipAssertionDetection?: boolean;
  /** Reuse an existing graph/semantic map instead of rebuilding. */
  graph?: DependencyGraph;
  semanticMap?: SemanticMap;
}

function severityRank(severity: Severity): number {
  return ["critical", "high", "medium", "low", "info"].indexOf(severity);
}

/**
 * Health score: start at 100 and subtract a penalty per finding, weighted by
 * severity and scaled by blast radius (a broken cell nothing reads is not the
 * same as one feeding half the model). Deliberately saturating, so a workbook
 * with 400 low-severity findings does not read as worse than one with a
 * broken balance sheet.
 */
export function computeHealth(findings: Finding[]): HealthScore {
  const breakdown = new Map<Severity, { count: number; penalty: number }>();
  let total = 0;

  for (const finding of findings) {
    const base = SEVERITY_WEIGHT[finding.severity];
    // Blast radius scales the penalty from 1x to 2x, saturating quickly.
    const scale = 1 + Math.min(1, Math.log10(1 + finding.blastRadius) / 3);
    const penalty = base * scale;
    total += penalty;
    const entry = breakdown.get(finding.severity) ?? { count: 0, penalty: 0 };
    entry.count++;
    entry.penalty += penalty;
    breakdown.set(finding.severity, entry);
  }

  // Diminishing returns: 100 * exp(-total/120) keeps the score in (0,100].
  const score = findings.length === 0 ? 100 : Math.round(100 * Math.exp(-total / 120));
  const band: HealthScore["band"] =
    score >= 85 ? "healthy" : score >= 65 ? "needs-attention" : score >= 40 ? "at-risk" : "critical";

  return {
    score,
    band,
    breakdown: [...breakdown.entries()]
      .map(([severity, entry]) => ({ severity, ...entry }))
      .sort((a, b) => severityRank(a.severity) - severityRank(b.severity)),
  };
}

/**
 * One row per cell. When several rules fire on the same address they describe
 * one defect from different angles, so the most severe becomes the finding and
 * the rest are attached to it. Input must already be sorted by severity.
 */
function dedupeByAddress(findings: Finding[]): Finding[] {
  const primary = new Map<string, Finding>();
  for (const finding of findings) {
    const key = finding.address.toUpperCase();
    const existing = primary.get(key);
    if (!existing) {
      primary.set(key, finding);
      continue;
    }
    existing.alsoFlaggedBy ??= [];
    existing.alsoFlaggedBy.push({
      rule: finding.rule,
      title: finding.title,
      severity: finding.severity,
    });
    // Keep an auto-fix if the primary finding has none to offer.
    if (!existing.autoFix && finding.autoFix) existing.autoFix = finding.autoFix;
  }
  return [...primary.values()];
}

export function runAudit(workbook: Workbook, options: AuditOptions = {}): AuditReport {
  const started = Date.now();
  const graph = options.graph ?? DependencyGraph.build(workbook);
  const semanticMap = options.semanticMap ?? buildSemanticMap(workbook, graph);

  const requested = options.rules;
  const rules: Rule[] = DETERMINISTIC_RULES.filter(
    (rule) => !requested || requested.includes(rule.id)
  );

  const context: RuleContext = { workbook, graph, semanticMap };
  let findings: Finding[] = [];
  const rulesRun: string[] = [];

  for (const rule of rules) {
    findings.push(...rule.run(context));
    rulesRun.push(rule.id);
  }

  // AUD-008 runs separately: it is value-based, not structure-based.
  if (!requested || requested.includes("AUD-008")) {
    const assertions = [
      ...(options.skipAssertionDetection ? [] : detectBalanceAssertions(workbook)),
      ...(options.assertions ?? []),
    ];
    if (assertions.length > 0) {
      findings.push(...checkBalanceAssertions(workbook, graph, assertions));
    }
    rulesRun.push("AUD-008");
  }

  if (options.sheets && options.sheets.length > 0) {
    const wanted = new Set(options.sheets.map((name) => name.toUpperCase()));
    findings = findings.filter((finding) => wanted.has(finding.sheet.toUpperCase()));
  }

  // Rank: severity first, then blast radius, then position — so the top of
  // the list is always the thing most worth fixing.
  findings.sort(
    (a, b) =>
      severityRank(a.severity) - severityRank(b.severity) ||
      b.blastRadius - a.blastRadius ||
      a.sheet.localeCompare(b.sheet) ||
      a.row - b.row ||
      a.col - b.col
  );

  findings = dedupeByAddress(findings);

  const findingsBySeverity: Record<Severity, number> = {
    critical: 0,
    high: 0,
    medium: 0,
    low: 0,
    info: 0,
  };
  const findingsByRule: Record<string, number> = {};
  for (const finding of findings) {
    findingsBySeverity[finding.severity]++;
    findingsByRule[finding.rule] = (findingsByRule[finding.rule] ?? 0) + 1;
  }

  const complete =
    graph.stats.opaqueNodes === 0 &&
    graph.stats.unresolvedRefs === 0 &&
    graph.stats.unparsedFormulas === 0 &&
    graph.stats.externalRefNodes === 0;

  return {
    findings,
    health: computeHealth(findings),
    rulesRun,
    coverage: {
      formulaCells: graph.stats.formulaCells,
      opaqueNodes: graph.stats.opaqueNodes,
      unresolvedRefs: graph.stats.unresolvedRefs,
      unparsedFormulas: graph.stats.unparsedFormulas,
      externalRefNodes: graph.stats.externalRefNodes,
      complete,
      caveat: complete
        ? "Every formula was parsed and every reference resolved: dependency coverage is complete."
        : `Dependency coverage is INCOMPLETE: ${graph.stats.opaqueNodes} opaque node(s), ` +
          `${graph.stats.unresolvedRefs} unresolved reference(s), ` +
          `${graph.stats.unparsedFormulas} unparsed formula(s), ` +
          `${graph.stats.externalRefNodes} external link(s). Findings below are accurate for what ` +
          `we can see, but impact analysis may miss paths through these cells (see AUD-011).`,
    },
    stats: {
      findingsBySeverity,
      findingsByRule,
      durationMs: Date.now() - started,
      llmCallsUsed: 0,
    },
  };
}

/** Findings whose auto-fix is LOW risk — the "Fix all safe issues" set. */
export function safeAutoFixes(report: AuditReport): Finding[] {
  return report.findings.filter((finding) => finding.autoFix?.risk === "low");
}

/**
 * Audit engine types (handoff §6) — the wedge feature.
 *
 * Hard requirement: every rule here is deterministic and runs with ZERO LLM
 * calls, so the whole audit can be demoed inside any compliance boundary.
 * The only LLM-assisted rules are AUD-012+ (model-risk judgment), which are
 * clearly labelled as judgment rather than fact and are opt-in.
 */

import { CellRef } from "../graph/graph";

export type Severity = "critical" | "high" | "medium" | "low" | "info";

/** How confident the rule is that this is genuinely a defect. */
export type Confidence = "certain" | "likely" | "possible";

export interface AutoFix {
  /** Human-readable description of what the fix would do. */
  description: string;
  /** Risk tier of applying it (see handoff §4). Only "low" is auto-applied. */
  risk: "low" | "medium" | "high";
  /** Target cell and the formula/value to write. */
  edits: Array<{ sheet: string; row: number; col: number; formula?: string; value?: string | number }>;
}

export interface Finding {
  rule: string;
  title: string;
  severity: Severity;
  confidence: Confidence;
  /** "Sheet!A1" of the primary offending cell. */
  address: string;
  sheet: string;
  row: number;
  col: number;
  /** Plain-language explanation, rendered from the rule's template. */
  explanation: string;
  /** The formula or value that triggered the finding. */
  evidence?: string;
  /** Cells to highlight when the user clicks "Trace". */
  trace: CellRef[];
  /** Number of downstream cells affected — drives ranking. */
  blastRadius: number;
  autoFix?: AutoFix;
  /**
   * Other rules that fired on this same cell. One defect often trips several
   * rules — a #REF! typed over a filled row is both an error cell and a
   * pattern break — and listing the cell once with its secondary rules noted
   * is far more readable than repeating it.
   */
  alsoFlaggedBy?: Array<{ rule: string; title: string; severity: Severity }>;
}

export interface RuleMeta {
  id: string;
  title: string;
  severity: Severity;
  /** One-line statement of what the rule looks for, shown in the UI. */
  description: string;
  /** True when the rule needs an LLM (AUD-012+). Default false. */
  requiresLlm?: boolean;
}

export interface AuditScope {
  /** Restrict to these sheets; empty = whole workbook. */
  sheets?: string[];
  /** Rule ids to run; empty = all deterministic rules. */
  rules?: string[];
  /** Include LLM-assisted judgment rules (AUD-012+). */
  includeJudgment?: boolean;
}

export interface BalanceAssertion {
  /** Display name, e.g. "Balance sheet ties". */
  name: string;
  sheet: string;
  /** A1 of the cell that must satisfy the assertion. */
  a1: string;
  /** Expected value (default 0). */
  expected?: number;
  tolerance?: number;
}

export interface HealthScore {
  /** 0-100. Weighted by severity and blast radius. */
  score: number;
  band: "healthy" | "needs-attention" | "at-risk" | "critical";
  breakdown: Array<{ severity: Severity; count: number; penalty: number }>;
}

export interface AuditReport {
  findings: Finding[];
  health: HealthScore;
  /** Rules that ran, so the UI can distinguish "clean" from "not checked". */
  rulesRun: string[];
  /** Honest coverage statement — see INV-4. */
  coverage: {
    formulaCells: number;
    opaqueNodes: number;
    unresolvedRefs: number;
    unparsedFormulas: number;
    externalRefNodes: number;
    complete: boolean;
    caveat: string;
  };
  stats: {
    findingsBySeverity: Record<Severity, number>;
    findingsByRule: Record<string, number>;
    durationMs: number;
    llmCallsUsed: number;
  };
}

export const SEVERITY_WEIGHT: Record<Severity, number> = {
  critical: 25,
  high: 12,
  medium: 5,
  low: 2,
  info: 0,
};

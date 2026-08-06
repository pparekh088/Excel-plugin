/**
 * Agent runtime (handoff §7):
 *
 *   intent -> context assembly (WIL) -> PLANNER -> user approval
 *     -> EXECUTOR (accumulates ONE change set) -> preview
 *     -> approval -> drift check (INV-8) -> apply -> recalculate
 *     -> VERIFIER -> repair loop (<=3) -> rollback offer -> EXPLAIN
 *
 * The runtime never partially applies: every step contributes edits to a
 * single change set which is applied atomically or not at all (INV-2).
 * Rollback is offered — and its fidelity honestly reported — whenever
 * verification fails after the repair budget is exhausted.
 */

import { runAudit } from "../audit/engine";
import { BalanceAssertion } from "../audit/types";
import {
  applyToWorkbook,
  checkDrift,
  explainChangeSet,
  proposeChangeSet,
  rollback,
} from "../changeset/engine";
import { ChangeSet, Edit, RollbackReport, isCellEdit } from "../changeset/types";
import { DependencyGraph } from "../graph/graph";
import { Workbook } from "../model/workbook";
import type { CellAddr, Node as AstNode } from "../parser/ast";
import { parseFormula } from "../parser/parser";
import { toFormulaText } from "../parser/serialize";
import { Simulator } from "../sim/simulator";
import { buildWil } from "../wil/serialize";
import { CostMeter, LlmProvider, TaskClass, tierFor } from "./llm";
import { Plan, isAutoApprovable, parsePlan, renderPlan } from "./plan";
import { renderToolCatalogue } from "./tools";
import { VerificationResult, errorCellSet, verify } from "./verify";

export type ApprovalDecision = "approve" | "reject";

export interface ApprovalRequest {
  kind: "plan" | "changeset";
  plan?: Plan;
  changeSet?: ChangeSet;
  /** Rendered text the UI shows. */
  text: string;
}

export type Approver = (request: ApprovalRequest) => Promise<ApprovalDecision>;

/** Turns a validated plan step into concrete edits. */
export type StepExecutor = (
  step: Plan["steps"][number],
  workbook: Workbook
) => Edit[] | Promise<Edit[]>;

export interface RunOptions {
  intent: string;
  taskClass?: TaskClass;
  provider: LlmProvider;
  approver: Approver;
  executor: StepExecutor;
  costMeter?: CostMeter;
  assertions?: BalanceAssertion[];
  /** Trusted-session mode: auto-approve plans that are entirely LOW risk. */
  trustedSession?: boolean;
  maxRepairs?: number;
  /** Injected for tests; defaults to a live drift check against the workbook. */
  now?: () => string;
}

export type RunOutcome =
  | "applied"
  | "applied-with-warnings"
  | "rejected"
  | "aborted-drift"
  | "rolled-back"
  | "plan-failed"
  | "no-op";

export interface RunResult {
  outcome: RunOutcome;
  plan?: Plan;
  changeSet?: ChangeSet;
  verification?: VerificationResult;
  rollbackReport?: RollbackReport;
  explanation?: string;
  repairAttempts: number;
  /** Everything the user should be told, in order. */
  transcript: string[];
  costUsd: number;
  llmCalls: number;
}

const PLANNER_SYSTEM = `You are Ledger, an autonomous Excel engineer.

You never write code that touches the workbook. You emit ordered, typed tool
calls only, chosen from the catalogue below. Anything not in the catalogue is
unavailable — do not invent tools or ask for arbitrary code execution.

Respond with a single JSON object:
{
  "summary": "one line describing what this plan achieves",
  "steps": [
    { "tool": "formula.set", "params": { ... }, "rationale": "why this step" }
  ]
}

Rules:
- Prefer inspection tools before mutations; never guess a range you have not read.
- Preserve existing formulas unless the intent explicitly requires changing them.
- Formulas are en-US (comma separators, en-US function names).
- Every mutating step needs a rationale a reviewer can check.`;

export async function runAgent(
  workbook: Workbook,
  options: RunOptions
): Promise<RunResult> {
  const transcript: string[] = [];
  const costMeter = options.costMeter ?? new CostMeter();
  const maxRepairs = options.maxRepairs ?? 3;
  let repairAttempts = 0;

  const finish = (outcome: RunOutcome, extra: Partial<RunResult> = {}): RunResult => ({
    outcome,
    repairAttempts,
    transcript,
    costUsd: costMeter.totalCostUsd,
    llmCalls: costMeter.totalCalls,
    ...extra,
  });

  // ---- 1. context assembly: WIL only, never raw grids (INV-6) ----------
  const graph = DependencyGraph.build(workbook);
  const wil = buildWil(workbook, graph);
  transcript.push(
    `Read the workbook: ${wil.stats.sheets} sheets, ${wil.stats.formulaCells} formulas ` +
      `in ${wil.stats.runNodes} blocks.`
  );

  // ---- 2. planning ------------------------------------------------------
  const plannerResponse = await options.provider.complete({
    tier: tierFor("planner"),
    messages: [
      { role: "system", content: PLANNER_SYSTEM },
      { role: "system", content: renderToolCatalogue() },
      { role: "user", content: `WORKBOOK SUMMARY\n${wil.text}` },
      { role: "user", content: `INTENT\n${options.intent}` },
    ],
  });
  costMeter.record(tierFor("planner"), plannerResponse.usage);

  const parsed = parsePlan(plannerResponse.text);
  if (!parsed.plan) {
    transcript.push(
      `I could not produce a valid plan: ${parsed.errors.map((error) => error.message).join("; ")}`
    );
    return finish("plan-failed");
  }
  const plan = parsed.plan;
  transcript.push(`Plan (${plan.steps.length} steps, risk ${plan.risk}):\n${renderPlan(plan)}`);

  // ---- 3. plan approval (skippable only for all-LOW plans) -------------
  const autoApprove = options.trustedSession === true && isAutoApprovable(plan);
  if (!autoApprove) {
    const decision = await options.approver({
      kind: "plan",
      plan,
      text: renderPlan(plan),
    });
    if (decision !== "approve") {
      transcript.push("Plan rejected; nothing was changed.");
      return finish("rejected", { plan });
    }
  } else {
    transcript.push("Trusted session: plan is entirely low risk, applying without prompting.");
  }

  // ---- 4. execution: accumulate ONE change set -------------------------
  const edits: Edit[] = [];
  for (const step of plan.steps) {
    const produced = await options.executor(step, workbook);
    edits.push(...produced);
  }
  if (edits.length === 0) {
    transcript.push("The plan produced no changes — nothing to apply.");
    return finish("no-op", { plan });
  }

  let changeSet = proposeChangeSet(workbook, {
    intent: options.intent,
    summary: plan.summary,
    edits,
    graph,
  });

  // ---- 5. change-set approval on the previewed diff --------------------
  const preview = explainChangeSet(changeSet);
  transcript.push(preview);
  const changeApproval =
    options.trustedSession === true && changeSet.risk === "low"
      ? "approve"
      : await options.approver({ kind: "changeset", changeSet, text: preview });
  if (changeApproval !== "approve") {
    changeSet.status = "aborted";
    changeSet.failureReason = "User declined the change set.";
    transcript.push("Change set declined; nothing was changed.");
    return finish("rejected", { plan, changeSet });
  }

  // ---- 6. drift check (INV-8) -----------------------------------------
  const drift = checkDrift(workbook, changeSet);
  if (!drift.clean) {
    changeSet.status = "aborted";
    changeSet.failureReason = "Workbook changed since the plan was made.";
    transcript.push(
      `Aborted before writing anything: ${drift.entries.length} cell(s) changed since I read ` +
        `them (${drift.entries.slice(0, 5).map((entry) => entry.address).join(", ")}). ` +
        `Someone else may be editing. Re-run to plan against the current state.`
    );
    return finish("aborted-drift", { plan, changeSet });
  }

  // ---- 7. apply + recalculate + verify, with repair loop ---------------
  const preExistingErrors = errorCellSet(workbook);
  const cyclesBefore = graph.stats.cycleCount;

  applyToWorkbook(workbook, changeSet);
  changeSet.status = "applied";
  changeSet.appliedAt = (options.now ?? (() => new Date().toISOString()))();
  Simulator.of(workbook).recalculate();

  let verification = verify(workbook, changeSet, {
    preExistingErrors,
    assertions: options.assertions,
    cyclesBefore,
  });

  while (!verification.ok && repairAttempts < maxRepairs) {
    repairAttempts++;
    transcript.push(
      `Verification failed (attempt ${repairAttempts}/${maxRepairs}): ` +
        verification.issues.map((issue) => issue.detail).join(" ")
    );

    const repairResponse = await options.provider.complete({
      tier: tierFor("executor"),
      messages: [
        { role: "system", content: PLANNER_SYSTEM },
        { role: "system", content: renderToolCatalogue() },
        {
          role: "user",
          content:
            `The change set was applied but verification failed:\n` +
            verification.issues.map((issue) => `- ${issue.detail}`).join("\n") +
            `\n\nPropose a corrective plan, or return {"steps": []} if you cannot fix it.`,
        },
      ],
    });
    costMeter.record(tierFor("executor"), repairResponse.usage);

    const repairPlan = parsePlan(repairResponse.text);
    if (!repairPlan.plan || repairPlan.plan.steps.length === 0) {
      transcript.push("No corrective plan was possible.");
      break;
    }

    const repairEdits: Edit[] = [];
    for (const step of repairPlan.plan.steps) {
      repairEdits.push(...(await options.executor(step, workbook)));
    }
    if (repairEdits.length === 0) break;

    // The repair joins the SAME change set so one rollback undoes everything.
    const repairSet = proposeChangeSet(workbook, {
      intent: `repair: ${options.intent}`,
      summary: repairPlan.plan.summary,
      edits: repairEdits,
    });
    applyToWorkbook(workbook, repairSet);
    Simulator.of(workbook).recalculate();

    changeSet = mergeChangeSets(changeSet, repairSet);
    verification = verify(workbook, changeSet, {
      preExistingErrors,
      assertions: options.assertions,
      cyclesBefore,
    });
  }

  // ---- 8. outcome, rollback offer, explanation ------------------------
  const explanation = explainChangeSet(changeSet);

  if (!verification.ok) {
    transcript.push(
      `I could not get this to a clean state after ${repairAttempts} repair attempt(s). ` +
        `Remaining problems: ${verification.issues.map((issue) => issue.detail).join(" ")}`
    );
    const decision = await options.approver({
      kind: "changeset",
      changeSet,
      text: `Roll back this change set?\n\n${explanation}`,
    });
    if (decision === "approve") {
      const report = rollback(workbook, changeSet);
      Simulator.of(workbook).recalculate();
      changeSet.status = "rolled-back";
      changeSet.rolledBackAt = (options.now ?? (() => new Date().toISOString()))();
      transcript.push(
        `Rolled back ${report.restoredCells} cell(s).` +
          (report.unrestorable.length > 0
            ? ` Could NOT restore: ${report.unrestorable.join(" ")}`
            : " Everything I changed was restored.")
      );
      return finish("rolled-back", {
        plan,
        changeSet,
        verification,
        rollbackReport: report,
        explanation,
      });
    }
    transcript.push("Keeping the changes despite the warnings, as instructed.");
    return finish("applied-with-warnings", { plan, changeSet, verification, explanation });
  }

  transcript.push(explanation);
  return finish("applied", { plan, changeSet, verification, explanation });
}

/** Fold a repair change set into the original so one rollback undoes both. */
function mergeChangeSets(original: ChangeSet, repair: ChangeSet): ChangeSet {
  const snapshots = [...original.snapshots];
  const known = new Set(
    snapshots.map((snap) => `${snap.sheet.toUpperCase()}!${snap.row},${snap.col}`)
  );
  // Only add snapshots for cells the original did not already capture —
  // the original's snapshot is the true "before" state.
  for (const snap of repair.snapshots) {
    const key = `${snap.sheet.toUpperCase()}!${snap.row},${snap.col}`;
    if (!known.has(key)) {
      snapshots.push(snap);
      known.add(key);
    }
  }
  return {
    ...original,
    edits: [...original.edits, ...repair.edits],
    snapshots,
    diff: [...original.diff, ...repair.diff],
    impact: {
      affectedCells: original.impact.affectedCells + repair.impact.affectedCells,
      affectedNodes: original.impact.affectedNodes + repair.impact.affectedNodes,
      affectedOutputs: [
        ...new Set([...original.impact.affectedOutputs, ...repair.impact.affectedOutputs]),
      ],
      affectedCharts: [
        ...new Set([...original.impact.affectedCharts, ...repair.impact.affectedCharts]),
      ],
      affectedPivots: [
        ...new Set([...original.impact.affectedPivots, ...repair.impact.affectedPivots]),
      ],
      opaqueDownstream: original.impact.opaqueDownstream + repair.impact.opaqueDownstream,
    },
  };
}

/** Default executor: maps the mutation tools onto concrete cell edits. */
export function createDefaultExecutor(): StepExecutor {
  return (step, workbook) => {
    const params = step.params as Record<string, unknown>;
    const sheet = typeof params.sheet === "string" ? params.sheet : "";
    const a1 = typeof params.a1 === "string" ? params.a1 : "";

    switch (step.tool) {
      case "formula.set": {
        const formula = typeof params.formula === "string" ? params.formula : "";
        const anchor = parseSingleCell(a1);
        if (!anchor || !formula) return [];
        const fillDown = typeof params.fillDown === "number" ? params.fillDown : 0;
        const fillRight = typeof params.fillRight === "number" ? params.fillRight : 0;
        const edits: Edit[] = [];
        for (let dr = 0; dr <= fillDown; dr++) {
          for (let dc = 0; dc <= fillRight; dc++) {
            edits.push({
              kind: "setFormula",
              sheet,
              row: anchor.row + dr,
              col: anchor.col + dc,
              // Translate relative references the way a fill would.
              formula: dr === 0 && dc === 0 ? formula : translateFormula(formula, dr, dc),
            });
          }
        }
        return edits;
      }
      case "range.write": {
        const values = Array.isArray(params.values) ? (params.values as unknown[][]) : [];
        const anchor = parseSingleCell(a1);
        if (!anchor) return [];
        const edits: Edit[] = [];
        values.forEach((row, dr) => {
          if (!Array.isArray(row)) return;
          row.forEach((value, dc) => {
            edits.push({
              kind: "setValue",
              sheet,
              row: anchor.row + dr,
              col: anchor.col + dc,
              value: value as never,
            });
          });
        });
        return edits;
      }
      case "number_format.set": {
        const format = typeof params.format === "string" ? params.format : "General";
        return expandRange(sheet, a1).map((cell) => ({
          kind: "setNumberFormat" as const,
          sheet,
          row: cell.row,
          col: cell.col,
          numberFormat: format,
        }));
      }
      case "range.clear":
        return expandRange(sheet, a1).map((cell) => ({
          kind: "clear" as const,
          sheet,
          row: cell.row,
          col: cell.col,
        }));
      case "sheet.create":
        return [
          {
            kind: "createSheet",
            name: typeof params.name === "string" ? params.name : "Sheet",
          },
        ];
      case "name.define":
        return [
          {
            kind: "defineName",
            name: typeof params.name === "string" ? params.name : "",
            refersTo: typeof params.refersTo === "string" ? params.refersTo : "",
          },
        ];
      default:
        // Inspection and control tools produce no edits.
        void workbook;
        return [];
    }
  };
}

function parseSingleCell(a1: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([0-9]+)/.exec(a1);
  if (!match) return null;
  let col = 0;
  for (const ch of match[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(match[2]) - 1, col: col - 1 };
}

function expandRange(sheet: string, a1: string): Array<{ row: number; col: number }> {
  void sheet;
  const parts = a1.split(":");
  const start = parseSingleCell(parts[0] ?? "");
  if (!start) return [];
  const end = parts[1] ? parseSingleCell(parts[1]) : start;
  if (!end) return [start];
  const cells: Array<{ row: number; col: number }> = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
    for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
      cells.push({ row, col });
    }
  }
  return cells;
}

/**
 * Shift relative references by (dr, dc), leaving $-anchored parts alone —
 * what Excel does when you fill a formula.
 *
 * This goes through the parser rather than a regex: a pattern like
 * /[A-Za-z]{1,3}[0-9]+/ happily matches "eet2" inside "Sheet2!B5" and
 * corrupts sheet names, defined names and string literals.
 */
export function translateFormula(formula: string, dr: number, dc: number): string {
  const parsed = parseFormula(formula);
  if (!parsed.ok) return formula; // never mangle something we do not understand

  const shift = (addr: CellAddr): CellAddr => ({
    row: addr.rowAbs ? addr.row : Math.max(0, addr.row + dr),
    col: addr.colAbs ? addr.col : Math.max(0, addr.col + dc),
    rowAbs: addr.rowAbs,
    colAbs: addr.colAbs,
  });

  const walk = (node: AstNode): AstNode => {
    switch (node.kind) {
      case "cell":
        return { ...node, addr: shift(node.addr) };
      case "range":
        return { ...node, start: shift(node.start), end: shift(node.end) };
      case "colRange":
        return {
          ...node,
          startCol: node.startAbs ? node.startCol : Math.max(0, node.startCol + dc),
          endCol: node.endAbs ? node.endCol : Math.max(0, node.endCol + dc),
        };
      case "rowRange":
        return {
          ...node,
          startRow: node.startAbs ? node.startRow : Math.max(0, node.startRow + dr),
          endRow: node.endAbs ? node.endRow : Math.max(0, node.endRow + dr),
        };
      case "unary":
      case "percent":
      case "implicitIntersection":
      case "spill":
        return { ...node, operand: walk(node.operand) };
      case "binary":
        return { ...node, left: walk(node.left), right: walk(node.right) };
      case "group":
        return { ...node, expr: walk(node.expr) };
      case "func":
        return { ...node, args: node.args.map(walk) };
      case "callExpr":
        return { ...node, callee: walk(node.callee), args: node.args.map(walk) };
      case "array":
        return { ...node, rows: node.rows.map((row) => row.map(walk)) };
      default:
        return node;
    }
  };

  return toFormulaText(walk(parsed.ast));
}

/** Convenience: run the deterministic audit as a "plan-free" agent action. */
export function auditAction(workbook: Workbook) {
  return runAudit(workbook);
}

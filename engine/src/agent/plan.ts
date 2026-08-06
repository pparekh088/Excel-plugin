/**
 * Plan representation and parsing (INV-1).
 *
 * The planner emits JSON steps. Anything that is not a known tool with
 * well-formed params is REJECTED here, before it can reach a change set —
 * this is the engine-side half of "the LLM emits tool calls only". The
 * server independently validates the same calls against the exported JSON
 * Schemas, so a malformed or hallucinated call has to get past both.
 */

import { toolByName } from "./tools";
import { RiskTier } from "../changeset/types";

export interface PlanStep {
  tool: string;
  params: Record<string, unknown>;
  /** Why this step is needed — shown to the user in the plan review. */
  rationale: string;
  risk: RiskTier | "none";
}

export interface Plan {
  steps: PlanStep[];
  /** One-line statement of what the plan achieves. */
  summary: string;
  /** Highest risk tier across the steps. */
  risk: RiskTier | "none";
}

export interface PlanParseError {
  message: string;
  stepIndex?: number;
  raw?: unknown;
}

export interface PlanParseResult {
  plan?: Plan;
  errors: PlanParseError[];
}

const RISK_ORDER: Array<RiskTier | "none"> = ["none", "low", "medium", "high"];

function maxRisk(risks: Array<RiskTier | "none">): RiskTier | "none" {
  return risks.reduce(
    (highest, risk) => (RISK_ORDER.indexOf(risk) > RISK_ORDER.indexOf(highest) ? risk : highest),
    "none" as RiskTier | "none"
  );
}

/** Pull the first JSON object out of a model response that may have prose. */
function extractJson(text: string): unknown {
  const fenced = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
  const candidate = fenced?.[1] ?? text;
  const start = candidate.indexOf("{");
  const end = candidate.lastIndexOf("}");
  if (start < 0 || end <= start) return undefined;
  try {
    return JSON.parse(candidate.slice(start, end + 1));
  } catch {
    return undefined;
  }
}

export function parsePlan(text: string): PlanParseResult {
  const errors: PlanParseError[] = [];
  const parsed = extractJson(text);
  if (parsed === undefined || typeof parsed !== "object" || parsed === null) {
    return { errors: [{ message: "Response did not contain a JSON object.", raw: text }] };
  }

  const record = parsed as Record<string, unknown>;
  const rawSteps = record.steps;
  if (!Array.isArray(rawSteps)) {
    return { errors: [{ message: "Plan has no `steps` array.", raw: parsed }] };
  }

  const steps: PlanStep[] = [];
  rawSteps.forEach((rawStep, index) => {
    if (typeof rawStep !== "object" || rawStep === null) {
      errors.push({ message: "Step is not an object.", stepIndex: index, raw: rawStep });
      return;
    }
    const step = rawStep as Record<string, unknown>;
    const toolName = step.tool;
    if (typeof toolName !== "string") {
      errors.push({ message: "Step has no `tool` name.", stepIndex: index, raw: rawStep });
      return;
    }
    // INV-1: unknown tools are rejected, never passed through.
    const spec = toolByName(toolName);
    if (!spec) {
      errors.push({
        message: `Unknown tool "${toolName}". The agent may only call tools in the catalogue.`,
        stepIndex: index,
      });
      return;
    }
    const params = step.params;
    if (params !== undefined && (typeof params !== "object" || params === null || Array.isArray(params))) {
      errors.push({ message: "Step `params` must be an object.", stepIndex: index });
      return;
    }
    steps.push({
      tool: toolName,
      params: (params as Record<string, unknown>) ?? {},
      rationale: typeof step.rationale === "string" ? step.rationale : "",
      risk: spec.risk,
    });
  });

  if (errors.length > 0) return { errors };
  if (steps.length === 0) {
    return { errors: [{ message: "Plan contains no valid steps." }] };
  }

  return {
    plan: {
      steps,
      summary: typeof record.summary === "string" ? record.summary : "(no summary provided)",
      risk: maxRisk(steps.map((step) => step.risk)),
    },
    errors: [],
  };
}

/** Human-readable plan for the approval UI. */
export function renderPlan(plan: Plan): string {
  const lines = [plan.summary, "", `Risk: ${plan.risk}`, ""];
  plan.steps.forEach((step, index) => {
    lines.push(`${index + 1}. ${step.tool} [${step.risk}]`);
    if (step.rationale) lines.push(`   ${step.rationale}`);
    const params = Object.entries(step.params)
      .map(([key, value]) => `${key}=${JSON.stringify(value)}`)
      .join(", ");
    if (params) lines.push(`   ${params}`);
  });
  return lines.join("\n");
}

/** A plan qualifies for trusted-session auto-approval only if it is all LOW risk. */
export function isAutoApprovable(plan: Plan): boolean {
  return plan.steps.every((step) => step.risk === "low" || step.risk === "none");
}

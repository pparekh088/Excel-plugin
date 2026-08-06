/**
 * AI.* custom functions (handoff §8), running in the separate custom-functions
 * JS runtime.
 *
 * Every call goes through the shared AiCoordinator (INV-7): 250ms debounce,
 * content-hash cache, request coalescing, and a hard per-recalc budget after
 * which cells return #AI_BUDGET! and the task pane raises a banner.
 *
 * Custom functions NEVER mutate other cells (INV-7). They read their arguments
 * and return a value — nothing else. There is no path from here to a write.
 */

import { AiCoordinator, BUDGET_ERROR, type AiRequest } from "ledger-engine";
import { createAuthProvider } from "../auth/provider";
import { getBackendUrl } from "../config";

/* global CustomFunctions */

const MODEL_TIER = "cheap";

interface BatchResponse {
  results: Array<string | number | boolean>;
}

/** Marks AI-derived cells so the audit engine can inventory them (§8). */
const AI_CELL_MARKER = "AI:";

/**
 * The custom-functions runtime is a separate JavaScript context from the task
 * pane unless the manifest declares a SHARED runtime — which ours does. With
 * it, this module and the task pane are the same instance, so the token cache
 * is shared and a recalculation does not re-acquire per batch.
 *
 * These calls used to go out with no Authorization header at all. Against a
 * backend in dev mode that works, which is exactly why it survived: the moment
 * the backend runs LEDGER_AUTH_MODE=entra, every AI.* cell in every workbook
 * returns an error, and the add-in has no way to say why.
 */
const auth = createAuthProvider();

async function callBackend(requests: AiRequest[]): Promise<Array<string | number | boolean>> {
  const token = await auth.getAccessToken();
  const response = await fetch(`${getBackendUrl()}/api/v1/ai/batch`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
    body: JSON.stringify({ requests }),
  });
  if (!response.ok) {
    // 401 is worth naming: it is the difference between "the model failed" and
    // "you are not signed in", and the user can only act on the second.
    if (response.status === 401 || response.status === 403) {
      throw new Error(
        `AI backend rejected the request (${response.status}) — sign in from the Ledger ` +
          `task pane, then recalculate.`
      );
    }
    throw new Error(`AI backend returned ${response.status}`);
  }
  const body = (await response.json()) as BatchResponse;
  return body.results;
}

const coordinator = new AiCoordinator({
  handler: callBackend,
  onBudgetExceeded: (used, budget) => {
    // The task pane listens for this and shows a banner; without it a user
    // sees #AI_BUDGET! in the grid with no explanation.
    try {
      window.dispatchEvent(
        new CustomEvent("ledger:ai-budget-exceeded", { detail: { used, budget } })
      );
    } catch {
      /* the custom-functions runtime may not share a window; the error value
         in the cell is still the primary signal */
    }
  },
});

/** Excel signals the start of a recalculation cycle by clearing prior results. */
export function startRecalcCycle(): void {
  coordinator.startRecalcCycle();
}

function flatten(range: unknown): unknown[] {
  if (!Array.isArray(range)) return [range];
  return (range as unknown[][]).flat();
}

async function ask(request: AiRequest): Promise<string | number | boolean> {
  const result = await coordinator.request(request);
  return result.value;
}

/**
 * Ask a question about a range of cells.
 * @customfunction AI.ASK
 * @param range Cells to consider.
 * @param prompt What to ask about them.
 * @returns The model's answer.
 */
export async function aiAsk(range: unknown, prompt: string): Promise<string | number | boolean> {
  return ask({ fn: "AI.ASK", inputs: flatten(range), prompt, model: MODEL_TIER });
}

/**
 * Extract a named field from a cell's text.
 * @customfunction AI.EXTRACT
 * @param cell Text to extract from.
 * @param field Field name, e.g. "invoice number".
 * @returns The extracted value, or #N/A when absent.
 */
export async function aiExtract(cell: unknown, field: string): Promise<string | number | boolean> {
  return ask({ fn: "AI.EXTRACT", inputs: [cell], prompt: field, model: MODEL_TIER });
}

/**
 * Classify a cell into one of the supplied categories.
 * @customfunction AI.CLASSIFY
 * @param cell Value to classify.
 * @param categories Comma-separated category list.
 * @returns One of the categories.
 */
export async function aiClassify(
  cell: unknown,
  categories: string
): Promise<string | number | boolean> {
  return ask({ fn: "AI.CLASSIFY", inputs: [cell], prompt: categories, model: MODEL_TIER });
}

/**
 * Fuzzy-match two entity names.
 * @customfunction AI.MATCH
 * @param a First value.
 * @param b Second value.
 * @returns TRUE when they refer to the same entity.
 */
export async function aiMatch(a: unknown, b: unknown): Promise<string | number | boolean> {
  return ask({
    fn: "AI.MATCH",
    inputs: [a, b],
    prompt: "Do these refer to the same entity?",
    model: MODEL_TIER,
  });
}

/**
 * Forecast future periods from a history range.
 *
 * The numbers come from exponential smoothing computed server-side, not from
 * a language model — an invented forecast is indistinguishable from a computed
 * one, and in a financial model that difference is the whole point.
 *
 * @customfunction AI.FORECAST
 * @param range Historical values, oldest first.
 * @param horizon How many periods ahead.
 * @param seasonLength Optional periods per season (12 monthly, 4 quarterly).
 * @returns The forecast values as a spilled column.
 */
export async function aiForecast(
  range: unknown,
  horizon: number,
  seasonLength?: number
): Promise<number[][] | string> {
  const history = flatten(range).filter((value): value is number => typeof value === "number");
  const response = await fetch(`${getBackendUrl()}/api/v1/ai/forecast`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ series: history, horizon, seasonLength: seasonLength ?? 0 }),
  });
  if (!response.ok) return `#ERROR! forecast backend returned ${response.status}`;
  const body = (await response.json()) as { values: number[] };
  return body.values.map((value) => [value]);
}

/** Cell marker helper, used by the WIL to inventory AI-derived values. */
export function aiMarker(fn: string): string {
  return `${AI_CELL_MARKER}${fn}`;
}

export { BUDGET_ERROR };

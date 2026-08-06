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
import { getActiveSessionId, postToBackend } from "../api/backend";

/* global CustomFunctions */

const MODEL_TIER = "cheap";

/** Marks AI-derived cells so the audit engine can inventory them (§8). */
const AI_CELL_MARKER = "AI:";

/**
 * A cell whose individual AI call failed shows this rather than poisoning the
 * whole batch. Distinct from BUDGET_ERROR: budget is a policy refusal the user
 * configured, this is a backend failure the user did not.
 */
export const AI_ERROR = "#AI_ERROR!";

/**
 * What the server actually returns from /ai/batch: one envelope per request,
 * because one bad cell must not fail the other 199 in the batch. Excel can
 * only display scalars, so the envelope MUST be unwrapped here — handing the
 * object through would render "[object Object]" in the grid.
 */
interface AiResultItem {
  ok: boolean;
  value?: unknown;
  error?: string | null;
}

interface BatchResponse {
  results: AiResultItem[];
}

/** Unwrap one server result envelope into the scalar the cell will hold. */
export function unwrapAiResult(item: AiResultItem): string | number | boolean {
  if (!item.ok) return AI_ERROR;
  const value = item.value;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  // A null/undefined or structured value is a contract violation, not an
  // answer. Show the error marker rather than coercing garbage into the grid.
  return AI_ERROR;
}

async function callBackend(requests: AiRequest[]): Promise<Array<string | number | boolean>> {
  const body = await postToBackend<BatchResponse>("/api/v1/ai/batch", {
    requests,
    // Ties this batch's spend to the workbook session the task pane opened.
    // Null is accepted (metering falls back to the principal's own bucket),
    // but with the shared runtime the session is normally present.
    session_id: getActiveSessionId(),
  });
  if (!Array.isArray(body.results) || body.results.length !== requests.length) {
    throw new Error(
      `AI backend returned ${body.results?.length ?? 0} result(s) for ${requests.length} request(s).`
    );
  }
  return body.results.map(unwrapAiResult);
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
  try {
    // Same authenticated path as every other backend call. Forecast used to
    // have its own bare fetch, which meant it was the one AI.* function still
    // posting anonymously after auth landed everywhere else.
    const body = await postToBackend<{ values: number[] }>("/api/v1/ai/forecast", {
      series: history,
      horizon,
      seasonLength: seasonLength ?? 0,
    });
    return body.values.map((value) => [value]);
  } catch (error) {
    return `#ERROR! ${error instanceof Error ? error.message : String(error)}`;
  }
}

/** Cell marker helper, used by the WIL to inventory AI-derived values. */
export function aiMarker(fn: string): string {
  return `${AI_CELL_MARKER}${fn}`;
}

export { BUDGET_ERROR };

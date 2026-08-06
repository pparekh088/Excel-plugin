/**
 * AI custom-function batching, caching and budget control (INV-7).
 *
 * Dragging =AI.CLASSIFY down 5,000 rows fires 5,000 independent calls unless
 * something coordinates them. This module is that coordinator, and it lives in
 * the engine (not the custom-functions runtime) so it is unit-testable without
 * Excel:
 *
 *  - request coalescing on a 250ms debounce, so a drag becomes a few batches
 *  - content-hash cache keyed on (inputs, prompt, model), so re-runs are free
 *  - identical in-flight requests share one result rather than duplicating
 *  - a hard circuit breaker: past the per-recalc budget, every further cell
 *    returns #AI_BUDGET! and a banner is raised, instead of quietly spending
 *
 * The breaker is deliberately blunt. A runaway recalculation on a large sheet
 * is the failure mode that produces an unexpected four-figure bill, and a
 * visible error in some cells is far better than that.
 */

export const BUDGET_ERROR = "#AI_BUDGET!";
export const DEFAULT_BUDGET = 200;
export const DEFAULT_DEBOUNCE_MS = 250;
export const DEFAULT_BATCH_SIZE = 25;

export interface AiRequest {
  /** Function name, e.g. "AI.CLASSIFY". */
  fn: string;
  /** Cell inputs, already resolved to values. */
  inputs: unknown[];
  /** Prompt or category list, depending on the function. */
  prompt: string;
  model: string;
}

export interface AiResult {
  value: string | number | boolean;
  /** True when served from cache — surfaced in the cost dashboard. */
  cached: boolean;
}

export type BatchHandler = (requests: AiRequest[]) => Promise<Array<string | number | boolean>>;

/** Stable content hash over (inputs, prompt, model) — the cache key. */
export function contentHash(request: AiRequest): string {
  const canonical = JSON.stringify({
    fn: request.fn,
    inputs: request.inputs,
    prompt: request.prompt,
    model: request.model,
  });
  // FNV-1a: short, fast, dependency-free, and collision risk is irrelevant
  // here because a collision only costs a wrong cache hit within one session.
  let hash = 0x811c9dc5;
  for (let i = 0; i < canonical.length; i++) {
    hash ^= canonical.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return `${request.fn}:${hash.toString(36)}:${canonical.length}`;
}

export interface CoordinatorOptions {
  handler: BatchHandler;
  budget?: number;
  debounceMs?: number;
  batchSize?: number;
  /** Injected for tests. */
  scheduler?: (callback: () => void, ms: number) => void;
  /** Raised the first time the budget is exhausted. */
  onBudgetExceeded?: (used: number, budget: number) => void;
}

export interface CoordinatorStats {
  requests: number;
  cacheHits: number;
  batches: number;
  llmCalls: number;
  budgetRejections: number;
  cacheHitRate: number;
}

interface Pending {
  request: AiRequest;
  resolve: (result: AiResult) => void;
  reject: (error: Error) => void;
}

export class AiCoordinator {
  private readonly cache = new Map<string, string | number | boolean>();
  private readonly inflight = new Map<string, Promise<AiResult>>();
  private queue: Pending[] = [];
  private timer: unknown = null;
  private budgetNotified = false;

  private counts = {
    requests: 0,
    cacheHits: 0,
    batches: 0,
    llmCalls: 0,
    budgetRejections: 0,
  };

  constructor(private readonly options: CoordinatorOptions) {}

  private get budget(): number {
    return this.options.budget ?? DEFAULT_BUDGET;
  }

  private get debounceMs(): number {
    return this.options.debounceMs ?? DEFAULT_DEBOUNCE_MS;
  }

  private get batchSize(): number {
    return this.options.batchSize ?? DEFAULT_BATCH_SIZE;
  }

  /** Reset per recalculation cycle — the budget is per cycle, not per session. */
  startRecalcCycle(): void {
    this.counts.llmCalls = 0;
    this.counts.budgetRejections = 0;
    this.budgetNotified = false;
  }

  clearCache(): void {
    this.cache.clear();
  }

  get stats(): CoordinatorStats {
    const total = this.counts.requests;
    return {
      ...this.counts,
      cacheHitRate: total === 0 ? 0 : this.counts.cacheHits / total,
    };
  }

  async request(request: AiRequest): Promise<AiResult> {
    this.counts.requests++;
    const key = contentHash(request);

    const cached = this.cache.get(key);
    if (cached !== undefined) {
      this.counts.cacheHits++;
      return { value: cached, cached: true };
    }

    // Two cells asking the identical question share one call.
    const existing = this.inflight.get(key);
    if (existing) {
      this.counts.cacheHits++;
      return existing;
    }

    if (this.counts.llmCalls >= this.budget) {
      this.counts.budgetRejections++;
      if (!this.budgetNotified) {
        this.budgetNotified = true;
        this.options.onBudgetExceeded?.(this.counts.llmCalls, this.budget);
      }
      return { value: BUDGET_ERROR, cached: false };
    }

    const promise = new Promise<AiResult>((resolve, reject) => {
      this.queue.push({ request, resolve, reject });
      this.scheduleFlush();
    });
    this.inflight.set(key, promise);
    return promise;
  }

  private scheduleFlush(): void {
    if (this.timer !== null) return;
    const schedule =
      this.options.scheduler ?? ((callback: () => void, ms: number) => setTimeout(callback, ms));
    this.timer = 1;
    schedule(() => {
      this.timer = null;
      void this.flush();
    }, this.debounceMs);
  }

  /** Send everything queued, in batches. Exposed for tests and for shutdown. */
  async flush(): Promise<void> {
    while (this.queue.length > 0) {
      const batch = this.queue.splice(0, this.batchSize);
      // Re-check the budget per batch: a long drag can exhaust it mid-flight.
      const remaining = this.budget - this.counts.llmCalls;
      if (remaining <= 0) {
        for (const pending of batch) {
          this.counts.budgetRejections++;
          this.inflight.delete(contentHash(pending.request));
          pending.resolve({ value: BUDGET_ERROR, cached: false });
        }
        if (!this.budgetNotified) {
          this.budgetNotified = true;
          this.options.onBudgetExceeded?.(this.counts.llmCalls, this.budget);
        }
        continue;
      }

      const send = batch.slice(0, remaining);
      const rejected = batch.slice(remaining);
      for (const pending of rejected) {
        this.counts.budgetRejections++;
        this.inflight.delete(contentHash(pending.request));
        pending.resolve({ value: BUDGET_ERROR, cached: false });
      }
      // The banner must appear the moment ANY cell is refused, including when
      // a batch is only partly over budget — otherwise a user sees
      // #AI_BUDGET! in the grid with nothing explaining it.
      if (rejected.length > 0 && !this.budgetNotified) {
        this.budgetNotified = true;
        this.options.onBudgetExceeded?.(this.counts.llmCalls + send.length, this.budget);
      }

      this.counts.batches++;
      this.counts.llmCalls += send.length;

      try {
        const results = await this.options.handler(send.map((pending) => pending.request));
        send.forEach((pending, index) => {
          const key = contentHash(pending.request);
          const value = results[index];
          if (value === undefined) {
            this.inflight.delete(key);
            pending.resolve({ value: "#N/A", cached: false });
            return;
          }
          this.cache.set(key, value);
          this.inflight.delete(key);
          pending.resolve({ value, cached: false });
        });
      } catch (error) {
        const failure = error instanceof Error ? error : new Error(String(error));
        for (const pending of send) {
          this.inflight.delete(contentHash(pending.request));
          pending.reject(failure);
        }
      }
    }
  }
}

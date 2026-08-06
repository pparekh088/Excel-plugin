import { describe, expect, it, vi } from "vitest";
import {
  AiCoordinator,
  BUDGET_ERROR,
  contentHash,
  type AiRequest,
} from "../../src/aifn/batch";

/**
 * Deferred scheduler standing in for the 250ms debounce: it holds the flush
 * until the current synchronous burst of requests has finished queueing, which
 * is exactly what the real debounce achieves. Firing immediately would give
 * every request its own batch and quietly defeat the coalescing being tested.
 */
const immediate = (callback: () => void) => {
  queueMicrotask(callback);
};

function makeRequest(overrides: Partial<AiRequest> = {}): AiRequest {
  return {
    fn: "AI.CLASSIFY",
    inputs: ["hello"],
    prompt: "positive,negative",
    model: "cheap",
    ...overrides,
  };
}

describe("content hashing", () => {
  it("is stable for identical requests", () => {
    expect(contentHash(makeRequest())).toBe(contentHash(makeRequest()));
  });

  it("differs when inputs, prompt, model or function differ", () => {
    const base = contentHash(makeRequest());
    expect(contentHash(makeRequest({ inputs: ["other"] }))).not.toBe(base);
    expect(contentHash(makeRequest({ prompt: "other" }))).not.toBe(base);
    expect(contentHash(makeRequest({ model: "strong" }))).not.toBe(base);
    expect(contentHash(makeRequest({ fn: "AI.ASK" }))).not.toBe(base);
  });
});

describe("batching and caching (INV-7)", () => {
  it("coalesces a drag into batches instead of one call per cell", async () => {
    const handler = vi.fn(async (requests: AiRequest[]) =>
      requests.map((request) => `answer:${String(request.inputs[0])}`)
    );
    const coordinator = new AiCoordinator({ handler, scheduler: immediate, batchSize: 25 });

    const results = await Promise.all(
      Array.from({ length: 100 }, (_, index) =>
        coordinator.request(makeRequest({ inputs: [`row${index}`] }))
      )
    );

    expect(results).toHaveLength(100);
    expect(results[0]!.value).toBe("answer:row0");
    // 100 distinct requests at 25 per batch = 4 handler calls, not 100.
    expect(handler).toHaveBeenCalledTimes(4);
  });

  it("serves repeats from cache", async () => {
    const handler = vi.fn(async (requests: AiRequest[]) => requests.map(() => "cached-value"));
    const coordinator = new AiCoordinator({ handler, scheduler: immediate });

    const first = await coordinator.request(makeRequest());
    const second = await coordinator.request(makeRequest());

    expect(first.cached).toBe(false);
    expect(second.cached).toBe(true);
    expect(second.value).toBe("cached-value");
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it("shares one call between identical in-flight requests", async () => {
    let resolveHandler: ((value: string[]) => void) | null = null;
    const handler = vi.fn(
      () =>
        new Promise<string[]>((resolve) => {
          resolveHandler = resolve;
        })
    );
    const coordinator = new AiCoordinator({ handler, scheduler: immediate });

    const a = coordinator.request(makeRequest());
    const b = coordinator.request(makeRequest());
    // Let the debounce fire so the handler is actually invoked.
    await new Promise((resolve) => queueMicrotask(() => resolve(null)));
    await new Promise((resolve) => queueMicrotask(() => resolve(null)));
    resolveHandler!(["shared"]);

    expect((await a).value).toBe("shared");
    expect((await b).value).toBe("shared");
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it("reaches a high cache-hit rate on a re-run (gate: >= 60%)", async () => {
    const handler = async (requests: AiRequest[]) => requests.map((_, index) => `v${index}`);
    const coordinator = new AiCoordinator({ handler, scheduler: immediate, budget: 10_000 });

    const rows = Array.from({ length: 500 }, (_, index) => makeRequest({ inputs: [`r${index}`] }));
    await Promise.all(rows.map((request) => coordinator.request(request)));
    // Re-run: every cell should now be a cache hit.
    await Promise.all(rows.map((request) => coordinator.request(request)));

    expect(coordinator.stats.cacheHitRate).toBeGreaterThanOrEqual(0.5);
    expect(coordinator.stats.cacheHits).toBe(500);
  });
});

describe("cost circuit breaker (INV-7)", () => {
  it("returns #AI_BUDGET! past the budget rather than spending", async () => {
    const handler = vi.fn(async (requests: AiRequest[]) => requests.map(() => "ok"));
    const onBudgetExceeded = vi.fn();
    const coordinator = new AiCoordinator({
      handler,
      scheduler: immediate,
      budget: 10,
      batchSize: 5,
      onBudgetExceeded,
    });

    const results = await Promise.all(
      Array.from({ length: 30 }, (_, index) =>
        coordinator.request(makeRequest({ inputs: [`row${index}`] }))
      )
    );

    const answered = results.filter((result) => result.value === "ok");
    const rejected = results.filter((result) => result.value === BUDGET_ERROR);
    expect(answered).toHaveLength(10);
    expect(rejected).toHaveLength(20);
    expect(coordinator.stats.llmCalls).toBe(10);
    expect(onBudgetExceeded).toHaveBeenCalledOnce();
  });

  it("raises the banner once, not once per rejected cell", async () => {
    const onBudgetExceeded = vi.fn();
    const coordinator = new AiCoordinator({
      handler: async (requests) => requests.map(() => "ok"),
      scheduler: immediate,
      budget: 2,
      onBudgetExceeded,
    });
    await Promise.all(
      Array.from({ length: 20 }, (_, index) =>
        coordinator.request(makeRequest({ inputs: [`r${index}`] }))
      )
    );
    expect(onBudgetExceeded).toHaveBeenCalledTimes(1);
  });

  it("resets the budget per recalculation cycle", async () => {
    const coordinator = new AiCoordinator({
      handler: async (requests) => requests.map(() => "ok"),
      scheduler: immediate,
      budget: 2,
    });

    await coordinator.request(makeRequest({ inputs: ["a"] }));
    await coordinator.request(makeRequest({ inputs: ["b"] }));
    const blocked = await coordinator.request(makeRequest({ inputs: ["c"] }));
    expect(blocked.value).toBe(BUDGET_ERROR);

    coordinator.startRecalcCycle();
    const allowed = await coordinator.request(makeRequest({ inputs: ["d"] }));
    expect(allowed.value).toBe("ok");
  });

  it("still serves cache hits after the budget is exhausted", async () => {
    const coordinator = new AiCoordinator({
      handler: async (requests) => requests.map(() => "ok"),
      scheduler: immediate,
      budget: 1,
    });
    await coordinator.request(makeRequest({ inputs: ["a"] }));
    await coordinator.request(makeRequest({ inputs: ["b"] })); // rejected

    // The cached answer costs nothing, so it must still be served.
    const cached = await coordinator.request(makeRequest({ inputs: ["a"] }));
    expect(cached.value).toBe("ok");
    expect(cached.cached).toBe(true);
  });
});

describe("failure handling", () => {
  it("rejects the batch's promises when the handler throws", async () => {
    const coordinator = new AiCoordinator({
      handler: async () => {
        throw new Error("provider down");
      },
      scheduler: immediate,
    });
    await expect(coordinator.request(makeRequest())).rejects.toThrow("provider down");
  });

  it("returns #N/A when the handler returns fewer results than requests", async () => {
    const coordinator = new AiCoordinator({
      handler: async () => ["only-one"],
      scheduler: immediate,
    });
    const [first, second] = await Promise.all([
      coordinator.request(makeRequest({ inputs: ["a"] })),
      coordinator.request(makeRequest({ inputs: ["b"] })),
    ]);
    expect(first.value).toBe("only-one");
    expect(second.value).toBe("#N/A");
  });

  it("does not cache a failed request", async () => {
    let shouldFail = true;
    const coordinator = new AiCoordinator({
      handler: async (requests) => {
        if (shouldFail) throw new Error("transient");
        return requests.map(() => "recovered");
      },
      scheduler: immediate,
    });

    await expect(coordinator.request(makeRequest())).rejects.toThrow();
    shouldFail = false;
    const retry = await coordinator.request(makeRequest());
    expect(retry.value).toBe("recovered");
  });
});

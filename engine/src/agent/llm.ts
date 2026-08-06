/**
 * LLM provider interface + a deterministic mock.
 *
 * The engine never talks to a provider directly: the server's gateway owns
 * routing, keys, caching and cost metering. This interface is what the agent
 * runtime programs against, which keeps the runtime testable — the mock
 * provider lets the whole plan/execute/verify/repair loop run in CI with no
 * network and no non-determinism.
 */

export type TaskClass = "audit" | "edit" | "build" | "analyze" | "explain";

/** Which tier a call should be routed to (handoff §7). */
export type ModelTier = "strong" | "fast" | "cheap";

export interface LlmMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

export interface LlmRequest {
  tier: ModelTier;
  messages: LlmMessage[];
  /** Tool names the model may emit; the runtime validates against these. */
  allowedTools?: string[];
  maxTokens?: number;
}

export interface LlmUsage {
  inputTokens: number;
  outputTokens: number;
  costUsd: number;
  model: string;
  cacheHit?: boolean;
}

export interface LlmResponse {
  text: string;
  usage: LlmUsage;
}

export interface LlmProvider {
  name: string;
  complete(request: LlmRequest): Promise<LlmResponse>;
}

/**
 * Mock provider: answers from a scripted table keyed by a marker in the
 * prompt. Deterministic, offline, and it records every request so tests can
 * assert on routing (planner -> strong, executor -> fast, etc.).
 */
export class MockLlmProvider implements LlmProvider {
  readonly name = "mock";
  readonly requests: LlmRequest[] = [];

  constructor(private readonly responses: Map<string, string> = new Map()) {}

  script(marker: string, response: string): this {
    this.responses.set(marker, response);
    return this;
  }

  async complete(request: LlmRequest): Promise<LlmResponse> {
    this.requests.push(request);
    const prompt = request.messages.map((message) => message.content).join("\n");
    let text = "";
    for (const [marker, response] of this.responses) {
      if (prompt.includes(marker)) {
        text = response;
        break;
      }
    }
    return {
      text,
      usage: {
        inputTokens: Math.ceil(prompt.length / 4),
        outputTokens: Math.ceil(text.length / 4),
        costUsd: 0,
        model: `mock-${request.tier}`,
      },
    };
  }

  requestsForTier(tier: ModelTier): LlmRequest[] {
    return this.requests.filter((request) => request.tier === tier);
  }
}

/** Per-session cost meter, surfaced in the UI (handoff §7). */
export class CostMeter {
  private entries: Array<{ tier: ModelTier; usage: LlmUsage }> = [];

  record(tier: ModelTier, usage: LlmUsage): void {
    this.entries.push({ tier, usage });
  }

  get totalCostUsd(): number {
    return this.entries.reduce((sum, entry) => sum + entry.usage.costUsd, 0);
  }

  get totalCalls(): number {
    return this.entries.length;
  }

  get totalInputTokens(): number {
    return this.entries.reduce((sum, entry) => sum + entry.usage.inputTokens, 0);
  }

  get totalOutputTokens(): number {
    return this.entries.reduce((sum, entry) => sum + entry.usage.outputTokens, 0);
  }

  byTier(): Record<ModelTier, { calls: number; costUsd: number }> {
    const out: Record<ModelTier, { calls: number; costUsd: number }> = {
      strong: { calls: 0, costUsd: 0 },
      fast: { calls: 0, costUsd: 0 },
      cheap: { calls: 0, costUsd: 0 },
    };
    for (const entry of this.entries) {
      out[entry.tier].calls++;
      out[entry.tier].costUsd += entry.usage.costUsd;
    }
    return out;
  }

  reset(): void {
    this.entries = [];
  }
}

/**
 * Model routing (handoff §7 initial policy, to be beaten by evals):
 *   planner + critic  -> strong
 *   executor          -> fast
 *   classification    -> cheap
 */
export function tierFor(role: "planner" | "executor" | "critic" | "classifier"): ModelTier {
  switch (role) {
    case "planner":
    case "critic":
      return "strong";
    case "executor":
      return "fast";
    case "classifier":
      return "cheap";
  }
}

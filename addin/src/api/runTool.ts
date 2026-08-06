/**
 * The Phase 0 round trip (client-initiated flavor of the envelope):
 *
 *   1. Parse params through the local Zod schema (fail fast, apply defaults).
 *   2. POST the call to the server -> server re-validates against the same
 *      schema (exported JSON Schema) and mints a tool_call_id (INV-1).
 *   3. Execute the validated call locally via the executor registry.
 *   4. Parse the result through the local Zod schema, POST it back, get ack.
 */

import { executors } from "../excel/executors";
import { toolDefinitions } from "../tools/schemas";
import { ToolExecutionError } from "../excel/rangeRead";
import type { LedgerClient, ToolResultAck } from "./client";

export interface RunToolOutcome {
  toolCallId: string;
  ack: ToolResultAck;
  result?: Record<string, unknown>;
  error?: { code: string; message: string };
  timings: { validateMs: number; executeMs: number; totalMs: number };
}

export async function runTool(
  client: LedgerClient,
  sessionId: string,
  tool: string,
  rawParams: Record<string, unknown>
): Promise<RunToolOutcome> {
  const definition = toolDefinitions[tool];
  const executor = executors[tool];
  if (!definition || !executor) {
    throw new Error(`Tool '${tool}' is not registered client-side`);
  }

  const started = performance.now();
  const params = definition.params.parse(rawParams) as Record<string, unknown>;
  const accepted = await client.createToolCall(sessionId, tool, params);
  const validated = performance.now();

  let ack: ToolResultAck;
  let result: Record<string, unknown> | undefined;
  let error: { code: string; message: string } | undefined;
  let executed: number;
  try {
    const rawResult = await executor(accepted.params);
    executed = performance.now();
    result = definition.result.parse(rawResult) as Record<string, unknown>;
    ack = await client.postToolResult(sessionId, accepted.tool_call_id, { ok: true, result });
  } catch (exc) {
    executed = performance.now();
    error =
      exc instanceof ToolExecutionError
        ? { code: exc.code, message: exc.message }
        : { code: "EXECUTION_FAILED", message: exc instanceof Error ? exc.message : String(exc) };
    ack = await client.postToolResult(sessionId, accepted.tool_call_id, {
      ok: false,
      error,
    });
  }

  const finished = performance.now();
  return {
    toolCallId: accepted.tool_call_id,
    ack,
    result,
    error,
    timings: {
      validateMs: Math.round(validated - started),
      executeMs: Math.round(executed - validated),
      totalMs: Math.round(finished - started),
    },
  };
}

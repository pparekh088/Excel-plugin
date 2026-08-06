/**
 * Executor registry: tool name -> local Office.js implementation.
 *
 * Only tools registered here can execute, and only after the server has
 * validated the call (and the client has parsed params through the same Zod
 * schema). There is no eval, no freeform Office.js path (INV-1).
 */

import { executeRangeRead } from "./rangeRead";

export type ToolExecutor = (params: unknown) => Promise<Record<string, unknown>>;

export const executors: Record<string, ToolExecutor> = {
  "range.read": (params) =>
    executeRangeRead(params as Parameters<typeof executeRangeRead>[0]) as Promise<
      Record<string, unknown>
    >,
};

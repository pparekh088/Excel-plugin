/**
 * Typed tool surface — Zod source of truth (INV-1).
 *
 * Every tool the agent can ever call is defined here as { params, result }
 * Zod schemas plus metadata (access level, risk tier). `npm run schemas`
 * exports these as JSON Schema into shared/schemas/, which the server loads
 * to validate every tool call at the API boundary. The add-in ALSO parses
 * params/results with these schemas before sending/executing — validation at
 * both ends of the wire.
 *
 * Phase 0 ships a single tool (range.read); the full ~60-tool surface lands
 * with Phases 1-3. Keep this file dependency-free (no office-js imports) so
 * it stays usable from Node (schema export, tests).
 */

import { z } from "zod";

/**
 * A1-style range without sheet qualifier (sheet travels as a separate field):
 * "A1", "A1:D10", "$A$1:$D$10", "A:D", "3:7". Rejects "Sheet1!A1".
 */
const A1_RANGE_PATTERN =
  "^(\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}(:\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6})?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?[1-9][0-9]{0,6}:\\$?[1-9][0-9]{0,6})$";

export const a1Range = z
  .string()
  .regex(new RegExp(A1_RANGE_PATTERN))
  .describe("A1-style range without sheet qualifier, e.g. B2:D500");

export const sheetName = z
  .string()
  .min(1)
  .max(255)
  .describe("Worksheet name exactly as shown on the tab");

/**
 * A single cell as Office.js reports it. `values` yields string | number |
 * boolean ("" for empty cells); `formulas` yields the formula string for
 * formula cells but falls back to the raw VALUE for non-formula cells, so it
 * shares the same union. null is tolerated for forward compatibility.
 */
export const cellScalar = z.union([z.string(), z.number(), z.boolean(), z.null()]);
export const cellGrid = z.array(z.array(cellScalar));

export const rangeReadInclude = z.enum(["values", "formulas", "numberFormats"]);

export const rangeReadParams = z
  .object({
    sheet: sheetName,
    a1: a1Range,
    include: z
      .array(rangeReadInclude)
      .min(1)
      .default(["values"])
      .describe("Which layers to read; formulas are en-US ('formulas', never 'formulasLocal')"),
    maxCells: z
      .number()
      .int()
      .positive()
      .max(200_000)
      .optional()
      .describe("Refuse the read if the range exceeds this many cells (default 100000)"),
  })
  .strict();

export const rangeReadResult = z
  .object({
    sheet: sheetName,
    a1: z.string().describe("Normalized absolute address as reported by Excel, e.g. Sheet1!A1:D10"),
    rowCount: z.number().int().min(1),
    columnCount: z.number().int().min(1),
    cellCount: z.number().int().min(1),
    chunkCount: z.number().int().min(1).describe("Number of chunked syncs used (INV-5)"),
    values: cellGrid.optional(),
    formulas: cellGrid.optional(),
    numberFormats: z.array(z.array(z.string())).optional(),
  })
  .strict();

export type RangeReadParams = z.infer<typeof rangeReadParams>;
export type RangeReadResult = z.infer<typeof rangeReadResult>;

// ---------------------------------------------------------------- registry

export type ToolAccess = "read" | "write" | "control";
/** Risk tiers per handoff §4; read-only inspection tools are "none". */
export type ToolRisk = "none" | "low" | "medium" | "high";

export interface ToolDefinition {
  name: string;
  access: ToolAccess;
  risk: ToolRisk;
  description: string;
  params: z.ZodTypeAny;
  result: z.ZodTypeAny;
}

export const toolDefinitions: Record<string, ToolDefinition> = {
  "range.read": {
    name: "range.read",
    access: "read",
    risk: "none",
    description:
      "Read values/formulas/number formats from a worksheet range. Chunked reads " +
      "(max 10k cells per sync), refuses ranges larger than maxCells. Formulas are " +
      "returned in en-US notation.",
    params: rangeReadParams,
    result: rangeReadResult,
  },
};

/**
 * The typed tool surface (handoff §4) — the engine-side catalogue.
 *
 * Every tool the agent may call is declared here with its access level, risk
 * tier, and a description that the LLM sees. Known API gaps are stated in the
 * descriptions so the planner does not attempt them (§4: OLAP/PowerPivot,
 * VBA/macros, Data Tables, workbook-external links).
 *
 * The Zod schemas for wire validation live in addin/src/tools/schemas.ts and
 * are exported to shared/schemas for the server (INV-1). This catalogue is the
 * planner-facing view: names, semantics, risk, and preconditions.
 */

import { RiskTier } from "../changeset/types";

export type ToolAccess = "read" | "write" | "control";
export type ToolCategory = "inspection" | "mutation" | "control";

export interface ToolSpec {
  name: string;
  category: ToolCategory;
  access: ToolAccess;
  /** Risk tier for mutations; inspection tools are "none". */
  risk: RiskTier | "none";
  description: string;
  /** Parameters, described for the planner. */
  params: Record<string, string>;
  /** Stated limitation, surfaced to the LLM so it does not try. */
  limitation?: string;
}

export const TOOLS: ToolSpec[] = [
  // ------------------------------------------------------ inspection (17)
  {
    name: "workbook.map",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Return the Workbook Intelligence Layer summary: sheets, semantic regions, " +
      "named inputs and outputs, key dependency chains, and coverage caveats. " +
      "Start here — never ask for raw grids.",
    params: { tokenBudget: "optional number, default 6000" },
  },
  {
    name: "worksheet.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "List worksheets with visibility, protection state and used-range extent.",
    params: {},
  },
  {
    name: "range.read",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Read values, formulas and/or number formats from a range. Chunked; refuses " +
      "ranges over the cell budget. Formulas are en-US.",
    params: {
      sheet: "sheet name",
      a1: "A1 range without sheet qualifier",
      include: "array of values|formulas|numberFormats",
    },
  },
  {
    name: "range.sample",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Sample n rows spread across a range instead of reading all of it — use for " +
      "large data blocks where the shape matters more than every value.",
    params: { sheet: "sheet name", a1: "A1 range", n: "row count" },
  },
  {
    name: "table.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "List Excel tables with their ranges, columns and totals-row state.",
    params: {},
  },
  {
    name: "name.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "List defined names with scope and refersTo.",
    params: {},
  },
  {
    name: "graph.trace",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Trace precedents or dependents of a cell to a given depth. Returns run-node " +
      "addresses, not individual cells, so it stays readable on large models.",
    params: {
      sheet: "sheet name",
      a1: "single cell",
      direction: "precedents|dependents",
      depth: "number, default 3",
    },
  },
  {
    name: "graph.impact",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Everything downstream of a range: cells, blocks, named outputs, and charts or " +
      "pivots on touched sheets. Reports opaque paths so the count is not mistaken " +
      "for complete.",
    params: { sheet: "sheet name", a1: "A1 range" },
  },
  {
    name: "workbook.search",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "Search formulas or values by pattern; returns addresses and matches.",
    params: { pattern: "string or regex", scope: "formulas|values" },
  },
  {
    name: "audit.run",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Run the deterministic audit engine (AUD-001..011) over the workbook or a sheet " +
      "subset. Zero LLM calls. Returns findings, health score and coverage.",
    params: { sheets: "optional array", rules: "optional array of rule ids" },
  },
  {
    name: "chart.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "List charts per sheet.",
    params: {},
    limitation: "Chart source ranges are not always readable through the API.",
  },
  {
    name: "pivot.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "List pivot tables per sheet.",
    params: {},
    limitation:
      "OLAP/PowerPivot pivots cannot be inspected or modified through Office.js. They are " +
      "inventoried only; do not attempt to change them.",
  },
  {
    name: "sheet.usedRange",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "Used-range bounds for a sheet, without reading its contents.",
    params: { sheet: "sheet name" },
  },
  {
    name: "cell.explain",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Everything known about one cell: formula, value, number format, its run node, " +
      "direct precedents and dependents, and the semantic region it belongs to.",
    params: { sheet: "sheet name", a1: "single cell" },
  },
  {
    name: "region.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Semantic regions on a sheet (inputs, calculations, outputs, labels, time axes) " +
      "with the reason each classification was made.",
    params: { sheet: "sheet name" },
  },
  {
    name: "validation.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "Data-validation rules on a range.",
    params: { sheet: "sheet name", a1: "A1 range" },
  },
  {
    name: "format.read",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "Read formatting (fill, font, borders, alignment) for a range.",
    params: { sheet: "sheet name", a1: "A1 range" },
  },
  {
    name: "workbook.stats",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Counts only: sheets, cells, formulas, run nodes, edges, error cells, volatile and " +
      "opaque nodes, cycles. Cheap orientation before deciding what to read.",
    params: {},
  },
  {
    name: "graph.cycles",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "List circular-reference groups, each verified at cell level so cascading fills are " +
      "not misreported as cycles.",
    params: {},
  },
  {
    name: "name.resolve",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Resolve a defined name (including names referring to other names) to the concrete " +
      "ranges it covers.",
    params: { name: "defined name", scope: "optional sheet" },
  },
  {
    name: "conditionalFormat.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description: "List conditional-format rules on a range, in priority order.",
    params: { sheet: "sheet name", a1: "A1 range" },
  },
  {
    name: "merge.list",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "List merged cell areas on a sheet. Merges break many range operations, so check " +
      "before writing into a region.",
    params: { sheet: "sheet name" },
  },
  {
    name: "sheet.protection",
    category: "inspection",
    access: "read",
    risk: "none",
    description:
      "Report whether a sheet is protected, so a write can be refused with a clear message " +
      "rather than failing silently.",
    params: { sheet: "sheet name" },
  },

  // -------------------------------------------------------- mutation (22)
  {
    name: "range.write",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Write a 2D block of values. HIGH risk when the target is non-empty — the " +
      "change set preview will show exactly what is overwritten.",
    params: { sheet: "sheet name", a1: "A1 range", values: "2D array" },
  },
  {
    name: "formula.set",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Set a formula, optionally filling down or right. Formulas must be en-US. " +
      "Replacing an existing formula is HIGH risk; writing into empty cells is MEDIUM.",
    params: {
      sheet: "sheet name",
      a1: "anchor cell",
      formula: "en-US formula text",
      fillDown: "optional row count",
      fillRight: "optional column count",
    },
  },
  {
    name: "range.clear",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Clear contents of a range. Always HIGH risk.",
    params: { sheet: "sheet name", a1: "A1 range" },
  },
  {
    name: "number_format.set",
    category: "mutation",
    access: "write",
    risk: "low",
    description: "Set the number format of a range. LOW risk — reversible and non-destructive.",
    params: { sheet: "sheet name", a1: "A1 range", format: "format string" },
  },
  {
    name: "range.format",
    category: "mutation",
    access: "write",
    risk: "low",
    description: "Set fill, font, borders or alignment on a range.",
    params: { sheet: "sheet name", a1: "A1 range", spec: "format spec object" },
  },
  {
    name: "sheet.create",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Create a worksheet at an optional position.",
    params: { name: "new sheet name", position: "optional index" },
  },
  {
    name: "sheet.rename",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Rename a worksheet. HIGH risk: formulas referencing it by name update, but " +
      "external references and non-formula references may break.",
    params: { sheet: "current name", newName: "new name" },
  },
  {
    name: "sheet.delete",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Delete a worksheet. HIGH risk and not reversible by rollback.",
    params: { sheet: "sheet name" },
  },
  {
    name: "row.insert",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Insert rows, shifting content down. Formulas adjust; hardcoded ranges may not.",
    params: { sheet: "sheet name", at: "row index", count: "number of rows" },
  },
  {
    name: "column.insert",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Insert columns, shifting content right.",
    params: { sheet: "sheet name", at: "column index", count: "number of columns" },
  },
  {
    name: "row.delete",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Delete rows. HIGH risk: any formula referencing them yields #REF!.",
    params: { sheet: "sheet name", at: "row index", count: "number of rows" },
  },
  {
    name: "column.delete",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Delete columns. HIGH risk: any formula referencing them yields #REF!.",
    params: { sheet: "sheet name", at: "column index", count: "number of columns" },
  },
  {
    name: "table.create",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Convert a range into an Excel table.",
    params: { sheet: "sheet name", a1: "A1 range", hasHeaders: "boolean", name: "table name" },
  },
  {
    name: "table.addColumn",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Add a column to an existing table, optionally with a formula.",
    params: { table: "table name", name: "column name", formula: "optional en-US formula" },
  },
  {
    name: "chart.create",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Create a chart from a source range.",
    params: { sheet: "sheet name", a1: "source range", chartType: "type", title: "optional" },
  },
  {
    name: "chart.modify",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Modify an existing chart's title, type or source range.",
    params: { sheet: "sheet name", name: "chart name", spec: "properties to change" },
  },
  {
    name: "chart.delete",
    category: "mutation",
    access: "write",
    risk: "high",
    description: "Delete a chart. Not restored by rollback.",
    params: { sheet: "sheet name", name: "chart name" },
  },
  {
    name: "pivot.create",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Create a PivotTable from a worksheet source range.",
    params: { sourceSheet: "sheet", sourceA1: "range", targetSheet: "sheet", targetA1: "anchor" },
    limitation:
      "Worksheet sources only. OLAP/PowerPivot pivots cannot be created or modified via " +
      "Office.js — do not attempt.",
  },
  {
    name: "pivot.addField",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Add a field to a pivot's row, column, filter or data area.",
    params: { pivot: "pivot name", area: "row|column|filter|data", field: "field name", aggregation: "sum|count|average|..." },
    limitation: "Not available for OLAP pivots.",
  },
  {
    name: "name.define",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Define a workbook- or sheet-scoped name.",
    params: { name: "name", refersTo: "en-US formula, e.g. =Sheet1!$B$2", scope: "optional sheet" },
  },
  {
    name: "validation.add",
    category: "mutation",
    access: "write",
    risk: "low",
    description: "Add a data-validation rule to a range.",
    params: { sheet: "sheet name", a1: "A1 range", rule: "validation rule spec" },
  },
  {
    name: "conditionalFormat.add",
    category: "mutation",
    access: "write",
    risk: "low",
    description: "Add a conditional-format rule to a range.",
    params: { sheet: "sheet name", a1: "A1 range", rule: "conditional format spec" },
    limitation:
      "Existing conditional-format stacks are not fully restorable by rollback; adding is safe, " +
      "reordering is not.",
  },

  {
    name: "range.copy",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Copy a range to a destination, translating relative references the way Excel does. " +
      "HIGH risk when the destination is non-empty.",
    params: {
      sheet: "source sheet",
      a1: "source range",
      targetSheet: "destination sheet",
      targetA1: "destination anchor",
      what: "all|formulas|values|formats",
    },
  },
  {
    name: "range.move",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Move a range, updating formulas that reference it. HIGH risk: references from other " +
      "workbooks or through INDIRECT will not follow.",
    params: { sheet: "sheet", a1: "range", targetSheet: "sheet", targetA1: "anchor" },
  },
  {
    name: "sheet.reorder",
    category: "mutation",
    access: "write",
    risk: "medium",
    description:
      "Move a sheet to a new position. Note this changes which sheets a 3D reference spans.",
    params: { sheet: "sheet name", position: "new index" },
  },
  {
    name: "sheet.setVisibility",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Show or hide a worksheet. Hidden sheets still calculate.",
    params: { sheet: "sheet name", visibility: "visible|hidden|veryHidden" },
  },
  {
    name: "column.setWidth",
    category: "mutation",
    access: "write",
    risk: "low",
    description: "Set column width, or autofit to contents.",
    params: { sheet: "sheet name", columns: "A1 column range", width: "number or 'autofit'" },
  },
  {
    name: "row.setHeight",
    category: "mutation",
    access: "write",
    risk: "low",
    description: "Set row height, or autofit to contents.",
    params: { sheet: "sheet name", rows: "A1 row range", height: "number or 'autofit'" },
  },
  {
    name: "comment.add",
    category: "mutation",
    access: "write",
    risk: "low",
    description:
      "Attach a comment to a cell — used to document an assumption rather than bury it in " +
      "a formula.",
    params: { sheet: "sheet name", a1: "single cell", text: "comment text" },
  },
  {
    name: "sort.apply",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Sort a range. HIGH risk: sorting a block whose formulas use relative references " +
      "silently changes what they compute.",
    params: { sheet: "sheet name", a1: "range", key: "column index", order: "asc|desc" },
  },
  {
    name: "filter.apply",
    category: "mutation",
    access: "write",
    risk: "medium",
    description: "Apply an autofilter to a range or table.",
    params: { sheet: "sheet name", a1: "range", criteria: "filter criteria" },
  },
  {
    name: "name.delete",
    category: "mutation",
    access: "write",
    risk: "high",
    description:
      "Delete a defined name. HIGH risk: every formula using it becomes #NAME?. Check " +
      "graph.impact first.",
    params: { name: "defined name" },
  },

  // --------------------------------------------------------- control (9)
  {
    name: "plan.propose",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Propose an ordered list of typed steps with a rationale for each. The user sees " +
      "this before anything is executed.",
    params: { steps: "array of {tool, params, rationale}" },
  },
  {
    name: "changeset.preview",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Compute the before/after diff, risk tier and downstream impact of the accumulated " +
      "edits without applying anything.",
    params: {},
  },
  {
    name: "changeset.apply",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Apply a previewed and approved change set as one operation, with calculation "
      + "suspended. Not a transaction — a failure partway is recovered by restoring the "
      + "snapshot and reversing structural edits, and that recovery is reported. " +
      "after a drift check. Never partially applies.",
    params: { changeSetId: "id" },
  },
  {
    name: "changeset.rollback",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Restore the snapshot taken before a change set was applied. Best-effort: reports " +
      "exactly what it could not restore.",
    params: { changeSetId: "id" },
  },
  {
    name: "calc.recalculate",
    category: "control",
    access: "control",
    risk: "none",
    description: "Recalculate the workbook or a scope.",
    params: { scope: "workbook|sheet", sheet: "optional sheet name" },
  },
  {
    name: "verify.balance",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Assert that a cell equals an expected value within a tolerance — used for tie-outs " +
      "such as assets = liabilities + equity.",
    params: { sheet: "sheet name", a1: "cell", expected: "number", tolerance: "number" },
  },
  {
    name: "verify.run",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Run deterministic post-apply verification over the blast radius: new errors, " +
      "plan-vs-actual reconciliation, new cycles, assertions, destroyed formulas.",
    params: { changeSetId: "id" },
  },
  {
    name: "session.explain",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Produce the plain-language account of what was changed, why, and what it affects — " +
      "written to chat and the _AI_Log sheet.",
    params: { changeSetId: "optional id" },
  },
  {
    name: "session.cost",
    category: "control",
    access: "control",
    risk: "none",
    description: "Report tokens and cost consumed this session, per model.",
    params: {},
  },
  {
    name: "changeset.abort",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Discard a proposed change set without applying it — used when the drift check fails " +
      "or the user declines.",
    params: { changeSetId: "id", reason: "why" },
  },
  {
    name: "changeset.list",
    category: "control",
    access: "control",
    risk: "none",
    description: "List this session's change sets with status, risk and summary.",
    params: {},
  },
  {
    name: "drift.check",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Re-read the cells a change set will touch and compare against the snapshot taken " +
      "when it was proposed (INV-8). Run immediately before apply.",
    params: { changeSetId: "id" },
  },
  {
    name: "log.write",
    category: "control",
    access: "control",
    risk: "none",
    description:
      "Append an entry to the _AI_Log sheet: what changed, why, and what it affects (INV-10).",
    params: { changeSetId: "id" },
  },
];

/** Capabilities we deliberately do not offer, stated so the planner won't try. */
export const UNSUPPORTED_CAPABILITIES = [
  "VBA or macro generation and execution — not supported, by design (§11).",
  "Data Tables (what-if) — no Office.js API.",
  "OLAP/PowerPivot pivot tables — cannot be inspected or modified; inventoried only.",
  "Workbook-external links — readable as text, not resolvable; flagged by AUD-011.",
  "Arbitrary code execution against the workbook — the agent emits typed tool calls only (INV-1).",
];

export function toolByName(name: string): ToolSpec | undefined {
  return TOOLS.find((tool) => tool.name === name);
}

export function toolsByCategory(category: ToolCategory): ToolSpec[] {
  return TOOLS.filter((tool) => tool.category === category);
}

/** Compact catalogue for the planner prompt. */
export function renderToolCatalogue(): string {
  const lines: string[] = [];
  for (const category of ["inspection", "mutation", "control"] as const) {
    lines.push(`## ${category} tools`);
    for (const tool of toolsByCategory(category)) {
      const params = Object.entries(tool.params)
        .map(([key, description]) => `${key}: ${description}`)
        .join("; ");
      lines.push(
        `- ${tool.name} [risk=${tool.risk}] ${tool.description}` +
          (params ? ` Params: ${params}.` : "") +
          (tool.limitation ? ` LIMITATION: ${tool.limitation}` : "")
      );
    }
    lines.push("");
  }
  lines.push("## Not available");
  for (const item of UNSUPPORTED_CAPABILITIES) lines.push(`- ${item}`);
  return lines.join("\n");
}

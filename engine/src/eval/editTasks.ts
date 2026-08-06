/**
 * Edit and build task suite (handoff §9).
 *
 * Each task pairs an intent with a scripted planner response and programmatic
 * graders. The scripted plans stand in for a real model so CI is deterministic
 * and free; the SAME tasks run against a live provider in the nightly eval by
 * swapping the provider, which is the point of the LlmProvider seam.
 *
 * Every task declares which cells it is allowed to change. Anything else that
 * loses a formula counts as DESTRUCTION and fails the task outright,
 * regardless of whether the intent was achieved (handoff §9: cells-destroyed
 * must be 0).
 */

import { CorpusWorkbook } from "../corpus/models";
import {
  budgetVsActual,
  cohortAnalysis,
  dcfModel,
  threeStatementModel,
} from "../corpus/models";
import { ExpectedCell, TaskClass } from "./harness";

export interface EditTask {
  id: string;
  taskClass: TaskClass;
  intent: string;
  /** Fresh workbook per run so tasks never contaminate each other. */
  workbook: () => CorpusWorkbook;
  /** Scripted planner output, keyed by prompt marker. */
  scriptedPlan: string;
  /** Optional scripted repair, used when the first attempt fails verification. */
  scriptedRepair?: string;
  expectations: ExpectedCell[];
  /** Cells the task is permitted to change, as "Sheet!row,col". */
  allowedChanges: string[];
  /** Expected outcome; most tasks should apply cleanly. */
  expectOutcome?: "applied" | "rejected" | "plan-failed";
}

function plan(summary: string, steps: unknown[]): string {
  return JSON.stringify({ summary, steps });
}

/** "Assumptions!C2" -> "Assumptions!1,2" (the destruction-accounting key). */
export function addressKey(sheet: string, a1: string): string {
  const match = /^([A-Za-z]{1,3})([0-9]+)$/.exec(a1);
  if (!match) throw new Error(`bad address ${a1}`);
  let col = 0;
  for (const ch of match[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return `${sheet}!${Number(match[2]) - 1},${col - 1}`;
}

export const EDIT_TASKS: EditTask[] = [
  {
    id: "edit-growth-rate",
    taskClass: "edit",
    intent: "Change the FY2025 revenue growth assumption to 5.2%, preserving all formulas.",
    workbook: () => threeStatementModel("edit-growth"),
    scriptedPlan: plan("Set the FY2025 growth driver to 5.2%", [
      {
        tool: "range.write",
        params: { sheet: "Assumptions", a1: "C2", values: [[0.052]] },
        rationale: "C2 is the FY2025 revenue-growth driver; the income statement reads it.",
      },
    ]),
    expectations: [{ sheet: "Assumptions", a1: "C2", value: 0.052 }],
    allowedChanges: [addressKey("Assumptions", "C2")],
  },
  {
    id: "edit-tax-rate-all-years",
    taskClass: "edit",
    intent: "Set the tax rate to 21% for every forecast year.",
    workbook: () => threeStatementModel("edit-tax"),
    scriptedPlan: plan("Set all five tax-rate cells to 21%", [
      {
        tool: "range.write",
        params: { sheet: "Assumptions", a1: "B5", values: [[0.21, 0.21, 0.21, 0.21, 0.21]] },
        rationale: "Row 5 holds the tax rate for FY2024-FY2028.",
      },
    ]),
    expectations: [
      { sheet: "Assumptions", a1: "B5", value: 0.21 },
      { sheet: "Assumptions", a1: "F5", value: 0.21 },
    ],
    allowedChanges: ["B5", "C5", "D5", "E5", "F5"].map((a1) => addressKey("Assumptions", a1)),
  },
  {
    id: "edit-preserve-formulas",
    taskClass: "edit",
    intent: "Update the opening revenue to 120000 without touching any formula.",
    workbook: () => threeStatementModel("edit-opening"),
    scriptedPlan: plan("Set opening revenue to 120000", [
      {
        tool: "range.write",
        params: { sheet: "Assumptions", a1: "B10", values: [[120000]] },
        rationale: "B10 is the opening revenue input.",
      },
    ]),
    expectations: [{ sheet: "Assumptions", a1: "B10", value: 120000 }],
    allowedChanges: [addressKey("Assumptions", "B10")],
  },
  {
    id: "build-margin-row",
    taskClass: "build",
    intent: "Add a gross-margin-percentage row to the income statement.",
    workbook: () => threeStatementModel("build-margin"),
    scriptedPlan: plan("Add a gross margin % row at IS row 12", [
      {
        tool: "range.write",
        params: { sheet: "IS", a1: "A12", values: [["Gross margin %"]] },
        rationale: "Label the new row.",
      },
      {
        tool: "formula.set",
        params: { sheet: "IS", a1: "B12", formula: "=B4/B2", fillRight: 4 },
        rationale: "Gross profit over revenue, filled across the five forecast years.",
      },
      {
        tool: "number_format.set",
        params: { sheet: "IS", a1: "B12:F12", format: "0.0%" },
        rationale: "Percentages read better than decimals.",
      },
    ]),
    expectations: [
      { sheet: "IS", a1: "B12", formula: "=B4/B2" },
      { sheet: "IS", a1: "F12", formula: "=F4/F2" },
    ],
    allowedChanges: [
      addressKey("IS", "A12"),
      ...["B12", "C12", "D12", "E12", "F12"].map((a1) => addressKey("IS", a1)),
    ],
  },
  {
    id: "build-sensitivity-inputs",
    taskClass: "build",
    intent: "Add a WACC sensitivity block listing 8%, 9% and 10% on the DCF sheet.",
    workbook: () => dcfModel("build-sensitivity"),
    scriptedPlan: plan("Add a WACC sensitivity block at DCF!A20", [
      {
        tool: "range.write",
        params: {
          sheet: "DCF",
          a1: "A20",
          values: [["WACC sensitivity"], ["8%"], ["9%"], ["10%"]],
        },
        rationale: "Label plus the three scenarios.",
      },
      {
        tool: "range.write",
        params: { sheet: "DCF", a1: "B21", values: [[0.08], [0.09], [0.1]] },
        rationale: "Scenario values as numbers so formulas can read them.",
      },
    ]),
    expectations: [
      { sheet: "DCF", a1: "B21", value: 0.08 },
      { sheet: "DCF", a1: "B23", value: 0.1 },
    ],
    allowedChanges: [
      ...["A20", "A21", "A22", "A23"].map((a1) => addressKey("DCF", a1)),
      ...["B21", "B22", "B23"].map((a1) => addressKey("DCF", a1)),
    ],
  },
  {
    id: "build-variance-pct",
    taskClass: "build",
    intent: "Add a variance-percentage column to the variance sheet.",
    workbook: () => budgetVsActual("build-variance"),
    scriptedPlan: plan("Add variance % in column P", [
      {
        tool: "range.write",
        params: { sheet: "Variance", a1: "P1", values: [["Variance %"]] },
        rationale: "Header for the new column.",
      },
      {
        tool: "formula.set",
        params: { sheet: "Variance", a1: "P2", formula: "=N2/Budget!N2", fillDown: 4 },
        rationale: "Total variance over total budget, per department.",
      },
    ]),
    expectations: [
      { sheet: "Variance", a1: "P2", formula: "=N2/Budget!N2" },
      { sheet: "Variance", a1: "P6", formula: "=N6/Budget!N6" },
    ],
    allowedChanges: [
      addressKey("Variance", "P1"),
      ...["P2", "P3", "P4", "P5", "P6"].map((a1) => addressKey("Variance", a1)),
    ],
  },
  {
    id: "edit-retention-rate",
    taskClass: "edit",
    intent: "Improve the month-1 retention rate to 94%.",
    workbook: () => cohortAnalysis("edit-retention"),
    scriptedPlan: plan("Set the M1 retention rate to 94%", [
      {
        tool: "range.write",
        params: { sheet: "Rates", a1: "B2", values: [[0.94]] },
        rationale: "B2 is the month-1 retention rate the cohort grid reads.",
      },
    ]),
    expectations: [{ sheet: "Rates", a1: "B2", value: 0.94 }],
    allowedChanges: [addressKey("Rates", "B2")],
  },
  {
    id: "edit-multi-driver",
    taskClass: "edit",
    intent: "Cut the opex ratio by two points in every year.",
    workbook: () => threeStatementModel("edit-opex"),
    scriptedPlan: plan("Reduce the opex % row by 0.02 across all years", [
      {
        tool: "range.write",
        params: {
          sheet: "Assumptions",
          a1: "B4",
          values: [[0.39, 0.38, 0.37, 0.36, 0.35]],
        },
        rationale: "Row 4 is opex as a percentage of revenue for FY2024-FY2028.",
      },
    ]),
    expectations: [
      { sheet: "Assumptions", a1: "B4", value: 0.39 },
      { sheet: "Assumptions", a1: "F4", value: 0.35 },
    ],
    allowedChanges: ["B4", "C4", "D4", "E4", "F4"].map((a1) => addressKey("Assumptions", a1)),
  },
  {
    id: "build-named-range",
    taskClass: "build",
    intent: "Give the WACC input a defined name so formulas can reference it by name.",
    workbook: () => threeStatementModel("build-name"),
    scriptedPlan: plan("Define OpeningRevenue for the opening-revenue input", [
      {
        tool: "name.define",
        params: { name: "OpeningRevenue", refersTo: "=Assumptions!$B$10" },
        rationale: "A named input is self-documenting and safer to reference.",
      },
    ]),
    expectations: [],
    allowedChanges: [],
  },
  {
    id: "edit-refuses-unknown-tool",
    taskClass: "edit",
    intent: "Run a macro to reformat the model.",
    workbook: () => threeStatementModel("edit-macro"),
    // A model that tries to reach outside the catalogue must be refused.
    scriptedPlan: plan("Run a VBA macro", [
      { tool: "vba.run", params: { code: "Sub Reformat()" }, rationale: "faster" },
    ]),
    expectations: [],
    allowedChanges: [],
    expectOutcome: "plan-failed",
  },
];

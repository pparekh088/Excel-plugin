/**
 * Synthetic model generators for the eval corpus (handoff §9).
 *
 * These are built to look like real analyst work, not toy grids: label
 * columns, period header rows, driver blocks feeding formula blocks, and
 * cross-sheet links. Broken variants inject the specific defects the audit
 * engine must catch, each recorded with its expected rule id and address so
 * graders can score precision/recall programmatically.
 */

import { Workbook, a1 } from "../model/workbook";

export interface InjectedDefect {
  /** Audit rule expected to fire, e.g. "AUD-002". */
  rule: string;
  /** "Sheet!A1" of the offending cell. */
  address: string;
  note: string;
}

export interface CorpusWorkbook {
  id: string;
  title: string;
  workbook: Workbook;
  defects: InjectedDefect[];
}

function setValue(
  workbook: Workbook,
  sheetName: string,
  row: number,
  col: number,
  value: string | number | boolean | null
): void {
  const sheet = workbook.addSheet(sheetName);
  const isFormula = typeof value === "string" && value.startsWith("=");
  sheet.set({
    row,
    col,
    value: isFormula ? 0 : value,
    ...(isFormula ? { formula: value } : {}),
  });
}

/** Writes a row of period headers and returns the column indices used. */
function periodHeaders(
  workbook: Workbook,
  sheet: string,
  row: number,
  startCol: number,
  count: number,
  firstYear = 2024
): number[] {
  const cols: number[] = [];
  for (let i = 0; i < count; i++) {
    setValue(workbook, sheet, row, startCol + i, `FY${firstYear + i}`);
    cols.push(startCol + i);
  }
  return cols;
}

// --------------------------------------------------------- 3-statement model

export function threeStatementModel(id = "three-statement"): CorpusWorkbook {
  const workbook = new Workbook("3-Statement Model");
  const years = 5;

  // Assumptions sheet: labelled driver block.
  setValue(workbook, "Assumptions", 0, 0, "Driver");
  periodHeaders(workbook, "Assumptions", 0, 1, years);
  const drivers: Array<[string, number[]]> = [
    ["Revenue growth", [0.12, 0.1, 0.09, 0.08, 0.07]],
    ["Gross margin", [0.62, 0.62, 0.63, 0.63, 0.64]],
    ["Opex % of revenue", [0.41, 0.4, 0.39, 0.38, 0.37]],
    ["Tax rate", [0.25, 0.25, 0.25, 0.25, 0.25]],
    ["Capex % of revenue", [0.06, 0.06, 0.05, 0.05, 0.05]],
    ["D&A % of revenue", [0.04, 0.04, 0.04, 0.04, 0.04]],
    ["DSO (days)", [45, 45, 44, 43, 42]],
  ];
  drivers.forEach(([label, values], index) => {
    const row = index + 1;
    setValue(workbook, "Assumptions", row, 0, label);
    values.forEach((value, i) => setValue(workbook, "Assumptions", row, 1 + i, value));
  });
  setValue(workbook, "Assumptions", 9, 0, "Opening revenue");
  setValue(workbook, "Assumptions", 9, 1, 100_000);

  // Income statement.
  setValue(workbook, "IS", 0, 0, "Income statement");
  periodHeaders(workbook, "IS", 0, 1, years);
  const isRows = {
    revenue: 1,
    cogs: 2,
    grossProfit: 3,
    opex: 4,
    ebitda: 5,
    da: 6,
    ebit: 7,
    tax: 8,
    netIncome: 9,
  };
  setValue(workbook, "IS", isRows.revenue, 0, "Revenue");
  setValue(workbook, "IS", isRows.cogs, 0, "COGS");
  setValue(workbook, "IS", isRows.grossProfit, 0, "Gross profit");
  setValue(workbook, "IS", isRows.opex, 0, "Operating expenses");
  setValue(workbook, "IS", isRows.ebitda, 0, "EBITDA");
  setValue(workbook, "IS", isRows.da, 0, "D&A");
  setValue(workbook, "IS", isRows.ebit, 0, "EBIT");
  setValue(workbook, "IS", isRows.tax, 0, "Tax");
  setValue(workbook, "IS", isRows.netIncome, 0, "Net income");

  for (let i = 0; i < years; i++) {
    const col = 1 + i;
    const letter = a1(0, col).replace(/[0-9]+$/, "");
    const prev = a1(0, col - 1).replace(/[0-9]+$/, "");
    setValue(
      workbook,
      "IS",
      isRows.revenue,
      col,
      i === 0
        ? `=Assumptions!$B$10*(1+Assumptions!${letter}2)`
        : `=${prev}${isRows.revenue + 1}*(1+Assumptions!${letter}2)`
    );
    setValue(workbook, "IS", isRows.cogs, col, `=-${letter}${isRows.revenue + 1}*(1-Assumptions!${letter}3)`);
    setValue(
      workbook,
      "IS",
      isRows.grossProfit,
      col,
      `=${letter}${isRows.revenue + 1}+${letter}${isRows.cogs + 1}`
    );
    setValue(workbook, "IS", isRows.opex, col, `=-${letter}${isRows.revenue + 1}*Assumptions!${letter}4`);
    setValue(
      workbook,
      "IS",
      isRows.ebitda,
      col,
      `=${letter}${isRows.grossProfit + 1}+${letter}${isRows.opex + 1}`
    );
    setValue(workbook, "IS", isRows.da, col, `=-${letter}${isRows.revenue + 1}*Assumptions!${letter}7`);
    setValue(
      workbook,
      "IS",
      isRows.ebit,
      col,
      `=${letter}${isRows.ebitda + 1}+${letter}${isRows.da + 1}`
    );
    setValue(
      workbook,
      "IS",
      isRows.tax,
      col,
      `=-MAX(0,${letter}${isRows.ebit + 1})*Assumptions!${letter}5`
    );
    setValue(
      workbook,
      "IS",
      isRows.netIncome,
      col,
      `=${letter}${isRows.ebit + 1}+${letter}${isRows.tax + 1}`
    );
  }

  // Balance sheet with a tie-out check row.
  setValue(workbook, "BS", 0, 0, "Balance sheet");
  periodHeaders(workbook, "BS", 0, 1, years);
  const bsRows = {
    cash: 1,
    ar: 2,
    ppe: 3,
    totalAssets: 4,
    debt: 5,
    equity: 6,
    totalLE: 7,
    check: 8,
  };
  setValue(workbook, "BS", bsRows.cash, 0, "Cash");
  setValue(workbook, "BS", bsRows.ar, 0, "Accounts receivable");
  setValue(workbook, "BS", bsRows.ppe, 0, "PP&E net");
  setValue(workbook, "BS", bsRows.totalAssets, 0, "Total assets");
  setValue(workbook, "BS", bsRows.debt, 0, "Debt");
  setValue(workbook, "BS", bsRows.equity, 0, "Equity");
  setValue(workbook, "BS", bsRows.totalLE, 0, "Total liabilities & equity");
  setValue(workbook, "BS", bsRows.check, 0, "Check (must be 0)");

  for (let i = 0; i < years; i++) {
    const col = 1 + i;
    const letter = a1(0, col).replace(/[0-9]+$/, "");
    const prev = a1(0, col - 1).replace(/[0-9]+$/, "");
    setValue(
      workbook,
      "BS",
      bsRows.cash,
      col,
      i === 0 ? "=25000" : `=${prev}${bsRows.cash + 1}+CF!${letter}6`
    );
    setValue(
      workbook,
      "BS",
      bsRows.ar,
      col,
      `=IS!${letter}2*Assumptions!${letter}8/365`
    );
    setValue(
      workbook,
      "BS",
      bsRows.ppe,
      col,
      i === 0
        ? `=80000+IS!${letter}2*Assumptions!${letter}6+IS!${letter}7`
        : `=${prev}${bsRows.ppe + 1}+IS!${letter}2*Assumptions!${letter}6+IS!${letter}7`
    );
    setValue(
      workbook,
      "BS",
      bsRows.totalAssets,
      col,
      `=SUM(${letter}${bsRows.cash + 1}:${letter}${bsRows.ppe + 1})`
    );
    setValue(workbook, "BS", bsRows.debt, col, i === 0 ? "=40000" : `=${prev}${bsRows.debt + 1}`);
    setValue(
      workbook,
      "BS",
      bsRows.equity,
      col,
      i === 0
        ? `=65000+IS!${letter}10`
        : `=${prev}${bsRows.equity + 1}+IS!${letter}10`
    );
    setValue(
      workbook,
      "BS",
      bsRows.totalLE,
      col,
      `=${letter}${bsRows.debt + 1}+${letter}${bsRows.equity + 1}`
    );
    setValue(
      workbook,
      "BS",
      bsRows.check,
      col,
      `=${letter}${bsRows.totalAssets + 1}-${letter}${bsRows.totalLE + 1}`
    );
  }

  // Cash flow.
  setValue(workbook, "CF", 0, 0, "Cash flow");
  periodHeaders(workbook, "CF", 0, 1, years);
  setValue(workbook, "CF", 1, 0, "Net income");
  setValue(workbook, "CF", 2, 0, "Add back D&A");
  setValue(workbook, "CF", 3, 0, "Change in AR");
  setValue(workbook, "CF", 4, 0, "Capex");
  setValue(workbook, "CF", 5, 0, "Net cash flow");
  for (let i = 0; i < years; i++) {
    const col = 1 + i;
    const letter = a1(0, col).replace(/[0-9]+$/, "");
    const prev = a1(0, col - 1).replace(/[0-9]+$/, "");
    setValue(workbook, "CF", 1, col, `=IS!${letter}10`);
    setValue(workbook, "CF", 2, col, `=-IS!${letter}7`);
    setValue(
      workbook,
      "CF",
      3,
      col,
      i === 0 ? `=-BS!${letter}3` : `=-(BS!${letter}3-BS!${prev}3)`
    );
    setValue(workbook, "CF", 4, col, `=-IS!${letter}2*Assumptions!${letter}6`);
    setValue(workbook, "CF", 5, col, `=SUM(${letter}2:${letter}5)`);
  }

  workbook.names.push({
    name: "TaxRate",
    scope: null,
    refersTo: "=Assumptions!$B$5:$F$5",
  });

  return { id, title: "3-statement model (clean)", workbook, defects: [] };
}

// ---------------------------------------------------------------- DCF model

export function dcfModel(id = "dcf"): CorpusWorkbook {
  const workbook = new Workbook("DCF Valuation");
  const years = 5;

  setValue(workbook, "DCF", 0, 0, "DCF");
  periodHeaders(workbook, "DCF", 0, 1, years);
  setValue(workbook, "DCF", 1, 0, "EBIT");
  setValue(workbook, "DCF", 2, 0, "Less: tax");
  setValue(workbook, "DCF", 3, 0, "NOPAT");
  setValue(workbook, "DCF", 4, 0, "Plus: D&A");
  setValue(workbook, "DCF", 5, 0, "Less: capex");
  setValue(workbook, "DCF", 6, 0, "Less: change in NWC");
  setValue(workbook, "DCF", 7, 0, "Free cash flow");
  setValue(workbook, "DCF", 8, 0, "Discount factor");
  setValue(workbook, "DCF", 9, 0, "PV of FCF");

  setValue(workbook, "Inputs", 0, 0, "WACC");
  setValue(workbook, "Inputs", 0, 1, 0.095);
  setValue(workbook, "Inputs", 1, 0, "Terminal growth");
  setValue(workbook, "Inputs", 1, 1, 0.025);
  setValue(workbook, "Inputs", 2, 0, "Tax rate");
  setValue(workbook, "Inputs", 2, 1, 0.25);
  setValue(workbook, "Inputs", 3, 0, "Net debt");
  setValue(workbook, "Inputs", 3, 1, 42_000);
  setValue(workbook, "Inputs", 4, 0, "Shares outstanding");
  setValue(workbook, "Inputs", 4, 1, 10_000);

  workbook.names.push({ name: "WACC", scope: null, refersTo: "=Inputs!$B$1" });
  workbook.names.push({ name: "TerminalGrowth", scope: null, refersTo: "=Inputs!$B$2" });

  const ebit = [52_000, 58_500, 64_800, 70_200, 75_100];
  for (let i = 0; i < years; i++) {
    const col = 1 + i;
    const letter = a1(0, col).replace(/[0-9]+$/, "");
    setValue(workbook, "DCF", 1, col, ebit[i]!);
    setValue(workbook, "DCF", 2, col, `=-${letter}2*Inputs!$B$3`);
    setValue(workbook, "DCF", 3, col, `=${letter}2+${letter}3`);
    setValue(workbook, "DCF", 4, col, `=${letter}2*0.08`);
    setValue(workbook, "DCF", 5, col, `=-${letter}2*0.09`);
    setValue(workbook, "DCF", 6, col, `=-${letter}2*0.02`);
    setValue(workbook, "DCF", 7, col, `=SUM(${letter}4:${letter}7)`);
    setValue(workbook, "DCF", 8, col, `=1/(1+WACC)^${i + 1}`);
    setValue(workbook, "DCF", 9, col, `=${letter}8*${letter}9`);
  }

  setValue(workbook, "DCF", 11, 0, "Sum of PV");
  setValue(workbook, "DCF", 11, 1, "=SUM(B10:F10)");
  setValue(workbook, "DCF", 12, 0, "Terminal value");
  setValue(workbook, "DCF", 12, 1, "=F8*(1+TerminalGrowth)/(WACC-TerminalGrowth)");
  setValue(workbook, "DCF", 13, 0, "PV of terminal value");
  setValue(workbook, "DCF", 13, 1, "=B13*F9");
  setValue(workbook, "DCF", 14, 0, "Enterprise value");
  setValue(workbook, "DCF", 14, 1, "=B12+B14");
  setValue(workbook, "DCF", 15, 0, "Equity value");
  setValue(workbook, "DCF", 15, 1, "=B15-Inputs!B4");
  setValue(workbook, "DCF", 16, 0, "Value per share");
  setValue(workbook, "DCF", 16, 1, "=B16/Inputs!B5");

  return { id, title: "DCF valuation (clean)", workbook, defects: [] };
}

// ------------------------------------------------------- budget vs actual

export function budgetVsActual(id = "budget-vs-actual"): CorpusWorkbook {
  const workbook = new Workbook("Budget vs Actual");
  const months = 12;
  const departments = ["Sales", "Marketing", "Engineering", "G&A", "Support"];

  setValue(workbook, "Budget", 0, 0, "Department");
  for (let m = 0; m < months; m++) {
    setValue(workbook, "Budget", 0, 1 + m, `M${m + 1}`);
    setValue(workbook, "Actual", 0, 1 + m, `M${m + 1}`);
    setValue(workbook, "Variance", 0, 1 + m, `M${m + 1}`);
  }
  setValue(workbook, "Budget", 0, months + 1, "Total");
  setValue(workbook, "Actual", 0, 0, "Department");
  setValue(workbook, "Actual", 0, months + 1, "Total");
  setValue(workbook, "Variance", 0, 0, "Department");
  setValue(workbook, "Variance", 0, months + 1, "Total");

  departments.forEach((department, index) => {
    const row = index + 1;
    setValue(workbook, "Budget", row, 0, department);
    setValue(workbook, "Actual", row, 0, department);
    setValue(workbook, "Variance", row, 0, department);
    for (let m = 0; m < months; m++) {
      const col = 1 + m;
      const letter = a1(0, col).replace(/[0-9]+$/, "");
      setValue(workbook, "Budget", row, col, 20_000 + index * 5_000 + m * 250);
      setValue(workbook, "Actual", row, col, 19_400 + index * 5_100 + m * 260);
      setValue(workbook, "Variance", row, col, `=Actual!${letter}${row + 1}-Budget!${letter}${row + 1}`);
    }
    const totalCol = months + 1;
    const first = a1(0, 1).replace(/[0-9]+$/, "");
    const last = a1(0, months).replace(/[0-9]+$/, "");
    setValue(workbook, "Budget", row, totalCol, `=SUM(${first}${row + 1}:${last}${row + 1})`);
    setValue(workbook, "Actual", row, totalCol, `=SUM(${first}${row + 1}:${last}${row + 1})`);
    setValue(workbook, "Variance", row, totalCol, `=SUM(${first}${row + 1}:${last}${row + 1})`);
  });

  const totalRow = departments.length + 1;
  for (const sheetName of ["Budget", "Actual", "Variance"]) {
    setValue(workbook, sheetName, totalRow, 0, "Total");
    for (let m = 0; m <= months; m++) {
      const col = 1 + m;
      const letter = a1(0, col).replace(/[0-9]+$/, "");
      setValue(
        workbook,
        sheetName,
        totalRow,
        col,
        `=SUM(${letter}2:${letter}${departments.length + 1})`
      );
    }
  }

  return { id, title: "Budget vs actual (clean)", workbook, defects: [] };
}

// ------------------------------------------------------- cohort analysis

export function cohortAnalysis(id = "cohort"): CorpusWorkbook {
  const workbook = new Workbook("Cohort Analysis");
  const cohorts = 8;
  const periods = 8;

  setValue(workbook, "Cohorts", 0, 0, "Cohort");
  for (let p = 0; p < periods; p++) setValue(workbook, "Cohorts", 0, 1 + p, `Month ${p}`);
  setValue(workbook, "Cohorts", 0, periods + 1, "Retention M3");

  for (let c = 0; c < cohorts; c++) {
    const row = c + 1;
    setValue(workbook, "Cohorts", row, 0, `2025-${String(c + 1).padStart(2, "0")}`);
    setValue(workbook, "Cohorts", row, 1, 1_000 + c * 120);
    for (let p = 1; p < periods; p++) {
      const col = 1 + p;
      const letter = a1(0, col).replace(/[0-9]+$/, "");
      const prev = a1(0, col - 1).replace(/[0-9]+$/, "");
      setValue(workbook, "Cohorts", row, col, `=${prev}${row + 1}*Rates!$B$${p + 1}`);
    }
    setValue(workbook, "Cohorts", row, periods + 1, `=D${row + 1}/B${row + 1}`);
  }

  setValue(workbook, "Rates", 0, 0, "Period");
  setValue(workbook, "Rates", 0, 1, "Retention");
  for (let p = 1; p < periods; p++) {
    setValue(workbook, "Rates", p, 0, `M${p}`);
    setValue(workbook, "Rates", p, 1, Number((0.92 - p * 0.03).toFixed(3)));
  }

  return { id, title: "Cohort analysis (clean)", workbook, defects: [] };
}

export const CLEAN_GENERATORS = [
  threeStatementModel,
  dcfModel,
  budgetVsActual,
  cohortAnalysis,
];

/**
 * The Phase 2 gate fixture: a deliberately broken 15-sheet DCF.
 *
 * Gate wording: "a deliberately broken 15-sheet DCF yields the correct top-5
 * issues". So the model is built at realistic scale (segment build-ups feeding
 * a consolidation, then a valuation) and five defects of clearly different
 * severity are planted. The test asserts the engine's TOP FIVE — ranked by
 * severity then blast radius — are exactly those five, which is a much
 * stronger claim than "it found them somewhere in the list".
 */

import { Workbook, a1 } from "../model/workbook";
import { CorpusWorkbook, InjectedDefect } from "./models";

const SEGMENTS = ["Retail", "Wholesale", "Digital", "Services", "Licensing", "Other"];
const YEARS = 6;

function put(
  workbook: Workbook,
  sheetName: string,
  row: number,
  col: number,
  value: string | number
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

const letter = (col: number): string => a1(0, col).replace(/[0-9]+$/, "");

/** One segment sheet: drivers on top, a revenue/margin build-up below. */
function buildSegment(workbook: Workbook, name: string, index: number): void {
  const sheet = `Seg_${name}`;
  put(workbook, sheet, 0, 0, `${name} segment`);
  for (let y = 0; y < YEARS; y++) put(workbook, sheet, 0, 1 + y, `FY${2024 + y}`);

  put(workbook, sheet, 1, 0, "Units");
  put(workbook, sheet, 2, 0, "Price");
  put(workbook, sheet, 3, 0, "Revenue");
  put(workbook, sheet, 4, 0, "Direct cost %");
  put(workbook, sheet, 5, 0, "Direct cost");
  put(workbook, sheet, 6, 0, "Segment margin");

  for (let y = 0; y < YEARS; y++) {
    const col = 1 + y;
    const L = letter(col);
    const P = letter(col - 1);
    put(workbook, sheet, 1, col, y === 0 ? 1_000 + index * 250 : `=${P}2*(1+Drivers!$B$${index + 2})`);
    put(workbook, sheet, 2, col, y === 0 ? 20 + index * 3 : `=${P}3*(1+Drivers!$B$10)`);
    put(workbook, sheet, 3, col, `=${L}2*${L}3`);
    put(workbook, sheet, 4, col, 0.55 + index * 0.02);
    put(workbook, sheet, 5, col, `=-${L}4*${L}5`);
    put(workbook, sheet, 6, col, `=${L}4+${L}6`);
  }
}

export function brokenBigDcf(id = "broken-big-dcf"): CorpusWorkbook {
  const workbook = new Workbook("Project Atlas DCF");

  // --- Drivers sheet -----------------------------------------------------
  put(workbook, "Drivers", 0, 0, "Driver");
  put(workbook, "Drivers", 0, 1, "Value");
  SEGMENTS.forEach((segment, index) => {
    put(workbook, "Drivers", index + 1, 0, `${segment} unit growth`);
    put(workbook, "Drivers", index + 1, 1, 0.08 - index * 0.005);
  });
  put(workbook, "Drivers", 9, 0, "Price inflation");
  put(workbook, "Drivers", 9, 1, 0.025);
  put(workbook, "Drivers", 10, 0, "Opex % of revenue");
  put(workbook, "Drivers", 10, 1, 0.18);
  put(workbook, "Drivers", 11, 0, "Tax rate");
  put(workbook, "Drivers", 11, 1, 0.24);
  put(workbook, "Drivers", 12, 0, "WACC");
  put(workbook, "Drivers", 12, 1, 0.092);
  put(workbook, "Drivers", 13, 0, "Terminal growth");
  put(workbook, "Drivers", 13, 1, 0.023);
  put(workbook, "Drivers", 14, 0, "Net debt");
  put(workbook, "Drivers", 14, 1, 128_000);
  put(workbook, "Drivers", 15, 0, "Shares");
  put(workbook, "Drivers", 15, 1, 24_500);
  put(workbook, "Drivers", 16, 0, "D&A % of revenue");
  put(workbook, "Drivers", 16, 1, 0.035);
  put(workbook, "Drivers", 17, 0, "Capex % of revenue");
  put(workbook, "Drivers", 17, 1, 0.042);
  put(workbook, "Drivers", 18, 0, "NWC % of revenue");
  put(workbook, "Drivers", 18, 1, 0.015);

  workbook.names.push({ name: "WACC", scope: null, refersTo: "=Drivers!$B$13" });
  workbook.names.push({ name: "TerminalGrowth", scope: null, refersTo: "=Drivers!$B$14" });
  workbook.names.push({ name: "TaxRate", scope: null, refersTo: "=Drivers!$B$12" });

  // --- 6 segment sheets --------------------------------------------------
  SEGMENTS.forEach((segment, index) => buildSegment(workbook, segment, index));

  // --- Consolidation -----------------------------------------------------
  put(workbook, "Consol", 0, 0, "Consolidated");
  for (let y = 0; y < YEARS; y++) put(workbook, "Consol", 0, 1 + y, `FY${2024 + y}`);
  put(workbook, "Consol", 1, 0, "Revenue");
  put(workbook, "Consol", 2, 0, "Direct cost");
  put(workbook, "Consol", 3, 0, "Gross margin");
  put(workbook, "Consol", 4, 0, "Opex");
  put(workbook, "Consol", 5, 0, "EBITDA");
  put(workbook, "Consol", 6, 0, "D&A");
  put(workbook, "Consol", 7, 0, "EBIT");

  for (let y = 0; y < YEARS; y++) {
    const col = 1 + y;
    const L = letter(col);
    const revenueTerms = SEGMENTS.map((segment) => `Seg_${segment}!${L}4`).join("+");
    const costTerms = SEGMENTS.map((segment) => `Seg_${segment}!${L}6`).join("+");
    put(workbook, "Consol", 1, col, `=${revenueTerms}`);
    put(workbook, "Consol", 2, col, `=${costTerms}`);
    put(workbook, "Consol", 3, col, `=${L}2+${L}3`);
    put(workbook, "Consol", 4, col, `=-${L}2*Drivers!$B$11`);
    put(workbook, "Consol", 5, col, `=${L}4+${L}5`);
    put(workbook, "Consol", 6, col, `=-${L}2*Drivers!$B$17`);
    put(workbook, "Consol", 7, col, `=${L}6+${L}7`);
  }

  // --- Free cash flow ----------------------------------------------------
  put(workbook, "FCF", 0, 0, "Free cash flow");
  for (let y = 0; y < YEARS; y++) put(workbook, "FCF", 0, 1 + y, `FY${2024 + y}`);
  put(workbook, "FCF", 1, 0, "EBIT");
  put(workbook, "FCF", 2, 0, "Tax");
  put(workbook, "FCF", 3, 0, "NOPAT");
  put(workbook, "FCF", 4, 0, "Add: D&A");
  put(workbook, "FCF", 5, 0, "Less: capex");
  put(workbook, "FCF", 6, 0, "Less: change in NWC");
  put(workbook, "FCF", 7, 0, "Unlevered FCF");

  for (let y = 0; y < YEARS; y++) {
    const col = 1 + y;
    const L = letter(col);
    put(workbook, "FCF", 1, col, `=Consol!${L}8`);
    put(workbook, "FCF", 2, col, `=-MAX(0,${L}2)*TaxRate`);
    put(workbook, "FCF", 3, col, `=${L}2+${L}3`);
    put(workbook, "FCF", 4, col, `=-Consol!${L}7`);
    put(workbook, "FCF", 5, col, `=-Consol!${L}2*Drivers!$B$18`);
    put(workbook, "FCF", 6, col, `=-Consol!${L}2*Drivers!$B$19`);
    put(workbook, "FCF", 7, col, `=SUM(${L}4:${L}7)`);
  }

  // --- Valuation ---------------------------------------------------------
  put(workbook, "Valuation", 0, 0, "Valuation");
  for (let y = 0; y < YEARS; y++) put(workbook, "Valuation", 0, 1 + y, `FY${2024 + y}`);
  put(workbook, "Valuation", 1, 0, "Unlevered FCF");
  put(workbook, "Valuation", 2, 0, "Discount factor");
  put(workbook, "Valuation", 3, 0, "PV of FCF");
  for (let y = 0; y < YEARS; y++) {
    const col = 1 + y;
    const L = letter(col);
    put(workbook, "Valuation", 1, col, `=FCF!${L}8`);
    put(workbook, "Valuation", 2, col, `=1/(1+WACC)^${y + 1}`);
    put(workbook, "Valuation", 3, col, `=${L}2*${L}3`);
  }
  put(workbook, "Valuation", 5, 0, "Sum of PV");
  put(workbook, "Valuation", 5, 1, "=SUM(B4:G4)");
  put(workbook, "Valuation", 6, 0, "Terminal value");
  put(workbook, "Valuation", 6, 1, "=G2*(1+TerminalGrowth)/(WACC-TerminalGrowth)");
  put(workbook, "Valuation", 7, 0, "PV of terminal value");
  put(workbook, "Valuation", 7, 1, "=B7*G3");
  put(workbook, "Valuation", 8, 0, "Enterprise value");
  put(workbook, "Valuation", 8, 1, "=B6+B8");
  put(workbook, "Valuation", 9, 0, "Equity value");
  put(workbook, "Valuation", 9, 1, "=B9-Drivers!B15");
  put(workbook, "Valuation", 10, 0, "Value per share");
  put(workbook, "Valuation", 10, 1, "=B10/Drivers!B16");

  // --- Supporting sheets (bring the count to 15) -------------------------
  put(workbook, "Sensitivity", 0, 0, "WACC / growth sensitivity");
  put(workbook, "Sensitivity", 1, 0, "Base value per share");
  put(workbook, "Sensitivity", 1, 1, "=Valuation!B11");
  put(workbook, "Comparables", 0, 0, "Peer");
  put(workbook, "Comparables", 0, 1, "EV/EBITDA");
  ["Peer A", "Peer B", "Peer C", "Peer D"].forEach((peer, index) => {
    put(workbook, "Comparables", index + 1, 0, peer);
    put(workbook, "Comparables", index + 1, 1, 9.5 + index * 0.8);
  });
  put(workbook, "Comparables", 6, 0, "Median");
  put(workbook, "Comparables", 6, 1, "=MEDIAN(B2:B5)");
  put(workbook, "Checks", 0, 0, "Model checks");
  put(workbook, "Checks", 1, 0, "Segment revenue check (must be 0)");
  for (let y = 0; y < YEARS; y++) {
    const L = letter(1 + y);
    const terms = SEGMENTS.map((segment) => `Seg_${segment}!${L}4`).join("+");
    put(workbook, "Checks", 1, 1 + y, `=Consol!${L}2-(${terms})`);
  }
  // Ratios come from the drivers sheet: everything outside the five planted
  // defects must be clean modelling, or the fixture measures its own sloppiness.
  put(workbook, "Drivers", 19, 0, "Receivables % of revenue");
  put(workbook, "Drivers", 19, 1, 0.11);
  put(workbook, "Drivers", 20, 0, "Payables % of direct cost");
  put(workbook, "Drivers", 20, 1, 0.09);
  put(workbook, "WorkingCapital", 0, 0, "Working capital");
  for (let y = 0; y < YEARS; y++) put(workbook, "WorkingCapital", 0, 1 + y, `FY${2024 + y}`);
  put(workbook, "WorkingCapital", 1, 0, "Receivables");
  put(workbook, "WorkingCapital", 2, 0, "Payables");
  put(workbook, "WorkingCapital", 3, 0, "Net working capital");
  for (let y = 0; y < YEARS; y++) {
    const col = 1 + y;
    const L = letter(col);
    put(workbook, "WorkingCapital", 1, col, `=Consol!${L}2*Drivers!$B$20`);
    put(workbook, "WorkingCapital", 2, col, `=-Consol!${L}3*Drivers!$B$21`);
    put(workbook, "WorkingCapital", 3, col, `=${L}2+${L}3`);
  }
  put(workbook, "Notes", 0, 0, "Prepared by the deal team. Do not distribute.");

  // ======================================================================
  // Five planted defects, deliberately of different severities so the
  // ranking (severity, then blast radius) is actually exercised.
  // ======================================================================
  const defects: InjectedDefect[] = [];

  // 1. CRITICAL — circular reference between EV and equity value.
  put(workbook, "Valuation", 8, 1, "=B6+B8+B10*0");
  put(workbook, "Valuation", 9, 1, "=B9-Drivers!B15");
  put(workbook, "Valuation", 10, 1, "=B10/Drivers!B16+0*B9");
  defects.push({
    rule: "AUD-005",
    address: "Valuation!B9",
    alsoAcceptable: ["Valuation!B10", "Valuation!B11"],
    note: "circular reference across enterprise value, equity value and per-share value",
  });

  // 2. CRITICAL — a #REF! deep in the consolidation, poisoning everything after it.
  put(workbook, "Consol", 1, 3, "=#REF!+Seg_Wholesale!D4");
  workbook.sheet("Consol")!.get(1, 3)!.value = "#REF!";
  defects.push({
    rule: "AUD-004",
    address: "Consol!D2",
    note: "deleted segment link leaves #REF! in consolidated revenue",
  });

  // 3. HIGH — a plugged constant over the FY2027 EBIT formula.
  put(workbook, "Consol", 7, 4, 184_500);
  defects.push({
    rule: "AUD-003",
    address: "Consol!E8",
    note: "EBIT overwritten with a hardcoded number",
  });

  // 4. HIGH — the FY2028 capex line breaks the row's fill pattern.
  put(workbook, "FCF", 5, 5, "=-Consol!F2*Drivers!$B$17");
  defects.push({
    rule: "AUD-001",
    address: "FCF!F6",
    note: "capex row points at the D&A driver instead of the capex driver",
  });

  // 5. MEDIUM — a magic tax rate baked into one year's tax line.
  put(workbook, "FCF", 2, 6, "=-MAX(0,G2)*0.31");
  defects.push({
    rule: "AUD-002",
    address: "FCF!G3",
    note: "hardcoded 31% tax rate instead of the TaxRate name",
  });

  return {
    id,
    title: "15-sheet DCF with five planted defects",
    workbook,
    defects,
  };
}

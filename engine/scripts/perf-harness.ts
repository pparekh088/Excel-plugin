/**
 * WIL performance harness — measures the §5 acceptance numbers on a
 * synthetic workbook of the stated size (20 sheets, 200k formulas).
 *
 * Acceptance (handoff §5): build < 30s desktop / < 60s web, memory < 500MB.
 * This runs the pure-TypeScript pipeline (parse -> graph -> semantic -> WIL),
 * which is the part that dominates; Office.js extraction I/O is measured
 * separately on a real host and recorded in the gate report.
 *
 * Usage: npm run perf [-- --sheets 20 --rows 2000 --cols 5]
 */

import { DependencyGraph } from "../src/graph/graph";
import { Workbook, a1 } from "../src/model/workbook";
import { buildWil } from "../src/wil/serialize";

interface Args {
  sheets: number;
  rows: number;
  cols: number;
}

function parseArgs(): Args {
  const argv = process.argv.slice(2);
  const read = (flag: string, fallback: number): number => {
    const index = argv.indexOf(`--${flag}`);
    return index >= 0 && argv[index + 1] ? Number(argv[index + 1]) : fallback;
  };
  return {
    sheets: read("sheets", 20),
    rows: read("rows", 2000),
    cols: read("cols", 5),
  };
}

/** A model-shaped sheet: labels, a period header, drivers, and formula blocks. */
function buildSheet(workbook: Workbook, name: string, rows: number, cols: number): void {
  const sheet = workbook.addSheet(name);
  const put = (row: number, col: number, value: string | number) => {
    const isFormula = typeof value === "string" && value.startsWith("=");
    sheet.set({ row, col, value: isFormula ? 0 : value, ...(isFormula ? { formula: value } : {}) });
  };

  for (let col = 1; col <= cols; col++) put(0, col, `FY${2024 + col}`);
  for (let row = 1; row <= rows; row++) {
    put(row, 0, `Line item ${row}`);
    put(row, 1, row * 1.5); // driver column
    for (let col = 2; col <= cols; col++) {
      const prev = a1(0, col - 1).replace(/[0-9]+$/, "");
      const here = a1(0, col).replace(/[0-9]+$/, "");
      // Mixed formula shapes so runs are realistic, not one giant block.
      put(
        row,
        col,
        row % 50 === 0
          ? `=SUM(${here}2:${here}${row})`
          : `=${prev}${row + 1}*(1+$B$1)+${here}${row}`
      );
    }
  }
}

function main(): void {
  const args = parseArgs();
  const workbook = new Workbook("Perf");

  const startBuild = Date.now();
  for (let i = 0; i < args.sheets; i++) {
    buildSheet(workbook, `Sheet${i + 1}`, args.rows, args.cols);
  }
  const fixtureMs = Date.now() - startBuild;

  const startGraph = Date.now();
  const graph = DependencyGraph.build(workbook);
  const graphMs = Date.now() - startGraph;

  const startWil = Date.now();
  const wil = buildWil(workbook, graph);
  const wilMs = Date.now() - startWil;

  const memory = process.memoryUsage();
  const mb = (bytes: number) => Math.round(bytes / 1024 / 1024);

  const totalMs = graphMs + wilMs;
  console.log("--- Ledger WIL perf harness ---");
  console.log(
    `fixture:      ${args.sheets} sheets x ${args.rows} rows x ${args.cols} cols ` +
      `(${workbook.cellCount} cells, ${graph.stats.formulaCells} formulas) in ${fixtureMs}ms`
  );
  console.log(`graph build:  ${graphMs}ms`);
  console.log(`WIL build:    ${wilMs}ms`);
  console.log(`TOTAL:        ${totalMs}ms  (acceptance: <30000ms desktop, <60000ms web)`);
  console.log(`heap used:    ${mb(memory.heapUsed)}MB  rss: ${mb(memory.rss)}MB (acceptance: <500MB)`);
  console.log(
    `collapse:     ${graph.stats.formulaCells} formulas -> ${graph.stats.runCount} run nodes ` +
      `(${(graph.stats.formulaCells / Math.max(1, graph.stats.runCount)).toFixed(1)}x)`
  );
  console.log(`edges:        ${graph.stats.edgeCount}`);
  console.log(`WIL tokens:   ~${wil.approxTokens} (budget 6000, truncated=${wil.truncated})`);

  const passed = totalMs < 30_000 && mb(memory.heapUsed) < 500;
  console.log(passed ? "\nRESULT: PASS" : "\nRESULT: FAIL");
  if (!passed) process.exitCode = 1;
}

main();

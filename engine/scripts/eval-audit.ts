/**
 * Audit-detection eval (handoff §9, Phase 2 gate).
 *
 * Scores the audit engine against corpus ground truth. Precision counts a
 * finding against us only when it is unrelated to any injected defect — see
 * src/eval/auditScoring.ts for why naive address matching is the wrong
 * measure for an audit engine. On CLEAN workbooks nothing is injected, so
 * every finding is a false positive, which is where the gate really bites.
 *
 * Gate: >= 95% precision across the corpus.
 *
 * Usage: npm run eval:audit -w engine [-- --verbose]
 */

import { runAudit } from "../src/audit/engine";
import { DependencyGraph } from "../src/graph/graph";
import { Simulator } from "../src/sim/simulator";
import { buildCorpus } from "../src/corpus";
import { scoreAudit } from "../src/eval/auditScoring";

const verbose = process.argv.includes("--verbose");

function main(): void {
  const corpus = buildCorpus();
  const rows: Array<{
    id: string;
    clean: boolean;
    expected: number;
    matched: number;
    findings: number;
    falsePositives: number;
    health: number;
    ms: number;
  }> = [];

  let totalFindings = 0;
  let totalFalsePositives = 0;
  let totalExpected = 0;
  let totalMatched = 0;
  const falsePositiveDetail: string[] = [];
  const missedDetail: string[] = [];

  for (const entry of corpus) {
    // Value-based rules (AUD-004, AUD-008) need current values.
    Simulator.of(entry.workbook).recalculate();

    const started = Date.now();
    const graph = DependencyGraph.build(entry.workbook);
    const report = runAudit(entry.workbook, { graph });
    const ms = Date.now() - started;

    if (report.stats.llmCallsUsed !== 0) {
      console.error(`FATAL: ${entry.id} used ${report.stats.llmCallsUsed} LLM calls`);
      process.exitCode = 1;
    }

    const score = scoreAudit(report.findings, entry.defects, entry.workbook, graph);

    totalFindings += score.findingsTotal;
    totalFalsePositives += score.falsePositives;
    totalExpected += score.expectedTotal;
    totalMatched += score.matched;
    falsePositiveDetail.push(...score.falsePositiveDetail.map((item) => `${entry.id}: ${item}`));
    missedDetail.push(...score.missedDetail.map((item) => `${entry.id}: ${item}`));

    rows.push({
      id: entry.id,
      clean: entry.defects.length === 0,
      expected: score.expectedTotal,
      matched: score.matched,
      findings: score.findingsTotal,
      falsePositives: score.falsePositives,
      health: report.health.score,
      ms,
    });

    if (verbose) {
      console.log(`\n=== ${entry.id} (${entry.title}) ===`);
      console.log(`health ${report.health.score} (${report.health.band})  ${ms}ms`);
      console.log(report.coverage.caveat);
      for (const finding of report.findings.slice(0, 12)) {
        console.log(
          `  [${finding.severity}] ${finding.rule} ${finding.address} blast=${finding.blastRadius}`
        );
        console.log(`      ${finding.explanation}`);
      }
      if (report.findings.length > 12) console.log(`  ... ${report.findings.length - 12} more`);
    }
  }

  console.log("\n--- Audit detection eval ---");
  console.log(
    "workbook".padEnd(30) +
      "exp".padStart(5) +
      "hit".padStart(5) +
      "found".padStart(7) +
      "FP".padStart(5) +
      "health".padStart(8) +
      "ms".padStart(6)
  );
  for (const row of rows) {
    console.log(
      row.id.padEnd(30) +
        String(row.expected).padStart(5) +
        String(row.matched).padStart(5) +
        String(row.findings).padStart(7) +
        String(row.falsePositives).padStart(5) +
        String(row.health).padStart(8) +
        String(row.ms).padStart(6)
    );
  }

  const precision =
    totalFindings === 0 ? 1 : (totalFindings - totalFalsePositives) / totalFindings;
  const recall = totalExpected === 0 ? 1 : totalMatched / totalExpected;

  console.log(
    `\nprecision ${(precision * 100).toFixed(1)}%  (${totalFindings - totalFalsePositives}/${totalFindings} findings related to a real defect)`
  );
  console.log(
    `recall    ${(recall * 100).toFixed(1)}%  (${totalMatched}/${totalExpected} injected defects found)`
  );

  const cleanRows = rows.filter((row) => row.clean);
  const cleanFindings = cleanRows.reduce((sum, row) => sum + row.findings, 0);
  console.log(
    `clean baselines: ${cleanFindings} finding(s) across ${cleanRows.length} clean workbooks ` +
      `(every one is a false positive)`
  );

  if (falsePositiveDetail.length > 0) {
    console.log(`\nFalse positives (${falsePositiveDetail.length}):`);
    for (const item of falsePositiveDetail.slice(0, 40)) console.log(`  ${item}`);
    if (falsePositiveDetail.length > 40) {
      console.log(`  ... ${falsePositiveDetail.length - 40} more`);
    }
  }
  if (missedDetail.length > 0) {
    console.log(`\nMissed defects (${missedDetail.length}):`);
    for (const item of missedDetail) console.log(`  ${item}`);
  }

  const gate = precision >= 0.95;
  console.log(`\nGATE (precision >= 95%): ${gate ? "PASS" : "FAIL"}`);
  if (!gate) process.exitCode = 1;
}

main();

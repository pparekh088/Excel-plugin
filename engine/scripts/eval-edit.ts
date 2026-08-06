/**
 * Edit/build task eval (handoff §9, Phase 3 gate).
 *
 * Gate: task success >= 85%, cells-destroyed = 0 across the suite.
 *
 * Runs the full agent loop — plan, approve, change set, drift check, apply,
 * recalculate, verify, repair — against the headless simulator with a scripted
 * provider, so CI needs no network and no Excel. Swap the provider for a live
 * one to score real models on the same tasks.
 *
 * Usage: npm run eval:edit -w engine [-- --verbose]
 */

import { MockLlmProvider, CostMeter } from "../src/agent/llm";
import { INTENT_MARKER, createDefaultExecutor, runAgent } from "../src/agent/runtime";
import { EDIT_TASKS } from "../src/eval/editTasks";
import {
  accountDestruction,
  gradeCells,
  snapshotWorkbook,
  summarize,
  type TaskRun,
} from "../src/eval/harness";
import { Simulator } from "../src/sim/simulator";

const verbose = process.argv.includes("--verbose");

async function main(): Promise<void> {
  const runs: TaskRun[] = [];

  for (const task of EDIT_TASKS) {
    const corpus = task.workbook();
    const workbook = corpus.workbook;
    Simulator.of(workbook).recalculate();

    const before = snapshotWorkbook(workbook);
    const provider = new MockLlmProvider().script(INTENT_MARKER, task.scriptedPlan);
    if (task.scriptedRepair) provider.script("verification failed", task.scriptedRepair);

    const meter = new CostMeter();
    const started = Date.now();
    const result = await runAgent(workbook, {
      intent: task.intent,
      provider,
      approver: async () => "approve",
      executor: createDefaultExecutor(),
      costMeter: meter,
    });
    const latencyMs = Date.now() - started;

    const expectedOutcome = task.expectOutcome ?? "applied";
    const outcomeOk = result.outcome === expectedOutcome;

    const destruction = accountDestruction(
      before,
      workbook,
      new Set(task.allowedChanges)
    );
    const grade =
      task.expectations.length > 0 ? gradeCells(workbook, task.expectations) : { passed: true, checks: [] };

    // Cells destroyed is a hard failure regardless of whether the intent was met.
    const success = outcomeOk && grade.passed && destruction.destroyed.length === 0;

    runs.push({
      taskId: task.id,
      taskClass: task.taskClass,
      success,
      cellsDestroyed: destruction.destroyed.length,
      destroyedAddresses: destruction.destroyed,
      toolCalls: result.plan?.steps.length ?? 0,
      latencyMs,
      costUsd: meter.totalCostUsd,
      detail: outcomeOk
        ? grade.passed
          ? destruction.destroyed.length === 0
            ? "ok"
            : `destroyed ${destruction.destroyed.join(", ")}`
          : grade.checks.filter((check) => !check.ok).map((check) => check.detail).join("; ")
        : `expected outcome ${expectedOutcome}, got ${result.outcome}`,
    });

    if (verbose) {
      console.log(`\n=== ${task.id} ===`);
      console.log(`intent: ${task.intent}`);
      console.log(`outcome: ${result.outcome} (expected ${expectedOutcome})`);
      for (const line of result.transcript) console.log(`  | ${line.split("\n")[0]}`);
    }
  }

  const summary = summarize(runs);

  console.log("\n--- Edit/build task eval ---");
  console.log(
    "task".padEnd(30) + "class".padEnd(8) + "ok".padStart(4) + "destroyed".padStart(11) +
      "steps".padStart(7) + "ms".padStart(6) + "  detail"
  );
  for (const run of runs) {
    console.log(
      run.taskId.padEnd(30) +
        run.taskClass.padEnd(8) +
        (run.success ? "  Y" : "  N").padStart(4) +
        String(run.cellsDestroyed).padStart(11) +
        String(run.toolCalls).padStart(7) +
        String(run.latencyMs).padStart(6) +
        `  ${run.detail}`
    );
  }

  console.log(
    `\nsuccess rate      ${(summary.successRate * 100).toFixed(1)}%  ` +
      `(${runs.filter((run) => run.success).length}/${runs.length})`
  );
  console.log(`cells destroyed   ${summary.totalCellsDestroyed}`);
  console.log(`mean tool calls   ${summary.meanToolCalls.toFixed(1)}`);
  console.log(`mean latency      ${summary.meanLatencyMs.toFixed(0)}ms`);

  const successGate = summary.successRate >= 0.85;
  const destructionGate = summary.totalCellsDestroyed === 0;
  console.log(`\nGATE (success >= 85%):      ${successGate ? "PASS" : "FAIL"}`);
  console.log(`GATE (cells destroyed = 0): ${destructionGate ? "PASS" : "FAIL"}`);
  if (!successGate || !destructionGate) process.exitCode = 1;
}

void main();

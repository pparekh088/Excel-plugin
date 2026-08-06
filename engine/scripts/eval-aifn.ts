/**
 * AI custom-function eval (handoff §10, Phase 4 gate).
 *
 * Gate: "5k-cell AI.CLASSIFY drag completes under budget with cache hit rate
 * >= 60% on re-runs."
 *
 * Simulates the drag against the real coordinator with a stub backend, so it
 * measures the batching, caching and budget logic rather than a provider.
 *
 * Usage: npm run eval:aifn -w engine
 */

import { AiCoordinator, BUDGET_ERROR, type AiRequest } from "../src/aifn/batch";

const CELLS = 5000;
const BUDGET = 200;
const CATEGORIES = "revenue,cogs,opex,capex,other";

/**
 * Two scenarios, because "completes under budget" and "the breaker fires" are
 * both behaviours worth proving:
 *
 *  A. 150 distinct values — a realistic classification column, where repeated
 *     line-item names mean the drag finishes well inside the budget.
 *  B. 400 distinct values — more unique work than the budget allows, where the
 *     breaker must stop spending and mark the remaining cells.
 */
const SCENARIOS = [
  { name: "A: 150 distinct values (fits the budget)", distinct: 150 },
  { name: "B: 400 distinct values (exceeds the budget)", distinct: 400 },
];

function makeRequests(distinct: number): AiRequest[] {
  return Array.from({ length: CELLS }, (_, index) => ({
    fn: "AI.CLASSIFY",
    inputs: [`line item ${index % distinct}`],
    prompt: CATEGORIES,
    model: "cheap",
  }));
}

async function drag(
  coordinator: AiCoordinator,
  requests: AiRequest[]
): Promise<{ answered: number; budgetErrors: number; ms: number }> {
  const started = Date.now();
  const results = await Promise.all(requests.map((request) => coordinator.request(request)));
  return {
    answered: results.filter((result) => result.value !== BUDGET_ERROR).length,
    budgetErrors: results.filter((result) => result.value === BUDGET_ERROR).length,
    ms: Date.now() - started,
  };
}

interface ScenarioResult {
  name: string;
  distinct: number;
  firstAnswered: number;
  firstBudgetErrors: number;
  firstCalls: number;
  batches: number;
  rerunCalls: number;
  rerunHitRate: number;
  breakerFired: boolean;
}

async function runScenario(name: string, distinct: number): Promise<ScenarioResult> {
  let backendCalls = 0;
  let batchesSeen = 0;
  let breakerFired = false;

  const coordinator = new AiCoordinator({
    handler: async (requests) => {
      batchesSeen++;
      backendCalls += requests.length;
      return requests.map((request) => `cat:${String(request.inputs[0])}`);
    },
    budget: BUDGET,
    batchSize: 25,
    // Defer just past the synchronous burst, standing in for the 250ms debounce.
    scheduler: (callback) => queueMicrotask(callback),
    onBudgetExceeded: () => {
      breakerFired = true;
    },
  });

  const requests = makeRequests(distinct);

  const first = await drag(coordinator, requests);
  const firstStats = coordinator.stats;
  const firstCalls = backendCalls;

  // A recalculation resets the per-cycle budget but NOT the cache — that is
  // the point: the second pass should cost close to nothing.
  coordinator.startRecalcCycle();
  const callsBeforeRerun = backendCalls;
  await drag(coordinator, requests);
  const secondStats = coordinator.stats;

  console.log(`\n${name}`);
  console.log(`  first pass:  answered ${first.answered}/${CELLS}, ` +
    `${first.budgetErrors} #AI_BUDGET!, ${firstCalls} backend calls in ${batchesSeen} batches`);
  console.log(`  re-run:      ${backendCalls - callsBeforeRerun} new backend calls, ` +
    `cache hit rate ${(((secondStats.cacheHits - firstStats.cacheHits) / CELLS) * 100).toFixed(1)}%`);
  if (breakerFired) {
    console.log(`  breaker:     fired at ${BUDGET} calls — remaining cells marked #AI_BUDGET!`);
  }

  return {
    name,
    distinct,
    firstAnswered: first.answered,
    firstBudgetErrors: first.budgetErrors,
    firstCalls,
    batches: batchesSeen,
    rerunCalls: backendCalls - callsBeforeRerun,
    rerunHitRate: (secondStats.cacheHits - firstStats.cacheHits) / CELLS,
    breakerFired,
  };
}

async function main(): Promise<void> {
  console.log("--- AI custom-function eval ---");
  console.log(`${CELLS}-cell AI.CLASSIFY drag, budget ${BUDGET} calls per recalc cycle`);

  const results: ScenarioResult[] = [];
  for (const scenario of SCENARIOS) {
    results.push(await runScenario(scenario.name, scenario.distinct));
  }

  const [fits, exceeds] = results as [ScenarioResult, ScenarioResult];

  // The gate: a realistic drag completes inside the budget, and re-running it
  // is nearly free.
  const completesUnderBudget =
    fits.firstBudgetErrors === 0 && fits.firstCalls <= BUDGET && !fits.breakerFired;
  const hitRateGate = fits.rerunHitRate >= 0.6;
  // And the breaker genuinely stops spending when there is more work than budget.
  const breakerWorks =
    exceeds.breakerFired &&
    exceeds.firstCalls === BUDGET &&
    exceeds.firstBudgetErrors > 0;

  console.log(`\nGATE (5k drag completes under budget):  ${completesUnderBudget ? "PASS" : "FAIL"}`);
  console.log(`GATE (re-run cache hit rate >= 60%):    ${hitRateGate ? "PASS" : "FAIL"}`);
  console.log(`CHECK (breaker stops overspending):     ${breakerWorks ? "PASS" : "FAIL"}`);

  if (!completesUnderBudget || !hitRateGate || !breakerWorks) process.exitCode = 1;
}

void main();

# Phase 4 gate report — AI functions + polish

Date: 2026-08-06 · Branch: `claude/ledger-excel-agent-8dqvfb`

Gate per handoff §10: AI.* functions with batching/caching/budget; dependency
visualizer; trusted-session mode; cost dashboard.
**Gate condition: 5k-cell `AI.CLASSIFY` drag completes under budget with cache
hit rate ≥ 60% on re-runs.**

## Results

| Gate condition | Target | Actual | Status |
| --- | --- | --- | --- |
| 5k-cell drag completes under budget | required | **5000/5000 answered, 150 backend calls, 0 `#AI_BUDGET!`** | **pass** |
| Re-run cache hit rate | ≥ 60% | **100%** (0 new backend calls on re-run) | **pass** |
| Budget circuit breaker stops overspending | required | breaker fires at exactly 200 calls, remaining cells marked `#AI_BUDGET!`, banner raised once | **pass** |
| `AI.*` functions | ASK, EXTRACT, CLASSIFY, MATCH, FORECAST | all five, registered in the manifest under the `AI` namespace | **partial — needs Excel to verify** |
| Dependency visualizer | required | layered layout over run nodes + SVG renderer | **pass** |
| Trusted-session mode | required | shipped in Phase 3; auto-approves only all-LOW-risk plans | **pass** |
| Cost dashboard | required | per-tier tokens and cost, cache savings, budget banner | **partial — needs a running backend to verify live** |

```
--- AI custom-function eval ---
5000-cell AI.CLASSIFY drag, budget 200 calls per recalc cycle

A: 150 distinct values (fits the budget)
  first pass:  answered 5000/5000, 0 #AI_BUDGET!, 150 backend calls in 6 batches
  re-run:      0 new backend calls, cache hit rate 100.0%

B: 400 distinct values (exceeds the budget)
  first pass:  answered 2600/5000, 2400 #AI_BUDGET!, 200 backend calls in 16 batches
  re-run:      200 new backend calls, cache hit rate 96.0%
  breaker:     fired at 200 calls — remaining cells marked #AI_BUDGET!

GATE (5k drag completes under budget):  PASS
GATE (re-run cache hit rate >= 60%):    PASS
CHECK (breaker stops overspending):     PASS
```

Two scenarios are measured deliberately. Scenario A is the gate: a realistic
classification column, where repeated line-item names mean 5,000 cells cost
150 calls. Scenario B proves the breaker actually stops spending when there is
more distinct work than budget — the failure mode that produces an unexpected
four-figure bill. A gate that only tested A would leave the expensive case
unproven.

Engine 1054 tests · add-in 36 · server 34, all green. Typecheck clean, ruff
clean, production build emits `functions.js` with the stable filename the
manifest requires.

## INV-7 coverage

- **Batching**: a 250 ms debounce coalesces a drag; 100 distinct requests
  become 4 handler calls at the default batch size, not 100.
- **Content-hash cache** keyed on (inputs, prompt, model), so a re-run is free
  and two cells asking the identical question share one call.
- **Request coalescing**: identical in-flight requests share a single promise
  rather than duplicating the call.
- **Budget breaker**: 200 calls per recalculation cycle by default, then
  `#AI_BUDGET!` and one banner. Cached answers are still served after the
  budget is exhausted, since they cost nothing.
- **Custom functions never mutate other cells**: they read arguments and
  return a value. There is no code path from the functions runtime to a write.

The banner mattered more than it first appears: an earlier version only raised
it when a whole batch was over budget, so a *partially* over-budget batch left
`#AI_BUDGET!` in the grid with nothing explaining it. Caught by the test that
asserts the banner fires exactly once.

## AI.FORECAST is statistics, not generation

`AI.FORECAST` computes Holt-Winters / Holt / simple exponential smoothing and
picks parameters by in-sample MAE. **No language model is involved.** A number
a model invented is indistinguishable from one it computed once it lands in a
cell, and in a financial model that difference is the whole point. The
implementation exists in both the engine (offline path) and the server, and a
test asserts the endpoint is deterministic across repeated calls.

## AUD-013: AI-derived value inventory

Handoff §8 requires AI.* cells to be tagged so the audit engine can inventory
them — "auditors will demand this". Implemented as a deterministic rule:

- Generative cells (`AI.ASK/EXTRACT/CLASSIFY/MATCH`) are listed as **info**
  when nothing reads them, and escalate to **medium** when they feed
  downstream calculations, because at that point an inferred value is
  propagating into computed results.
- `AI.FORECAST` is inventoried separately as *computed*, not generative.
- Cells stranded at `#AI_BUDGET!` are **high**: they hold no answer at all.

It uses zero LLM calls, so the zero-LLM demo still holds with AI functions
present in the workbook.

## Not verifiable in this environment

- **Custom functions in a real Excel host.** The manifest declares the `AI`
  namespace with script/page/metadata URLs and the build emits `functions.js`
  at a stable path, but registration, the streaming path, and Excel's
  recalculation behaviour need a host. On the sideload checklist.
- **The cost dashboard against a live backend.** The panel polls
  `/agent/cost` every 5 s and listens for the budget event; the endpoint is
  tested server-side, the panel is not (no DOM test harness in this package).

## Deviations

- **No streaming into cells.** §8 mentions the custom-functions streaming
  pattern. Batched request/response fits every function we ship — none of them
  produce progressive output — and streaming would complicate the budget
  accounting for no user-visible gain. Revisit if a long-running AI function
  is added. Logged as D-022.
- **`AI.*` functions route to the cheapest tier unconditionally.** They run
  thousands of times over small inputs. A per-function tier override is a
  config change away but has no evidence behind it yet.

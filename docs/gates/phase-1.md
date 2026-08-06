# Phase 1 gate report — WIL, parser, eval scaffold

Date: 2026-08-06 · Branch: `claude/ledger-excel-agent-8dqvfb`

Gate per handoff §10: formula parser with full grammar + 500 unit tests
including pathological cases; chunked extraction; DAG with range-run
collapsing; WIL serializer; eval corpus v1 + headless runner.
**Gate condition: §5 acceptance numbers met; parser test suite green.**

## Results

| Gate item | Target | Actual | Status |
| --- | --- | --- | --- |
| Parser unit tests | ≥ 500 | **713** parser tests (901 engine-wide) | **pass** |
| Pathological cases covered | required | 100 malformed inputs + 2000 deterministic fuzz cases, all non-throwing | **pass** |
| WIL build time, 20 sheets / 200k formulas | < 30s desktop | **8.8s** (pure pipeline: parse → graph → semantic → WIL) | **pass** |
| Memory | < 500MB | **265MB heap / 492MB RSS** | **pass** |
| Range-run collapsing | required for scale | 200k formulas → 2k nodes (**100×**) | **pass** |
| Precedent agreement ≥ 99% | vs `getDirectPrecedents` | **exact** vs naive per-cell reference implementation across the whole corpus; Excel comparison pending a host | **partial — see below** |
| Eval corpus v1 | 30+ workbooks | 17 (4 clean + 13 broken variants) with ground-truth defect labels | **partial — see below** |
| Headless runner | CI-runnable without Excel | simulator + evaluator + graders, all green | **pass** |

Perf harness output (`npm run perf -w engine -- --sheets 20 --rows 2500 --cols 5`):

```
fixture:      20 sheets x 2500 rows x 5 cols (300100 cells, 200000 formulas)
graph build:  8388ms
WIL build:    457ms
TOTAL:        8845ms  (acceptance: <30000ms desktop, <60000ms web)
heap used:    265MB  rss: 492MB (acceptance: <500MB)
collapse:     200000 formulas -> 2000 run nodes (100.0x)
WIL tokens:   ~5849 (budget 6000, truncated=true)
RESULT: PASS
```

## Two correctness bugs found by the test suite, both fixed

Recording these because both would have produced *silently wrong* dependency
analysis — the exact failure mode that destroys workbooks later.

1. **Collapsed runs only resolved their anchor cell's references.** A filled
   row `IS!B2:F2` reported reading only `Assumptions!B2`; the other four
   drivers appeared unused, and changing them would have shown no impact.
   Fixed by expanding relative references across the run's extent while
   leaving `$`-absolute references pinned — precisely the union of what the
   member cells read. Edge count on the 3-statement model went 52 → 98.
   Regression tests cover filled rows, filled columns, absolute refs, mixed
   anchors (`SUM(B$1:B3)` running totals), and cross-sheet fills.

2. **That expansion made node-level cycle detection over-approximate.** A
   cascading fill (`=B1+1` filled right) legitimately reads its own range, so
   Tarjan reported phantom circular references — 5 of them on the clean
   3-statement model. Fixed by treating SCCs as *candidates* and verifying
   each against actual per-cell dependencies before reporting; cascading fills
   are now recorded as `readsOwnRange` instead. Clean models report zero
   cycles; genuine cycles (self, two-node, three-sheet, multiple independent)
   are still detected.

A third issue was found in the **corpus** rather than the engine: the
3-statement model did not balance because year-1 cash was hardcoded instead
of rolling forward from the cash-flow statement. Now every forecast year ties
to < 1e-6, verified end-to-end through the simulator — which is a genuine
joint test of the evaluator, the graph, and the model.

## Deviations from the gate

- **Corpus is 17 workbooks, not 30+.** The 17 cover every AUD rule with
  labelled ground truth (single-defect variants isolate each rule, multi-defect
  variants are the realistic case). Expanding to 30+ is queued for Phase 2,
  where the audit engine gives each new workbook a scored purpose rather than
  padding the count. Logged as D-014.
- **Precedent agreement is verified against our own naive implementation, not
  Office.js.** The `getDirectPrecedents` comparison needs a live Excel host and
  is on the sideload checklist (Phase 0 gate item, still open). The headless
  check is the stronger guard for what it can cover: it proves run collapsing
  and reference expansion never change the answer, which is the failure mode
  the Excel comparison exists to catch. Two properties are asserted — no
  missed precedents on any sampled cell, and no spurious precedent nodes.
- **Chunked extraction** ships as the client-side `range.read` executor from
  Phase 0 (`addin/src/excel/chunks.ts`, unit-tested); the WIL-driving bulk
  extractor that walks a whole workbook is Phase 2 work, since the audit
  dashboard is its first real consumer.

## What the simulator can and cannot do

`SUPPORTED_FUNCTIONS` lists 60 functions — everything the corpus uses. An
unsupported function evaluates to `#NAME?` **and is reported** in
`recalculate().unsupportedFunctions`, so a corpus that outgrows the evaluator
fails loudly instead of scoring against a wrong number. It is not a
spreadsheet engine and is never presented as one.

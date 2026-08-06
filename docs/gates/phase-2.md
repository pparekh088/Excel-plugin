# Phase 2 gate report — Audit engine + read-only UI

Date: 2026-08-06 · Branch: `claude/ledger-excel-agent-8dqvfb`

Gate per handoff §10: AUD-001..011, health dashboard, trace-highlighting,
zero-LLM mode. **Gate conditions: ≥95% precision on the audit corpus; a
deliberately broken 15-sheet DCF yields the correct top-5 issues; demo runs
with network to the LLM disabled.**

## Results

| Gate condition | Target | Actual | Status |
| --- | --- | --- | --- |
| Precision on audit corpus | ≥ 95% | **100%** (38/38 findings trace to a real defect) | **pass** |
| Recall | not specified | **100%** (29/29 injected defects found) | **pass** |
| False positives on clean models | implied 0 | **0** across 4 clean baselines | **pass** |
| Broken 15-sheet DCF top-5 | correct | no noise in top 5; all 5 roots in top 8; criticals first | **pass** |
| Zero-LLM operation | required | `llmCallsUsed === 0` asserted in tests **and** in the eval, which fails the run if non-zero | **pass** |
| Rules implemented | AUD-001..011 | 001–011 (AUD-012+ judgment rules deferred, see below) | **pass** |
| Health score + dashboard | required | score, band, severity grouping, per-finding explanation and evidence | **pass** |
| Trace highlighting | required | implemented with exact fill restore | **partial — needs Excel to verify** |

Corpus eval (`npm run eval:audit -w engine`):

```
precision 100.0%  (38/38 findings related to a real defect)
recall    100.0%  (29/29 injected defects found)
clean baselines: 0 finding(s) across 4 clean workbooks
GATE (precision >= 95%): PASS
```

Engine tests: 951 green. Add-in tests: 29 green (includes an integration test
that exercises the exact import path and call sequence the task pane uses).
Typecheck clean in both packages; production webpack build succeeds.

## How precision is actually measured

Naive address matching is the wrong measure for an audit engine, and using it
would have made these numbers meaningless. Three reasons, all of which
appeared on the first eval run:

1. **One defect legitimately trips several rules.** A magic number typed over
   a filled row is both a hardcode (AUD-002) and a pattern break (AUD-001).
2. **Errors propagate.** One injected `#REF!` makes every downstream cell an
   error cell. A report that hid the blast radius would be worse, not better.
3. **Findings anchor at run or group granularity** — a cycle finding anchors
   at one member of the cycle; a volatile finding at a collapsed run's first
   cell.

So a finding counts against precision only when it is unrelated to any
injected defect: not at an injection site, not downstream of one, not covering
one. On clean workbooks nothing is injected, so **every** finding is a false
positive — which is where the gate genuinely bites, and why the four clean
baselines matter as much as the fourteen broken ones.

## Four precision problems the eval caught

Recording these because each was found by measurement, not inspection, and
each would have made the product feel unreliable in a pilot.

1. **AUD-001 missed anomalies in fragmented blocks.** The original rule
   compared neighbouring nodes with matching bounds; once other defects punch
   holes in a block, the surviving runs are irregular rectangles and the
   comparison fails. It now inspects the cells *flanking* a suspect.
2. **AUD-001 fired on coincidentally-similar rows.** In a financial model,
   vertically adjacent rows are different line items that can share an R1C1
   signature by chance — gross profit and EBITDA are both "sum of the two rows
   above". The rule now additionally requires the suspect's formula not to
   recur along the perpendicular axis, which is what distinguishes a fill
   pattern from a coincidence.
3. **AUD-003 flagged legitimate seed values.** A year-0 units figure at the
   head of a forecast row is correct modelling. A plug is *flanked* by
   formulas on both sides; a seed only has formulas after it. Position is the
   discriminator.
4. **One deleted precedent produced fifteen identical critical findings**,
   burying every other issue. AUD-004 now reports only root errors and states
   how far the error spread; findings are additionally deduplicated by address
   so a cell tripping several rules appears once with the rest attached
   (`alsoFlaggedBy`). The 15-sheet DCF report went from 60 findings to 6.

Point 4 is the difference between a report a reviewer reads and one they
close.

## The 15-sheet DCF fixture

`engine/src/corpus/bigDcf.ts` builds a realistic deal model: 6 segment
build-ups → consolidation → free cash flow → valuation, plus sensitivity,
comparables, working capital, checks and notes sheets (15 total, 200+
formulas). Five defects of deliberately different severity are planted:
a circular EV/equity loop, a `#REF!` in consolidated revenue, a plugged EBIT,
a capex line pointing at the wrong driver, and a hardcoded tax rate.

The tests assert something stronger than "found them somewhere": no noise in
the top five, all five roots within the top eight, criticals ranked first,
propagated errors grouped under their root, and the whole report ≤ 12 findings.
(The five roots occupy the top of the list rather than exactly the first five
slots because the planted `#REF!` also breaks a downstream tie-out check —
a correct high-severity finding of its own.)

## Not verifiable in this environment

- **Trace highlighting on a real grid.** `addin/src/excel/highlight.ts`
  snapshots each cell's existing fill before painting and restores exactly
  that on clear, so clicking "Trace" can never leave a user's formatting
  altered. The restore path needs a live host to verify, and is on the
  sideload checklist along with the still-open Phase 0 items.
- **Bulk extraction against a real workbook.** `addin/src/excel/extract.ts`
  is written to INV-5 (chunked ≤10k cells/sync, `untrack()` per chunk,
  used-range bounded, `formulas` never `formulasLocal`) and its chunk planner
  is unit-tested, but end-to-end timing on a 200k-formula workbook must be
  measured on a host.

## Deviations

- **AUD-012+ (model-risk judgment rules) are not implemented.** They are the
  only LLM-assisted rules in §6 and are explicitly "flagged, never
  auto-fixed". Shipping them now would put an LLM dependency into the wedge
  feature whose entire value is running without one. They belong with the
  agent runtime in Phase 3, where the gateway and cost metering already exist.
  Logged as D-016.
- **"Fix all safe issues" is modelled but not wired.** `safeAutoFixes()`
  returns only LOW-risk fixes and every rule declares its risk tier — AUD-001's
  fix is HIGH because it overwrites an existing formula. Applying them
  requires the change-set engine (INV-2/INV-3: propose → preview → approve →
  apply → log), which is Phase 3. Offering a one-click fix without snapshots
  and rollback would violate the invariants. Logged as D-017.

# Phase 3 gate report — Agent + change sets

Date: 2026-08-06 · Branch: `claude/ledger-excel-agent-8dqvfb`

Gate per handoff §10: full tool surface, planner/executor/verifier/repair,
change-set preview/apply/rollback, drift check, `_AI_Log`.
**Gate conditions: eval edit-task success ≥ 85%; cells-destroyed = 0 across
the suite; rollback restores values + formulas byte-identical on the corpus.**

## Results

| Gate condition | Target | Actual | Status |
| --- | --- | --- | --- |
| Edit-task success rate | ≥ 85% | **100%** (10/10) | **pass** |
| Cells destroyed across the suite | 0 | **0** | **pass** |
| Rollback byte-identical on the corpus | required | **exact** on all 18 corpus workbooks, fingerprinting every cell's value, formula and number format | **pass** |
| Tool surface | ~60 | **68** (23 inspection, 32 mutation, 13 control) | **pass** |
| Planner / executor / verifier / repair | required | implemented, repair budget 3, then rollback offer | **pass** |
| Change-set preview / apply / rollback | required | implemented in engine and add-in, with review UI | **pass** |
| Drift check (INV-8) | required | engine-side and live-workbook versions, both refuse to write on drift | **pass** |
| `_AI_Log` (INV-10) | required | implemented; sheet created on first write, user-deletable | **partial — needs Excel to verify** |

```
--- Edit/build task eval ---
success rate      100.0%  (10/10)
cells destroyed   0
mean tool calls   1.3
GATE (success >= 85%):      PASS
GATE (cells destroyed = 0): PASS
```

Engine 1013 tests · add-in 36 tests · server 23 tests, all green. Typecheck
clean in both TypeScript packages; ruff clean; production build succeeds.

## Two bugs the tests caught, both workbook-corrupting

1. **`translateFormula` mangled sheet names.** Filling a formula shifts its
   relative references, and the first implementation did that with a regex.
   `/[A-Za-z]{1,3}[0-9]+/` matches `eet2` inside `Sheet2!B5`, so filling
   `=Sheet2!B5` down three rows produced `=ShEET5!B8` — a silent reference to
   a sheet that does not exist. It now shifts the parsed AST, which makes
   absolute anchors, defined names, string literals and sheet prefixes
   untouchable by construction. This is exactly why INV-4 exists: we have a
   parser, and string manipulation of formulas is never acceptable.

2. **The repair loop could never converge.** The verifier reconciled *every*
   historical edit against the final workbook state, so once a repair
   overwrote the first attempt's cell, the superseded edit was reported as a
   permanent plan mismatch. Only the last edit per cell describes the intended
   end state. Without this, every repair would have ended in a rollback offer
   even when it succeeded.

## Invariant coverage

- **INV-1 (no freeform execution):** the planner's output is parsed and every
  step checked against the tool catalogue; unknown tools are rejected before a
  change set exists. Tested with a plan emitting `shell.exec` and one emitting
  `vba.run` — both refused, workbook untouched.
- **INV-2 (no silent writes):** the runtime asks for approval twice — once on
  the plan, once on the previewed diff — and applies one change set atomically.
  Trusted-session mode skips prompting only when *every* step is LOW risk;
  tested that a HIGH-risk plan still prompts in a trusted session.
- **INV-3 (snapshot before write):** snapshots capture value, formula and
  number format, including "this cell did not exist". Rollback restores all
  three and deletes cells that were absent.
- **INV-8 (co-authoring):** drift is checked immediately before applying, in
  both the engine and the live-workbook writer. A test mutates the workbook
  between approval and apply and asserts that nothing is written and the
  concurrent edit survives intact. A recalculated *value* on a formula cell is
  correctly not treated as drift.
- **INV-10 (explanation):** `explainChangeSet` states what changed, why, and
  what it affects downstream, including a "these figures are a LOWER BOUND"
  warning when opaque references sit in the blast radius.

## Server side

The LLM gateway routes roles to tiers (planner/critic → strong, executor →
fast, classifier → cheap), meters cost per session, and marks the tool
catalogue and WIL as cacheable since they are large and stable within a
session. Anthropic and OpenAI adapters share one interface and import their
SDKs lazily, so **the server runs with no API keys and no provider packages
installed** — which is what keeps the zero-LLM audit demo honest.

The server also keeps the durable change-set record: a change set that reached
a workbook must be recoverable after the add-in closes, because "what did the
agent do to my model" is an auditor's question.

## Not verifiable in this environment

- **The live-workbook write path.** `addin/src/excel/writer.ts` implements
  drift-check-then-write, calculation suspension re-armed per sync (Q-006),
  batching, and restore-from-snapshot on partial failure. It is tested against
  an Office.js fake that records call ordering, which covers the logic but not
  real payload limits or real calculation semantics. On the sideload checklist.
- **`_AI_Log` sheet creation** — same reason.
- **Model routing quality.** Routing is the handoff's initial policy, not an
  eval result. Scoring real models per task class needs API keys; the seam is
  in place (swap `MockLlmProvider` for a live provider and the same task suite
  scores it), and the eval already reports cost and tool calls per run.

## Deviations

- **The agent runtime lives in the engine, not the server** (D-011 extended).
  The runtime orchestrates deterministic components that all live in
  TypeScript; putting the loop in Python would mean shipping the workbook
  model across the wire on every step. The server owns the gateway, the cost
  meter and the durable record. Logged as D-019.
- **Change-set records are in-process on the server**, not Redis. Sessions are
  already Redis-backed; the records move when the deployment target is
  multi-worker. Noted in PROGRESS rather than as a decision, since nothing
  about the design depends on it.

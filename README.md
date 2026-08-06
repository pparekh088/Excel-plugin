# Ledger — autonomous Excel engineering agent

Ledger is an Excel AI add-in built as an **autonomous spreadsheet engineer**,
not a chat sidebar: it understands entire workbooks, audits financial models
deterministically, plans and applies changes through reviewed and reversible
change sets, and verifies its own work.

The LLM is replaceable. The IP is everything around it.

## What it does

**Audit — the wedge.** Eleven deterministic rules (formula inconsistency,
hardcodes, broken chains, error blast radius, circular references, balance
tie-outs, volatile-function overuse, external/opaque reference inventory, and
more) find real defects in financial models with **zero LLM calls**. That is a
hard property, asserted in tests and in the eval: the whole audit runs inside
any compliance boundary with the network to every model provider switched off.
On an 18-workbook corpus it scores **100% precision and 100% recall**, with
zero findings on four clean baselines.

**Understand.** A formula parser we wrote ourselves (never `getPrecedents()`)
feeds a dependency graph where contiguous cells sharing an R1C1 signature
collapse into one node — 500,000 formulas become 5,000 nodes, which is what
makes whole-workbook analysis possible at all. The Workbook Intelligence Layer
serializes that into a token-budgeted summary that leads with what it *cannot*
see, so the agent never assumes coverage it does not have.

**Change.** Every mutation goes through a change set: propose → preview
(before/after diff, risk tier, downstream impact) → approve → drift check →
apply → verify → repair or roll back → explain. Rollback restores values,
formulas and number formats **byte-identically**, verified across the whole
corpus.

Two things it deliberately will not do. It will not revert a cell somebody
edited after we applied — that cell is reported as a conflict and left exactly
as it is, because undoing our mistake at the cost of their work is the worst
thing this system could do. And it will not delete a sheet it created if
somebody has since put data on it. Both make rollback *partial*, and the report
says which parts, in those words. Structural edits are reversed by inverses
planned before the apply (create→delete, rename→rename back, defineName→restore
the prior definition), because there is no atomic commit in Office.js to lean
on — this is compensation, and it is described as such.

**Never destroy.** Cells-destroyed is a hard gate at zero. A protected sheet,
a merged cell, a concurrent edit by a colleague, or a formula the parser does
not understand each produce a clear refusal — never a partial write and never
a silent one.

## Results

| Gate | Target | Actual |
| --- | --- | --- |
| Audit precision | ≥ 95% | **100%** (recall 100%, 0 false positives on clean models) |
| Edit-task success | ≥ 85% | **100%** |
| Cells destroyed | 0 | **0** |
| Rollback fidelity | byte-identical | **exact** across all 18 corpus workbooks (cells nobody edited after apply — D-026) |
| WIL build, 500k formulas | < 30s, < 500MB | **19.1s**, **221MB** retained |
| 5k-cell `AI.CLASSIFY` drag | under budget | **150** calls, free on re-run |

Tests: engine **1136**, add-in **46**, server **92**. See `PROGRESS.md` for
the full picture including the sideload checklist — the work that genuinely
needs a live Excel host and has not been run.

## Repo layout

```
engine/           Deterministic TypeScript. No DOM, no Office dependencies.
  parser/         Formula parser: full en-US grammar, 713 tests, never throws
  graph/          Dependency graph with range-run collapsing, cycle detection
  wil/            Semantic mapping, token-budgeted summary, visualizer
  audit/          AUD-001..013 — the wedge, zero LLM calls
  changeset/      Snapshot, diff, impact, drift, rollback, hazards
  agent/          68 typed tools, planner/executor/verifier/repair loop
  sim/            Headless Excel: evaluator + recalculation, so CI needs no Excel
  eval/, corpus/  18 workbooks with labelled defects, graders, scoring
addin/            Office.js task pane (React 18 + Fluent UI, TS strict)
  src/excel/      Chunked extraction, change-set writer, trace highlighting
  src/functions/  AI.* custom functions in their own runtime
server/           FastAPI: sessions, auth, LLM gateway, change-set records, telemetry
shared/schemas/   Generated JSON Schemas — the cross-language tool contract
docs/gates/       One report per phase, with measured numbers and open items
```

## Quickstart

```bash
npm ci                      # engine + addin (npm workspaces)

npm test -w engine          # 1136 tests
npm run eval:audit -w engine  # precision/recall against the corpus
npm run eval:edit  -w engine  # agent edit tasks, cells-destroyed
npm run eval:aifn  -w engine  # AI budget + cache gates
npm run perf       -w engine  # WIL build time and memory

cd server
uv venv .venv && uv pip install -p .venv/bin/python -e ".[dev]"
.venv/bin/pytest
.venv/bin/uvicorn ledger_server.main:app --reload --port 8000
```

Sideload the add-in:

```bash
cd addin
npx office-addin-dev-certs install   # first run on a new machine
npm run dev-server                   # https://localhost:3000
npm run start:desktop                # or upload manifest.xml on Excel web
```

The server runs with **no API keys and no provider packages installed** — it
defaults to a mock LLM provider. The audit engine needs no backend at all.

## The invariants

These are not aspirations; they are enforced and tested.

- **INV-1** The LLM emits typed tool calls only. Unknown tools are rejected
  before a change set exists — validated by Zod client-side and by the exported
  JSON Schema server-side.
- **INV-2** No silent writes. Propose → preview → approve → apply → log.
  Trusted-session mode auto-approves only plans that are entirely LOW risk.
- **INV-3** Snapshot before write; structural edits get a planned inverse.
  Rollback refuses to overwrite a cell a human edited after we applied, and
  reports what it could not restore rather than glossing over it.
- **INV-4** We build the dependency graph from our own parse, never
  `getPrecedents()`.
- **INV-5** Chunked, budgeted I/O — ≤10k cells per sync, tracked objects
  released.
- **INV-6** The LLM sees the WIL summary, never a raw grid.
- **INV-7** AI functions are batched, cached, and hard-capped per recalc cycle.
- **INV-8** Drift check immediately before apply; a colleague's edit aborts it.
- **INV-9** Audit, parser, diff and verifier are pure deterministic TypeScript.
- **INV-10** Every change set produces a plain-language explanation.

## Documentation

- `DECISIONS.md` — 27 logged decisions and deviations, with reasoning
- `PLATFORM_QUIRKS.md` — Office.js behaviour traps (product IP)
- `PROGRESS.md` — status, gate results, sideload checklist, known gaps
- `docs/gates/phase-{0..5}.md` — per-phase reports including what failed first

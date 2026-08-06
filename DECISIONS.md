# Decisions log

Deviations from the handoff and consequential choices, with reasoning. Newest last.

## D-001 · Repo layout: `addin/` + `server/` + `shared/schemas/` (2026-08-06)

Two independently buildable packages (npm and uv/pip respectively) plus a
checked-in generated-schema directory as the contract between them. No
monorepo tooling (nx/turbo) until there is a third package that needs it.

## D-002 · Hand-rolled add-in scaffold instead of `yo office` (2026-08-06)

The Yeoman generator is interactive and pins its own template choices. We
reproduce its structure (webpack + dev-certs + manifest + taskpane/commands
entries) by hand: same capabilities, reproducible in CI, no generator cruft.
Sideload flow is identical (`npm run start:desktop`, or manifest upload on web).

## D-003 · Zod is the tool-schema source of truth; server validates generated JSON Schema (2026-08-06)

Per handoff §4, tools are defined once as Zod schemas
(`addin/src/tools/schemas.ts`). `npm run schemas` exports JSON Schema
(draft 2020-12) plus a metadata manifest into `shared/schemas/`, which is
checked in so the Python server needs no Node toolchain. The server loads
these at startup and validates every tool call's params AND result (INV-1 at
the API boundary); the client additionally parses with the same Zod schemas
before sending/executing (both ends of the wire). CI will regenerate and diff
to detect drift. Zod v4's native `z.toJSONSchema` is used — one dependency
fewer than `zod-to-json-schema`.

## D-004 · Tool execution topology: validated envelope, client executes (2026-08-06)

The workbook is only reachable from Office.js inside the add-in, so tools are
*executed* client-side but *validated and logged* server-side. Phase 0
implements the client-initiated flavor (add-in asks the server to validate,
executes, posts the result). The agent runtime (Phase 3) reuses the same
envelope with the server minting the calls and pushing them to the add-in
(transport TBD: SSE or long-poll — decide in Phase 3, log here).

## D-005 · Python 3.11 floor instead of 3.12 (2026-08-06)

Handoff says Python 3.12; this environment provides 3.11. Code is written to
run on both (`requires-python >= 3.11`, no 3.12-only syntax). Revisit when the
deployment target is fixed; nothing in the backend depends on 3.12 features.

## D-006 · Dev auth mode with production guard (2026-08-06)

`LEDGER_AUTH_MODE=dev` (no token) exists for local development because the
real Entra/NAA wiring needs the org's app-registration IDs, which are not in
the repo. The server refuses to start with `auth_mode=dev` when
`LEDGER_ENV=production`, so the bypass cannot reach a deployment. The Entra
path (JWKS validation, audience `api://<client-id>`) is implemented and
awaiting real IDs; the add-in has an `AuthProvider` seam where NAA
(`createNestablePublicClientApplication`) plugs in.

## D-007 · Sessions: Redis when configured, in-memory fallback for dev (2026-08-06)

Same async interface, selected by `LEDGER_REDIS_URL`. The in-memory store is
single-process only and reports itself as `memory` in `/healthz` so it can
never be mistaken for a production setup. TTL 8h on both.

## D-008 · `range.read` include key is `numberFormats` (handoff wrote `formats`) (2026-08-06)

`formats` is ambiguous (fills, fonts, borders are also formats). The layer
maps 1:1 to Office.js `Range.numberFormat`, and INV-3 snapshots are defined as
values + formulas + *number formats*, so the tool key says exactly that.
A future `range.format` mutation tool will handle the broader format spec.

## D-009 · `formulas` grids share the value union (2026-08-06)

Office.js `Range.formulas` returns the *value* (number/bool/string) for
non-formula cells, not an empty string. The schemas encode that honestly:
`values` and `formulas` are both `(string|number|boolean|null)[][]`. Anything
stricter would reject real workbooks on day one.

## D-010 · Icons are generated placeholders (2026-08-06)

`addin/assets/icon-*.png` are programmatically generated (navy "L") so the
manifest validates and sideload works. Replace with real brand assets before
any pilot.

## D-011 · Deterministic engine is a shared TypeScript package, not Python (2026-08-06)

The handoff puts the Workbook Engine and Audit Engine in the FastAPI backend.
We put them in `engine/` — a pure-TypeScript package with no DOM or Office
dependencies — because three consumers need the same code and two of them are
JavaScript: (a) the add-in, for client-side degraded mode when the backend is
unreachable (§3 requires audit rules to still run), (b) the headless eval
harness in CI, (c) the server, via a thin call-out when server-side analysis
is wanted. Implementing the parser and audit rules twice would guarantee
divergence between what the add-in flags offline and what the server flags
online, which is precisely the trust-destroying bug for an audit product.
The Python backend keeps sessions, auth, the LLM gateway, change-set storage,
and the agent loop.

## D-012 · Runs expand relative references; absolute references stay pinned (2026-08-06)

A range-run node stands for many cells, so its dependency set must be the
union of what those cells read. Relative bounds sweep with the fill, `$`-fixed
bounds do not. Without this a filled row reports only its leftmost cell's
precedents (see Phase 1 gate report). The union is a bounding box, which for a
contiguous fill is exact.

## D-013 · Cycle detection verifies candidates at cell level (2026-08-06)

Because runs read their own range in cascading fills, node-level SCCs
over-approximate. Tarjan produces candidates; each is verified against actual
per-cell dependencies before being reported to a user as a circular reference.
Cost is bounded by the candidate's size. A false "your model has a circular
reference" is worse than a slightly slower audit.

## D-014 · Eval corpus starts at 17 workbooks, not 30+ (2026-08-06)

The gate asks for 30+. We ship 17 (4 clean baselines + 13 broken variants)
that between them exercise every audit rule with labelled ground truth, and
expand in Phase 2 where each addition can be scored rather than counted. Clean
baselines matter as much as broken ones: any finding on a clean model is a
false positive, which is how the ≥95% precision gate is actually enforced.

## D-015 · The simulator reports what it cannot evaluate (2026-08-06)

The headless evaluator implements the ~60 functions the corpus uses. Rather
than returning 0 for anything else, unsupported functions yield `#NAME?` and
are listed in `recalculate().unsupportedFunctions`, so an eval that outgrows
the evaluator fails loudly instead of silently scoring against a wrong value.

## D-016 · AUD-012+ judgment rules deferred to Phase 3 (2026-08-06)

The model-risk heuristics (growth > threshold, negative margins, terminal
assumptions) are the only LLM-assisted rules in §6. Shipping them in Phase 2
would put an LLM dependency inside the wedge feature whose entire value
proposition is running without one — "demo it with the network off" stops
being true the moment one rule needs a model. They land in Phase 3 alongside
the gateway and cost metering, clearly labelled as judgment rather than fact,
and opt-in via `AuditScope.includeJudgment`.

## D-017 · "Fix all safe issues" needs the change-set engine (2026-08-06)

Every rule declares an auto-fix risk tier and `safeAutoFixes()` filters to
LOW, but nothing is applied yet. Applying fixes requires snapshots, preview,
approval and rollback (INV-2, INV-3). A one-click fix button without those is
exactly the silent-write failure the invariants exist to prevent, so the
button ships in Phase 3 with the change-set engine behind it.

## D-018 · The add-in consumes the engine from source, not a built package (2026-08-06)

`ledger-engine` is aliased to `engine/src/index.ts` in webpack, tsconfig and
vitest rather than being built to `dist/` and installed. This guarantees the
audit logic running in the task pane is byte-identical to the one the eval
harness scores — the divergence risk D-011 exists to prevent. Cost is a
slightly slower add-in build; revisit if build time becomes a problem.

## D-019 · The agent runtime lives in the engine, not the server (2026-08-06)

Extends D-011. The plan/execute/verify/repair loop orchestrates components
that are all deterministic TypeScript — parser, graph, change-set engine,
verifier, simulator. Running the loop in Python would mean shipping the
workbook model across the wire on every step, and would fork the logic between
the live add-in path and the eval harness. The server keeps what genuinely
belongs to it: sessions, auth, the LLM gateway with routing and cost metering,
and the durable record of every change set that touched a workbook.

## D-020 · Formula translation goes through the parser, never a regex (2026-08-06)

Filling a formula shifts its relative references. The obvious implementation
is a regex over `[A-Za-z]{1,3}[0-9]+`, and it is wrong: that pattern matches
`eet2` inside `Sheet2!B5`, so filling `=Sheet2!B5` produced `=ShEET5!B8`.
Translation now walks the parsed AST, which makes absolute anchors, defined
names, string literals and sheet prefixes untouchable by construction. Recorded
because the temptation to string-manipulate a formula will recur, and the
answer is always no.

## D-021 · The verifier reconciles only the last edit per cell (2026-08-06)

A repair round legitimately overwrites what the first attempt wrote. Checking
every historical edit against the final state reported the superseded edit as
a permanent plan mismatch, so the repair loop could never converge and every
run ended in a rollback offer. The net effect of a change set is what the user
approved and what must be verified.

## D-022 · No streaming into cells for AI.* functions (2026-08-06)

§8 mentions the custom-functions streaming pattern. None of the five functions
we ship produce progressive output — they return one value — and streaming
would complicate budget accounting (a partially-streamed result is neither
cacheable nor countable) for no user-visible gain. Revisit if a long-running
AI function is added.

## D-023 · AI.FORECAST never touches a language model (2026-08-06)

The forecast is Holt-Winters / Holt / simple exponential smoothing with
parameters chosen by in-sample MAE, computed identically in the engine and the
server. A generated number and a computed number are indistinguishable once
they are in a cell, and in a financial model that is the difference between a
forecast and a guess. A model may narrate the result; it may never produce it.

## D-024 · AI-derived cells are an audit rule, not a UI badge (2026-08-06)

AUD-013 inventories AI.* cells in the same report as every other finding,
escalating from info to medium when a generated value feeds downstream
calculations. Putting it in the audit report rather than a separate panel means
it lands in front of the person signing off on the model, which is who §8 says
demands it. It uses zero LLM calls, so the zero-LLM demo still holds.

## D-025 · The system role is server-owned; workbook content is fenced data (2026-08-06)

The planner used to build its system prompt from a client-supplied tool
catalogue plus the WIL. Both are attacker-reachable — the WIL is generated
*from* the workbook, so every sheet name, defined name, comment and string
literal in it is content somebody else may have written — and putting them in
the system role hands the highest-trust position to the least-trusted input.

The system prompt and tool catalogue are now constants the server and engine
own; supplying either over the wire is a 422. Workbook content travels in a
user turn inside a `<workbook_context>` fence, with fence delimiters stripped
from the content first so it cannot close its own fence. The user's intent goes
last, so the most recent instruction the model reads is the user's.

This is a boundary, not a filter. It does not make injection impossible — a
sufficiently persuasive payload inside the fence may still steer a plan. What
it guarantees is that steering a plan is all it can do: the model can only emit
typed tool calls, and every mutation still goes through preview and approval.
That layering is the actual defence; the fence just removes the free win.

## D-026 · Rollback yields to a human edit rather than winning (2026-08-06)

Rollback used to restore the snapshot unconditionally. The failure mode: we
write F27, the user fixes F27 by hand, our verification fails on some unrelated
cell, the user accepts the rollback offer, and their fix is gone. The undo of
our mistake destroys their work — the single most damaging thing this system
could do, and it looked like correct behaviour from inside the code.

Change sets now record `appliedState`: what each cell held immediately after we
wrote it. Rollback compares live against that, using the same rule as drift
detection (formula for formula cells, value for constants, so a recalculation
is never mistaken for an edit). A cell that still matches is provably ours and
is restored; a cell that differs is somebody's newer work and is left exactly
as it is, reported as a `RollbackConflict` naming what we wrote and what is
there now. Missing `appliedState` is treated as a conflict too — absent
evidence is not evidence of safety.

`force` restores over a conflict, and exists solely to serve an explicit user
decision made with the conflict list in front of them. Nothing in the agent
loop sets it.

Two honest limits. First, this is post-apply protection only: we detect that a
cell changed, not who changed it, so an edit made by a *second* Ledger session
reads as a human edit and is likewise left alone (which is the safe direction).
Second, a conflict means the rollback is genuinely partial — some of our change
survives in the workbook — so `ok: true` on a report with conflicts means
"completed as designed", not "the workbook is back to its prior state". The
transcript says so in those words.

## D-027 · Structural edits are reversed by planned inverses, not by snapshots (2026-08-06)

A change set that created a sheet and wrote formulas onto it was half a
transaction: the cells were reversible, the sheet was not. Rollback restored
the cells, listed the sheet under `unrestorable`, and reported `ok: true` — so
the workbook was left holding an empty sheet the user never asked for, and the
report called that a success.

Cell edits are reversible because a cell has a prior state to snapshot. A sheet
that did not exist has none, so reversing it needs an explicit inverse:
create→delete, rename→rename back, defineName→delete (or restore the definition
we overwrote), createTable→delete. Those inverses are planned at propose time,
against the pre-apply workbook, because afterwards we can no longer tell
whether we were the ones who created the thing. `planCompensation` is pure and
lives in the engine; each host executes the plan (`applyCompensation` for the
simulator, `undoStructural` for Office.js), which keeps the deterministic layer
free of Excel.

Every inverse carries the same guard cell rollback uses — reverse it only if we
can still prove it is ours:

- a sheet we created that now holds cells we did not write is NOT deleted
- a rename is not undone if a sheet has reappeared under the old name
- a defined name that has been repointed since is left where it is
- an edit that applied as a no-op (the sheet already existed) has no inverse at
  all, and saying so is the whole point of planning at propose time

Live applies now run one sync per structural edit rather than batching them.
That costs a round trip per edit — there are rarely more than a couple — and
buys the ability to say exactly which ones landed when a later write throws.
Batched, a failure would leave us knowing only that "some of them" applied,
and an inverse we cannot aim is worse than none.

Two honest limits. This is compensation, not a transaction: there is no atomic
commit in Office.js, so a failure between the structural sync and the cell
writes is recovered by running inverses, and an inverse can itself fail (each
one is caught and reported rather than stranding the rest). And a *guarded*
inverse that refuses is a genuinely partial rollback — the report names what
stayed and why, rather than counting it as restored.

## D-028 · A repair inherits the approval it stays inside (2026-08-06)

The repair loop was the one place a write reached the workbook without the
user seeing it. Verification fails, the model proposes a corrective plan, the
plan is applied — up to three times, to any cell, at any risk tier. INV-2 says
no silent writes, and a repair is exactly when that matters most: a repair only
happens because the agent has already got something wrong.

Two obvious policies are both wrong. Approving every repair individually turns
the ordinary case — "the formula I just wrote needs a small correction" — into
a dialog the user learns to click through, which is how approval fatigue
destroys the value of asking at all. Approving none lets a failed verification
become a licence to edit anywhere.

So the boundary is scope, not count. The user approved a specific set of cells
at a specific risk tier. A repair confined to those cells, at no more risk than
they accepted, is finishing the job they said yes to and applies without
asking. A repair that reaches a cell outside that set, makes any structural
change, or escalates the risk tier goes back to them with the reasons named.
Rejecting leaves the workbook as it is and the usual rollback offer follows.

`classifyRepair` is pure and deterministic — the decision is computed from the
two change sets, never asked of the model. A boundary a model could talk its
way past would be decorative.

One subtlety that took a bug to find. Risk must be judged against the state the
user was SHOWN, not against the live workbook: once our change is applied,
every correction to our own formula reads as "overwriting existing logic" and
scores high, so every repair would escalate and the policy would collapse into
"always ask". `riskOf` now takes an optional prior-state map, and the repair
classifier passes the approved change set's snapshots.

There is deliberately no "never ask" setting. `repairApproval: "always"` exists
for callers who want every fix confirmed; the in-scope default is the floor.

## D-029 · The agent loop runs on a WorkbookHost, not on a Workbook (2026-08-06)

The agent loop was written directly against the in-memory model and our own
evaluator. That made CI possible and made its results weaker than they looked:
every gate in this repo — 100% edit success, zero cells destroyed, byte-identical
rollback — was proved against the SIMULATOR, and the simulator is our model of
Excel, not Excel. The loop had no way to run in a real workbook at all.

`WorkbookHost` is the seam: read, checkDrift, apply, refresh, rollback.
`SimulatorWorkbookHost` is what CI exercises; `OfficeJsWorkbookHost` is what
ships. `runAgent` still accepts a bare `Workbook` and wraps it in the simulator
host, so every existing test and eval is unchanged.

The step that matters is `refresh`. Verification used to run against our own
model immediately after we wrote to it, which proves the model self-consistent
and nothing else. It now runs against whatever the host reports AFTER
recalculating. On Office.js that is a real re-extraction, so verification can
catch the failures that only exist in the real host: value coercion on write,
implicit intersection, and a table's calculated column rewriting our formula
(Q-010). Those are exactly the ones a user would describe as "the agent broke
my workbook", and no amount of simulator testing can find them.

It re-extracts the whole workbook rather than re-reading the touched cells,
because verification is about blast radius — a formula we wrote can put #REF!
three sheets away. It is the expensive step of a run; the chunked extractor
keeps it inside the INV-5 budget.

Hosts also declare `capabilities`, and `RunResult` carries `host`. A pass on
the simulator and a pass on Excel are different claims, so the transcript ends
with which one it was, in words: "verified against the headless simulator,
which is our model of Excel, not Excel".

New outcome `write-failed`: the host can now refuse (protected sheet, drift
detected live), and the loop reports that rather than proceeding to verify a
write that never happened.

The limit worth stating: the Office.js host is covered by a fake in tests, and
a fake cannot reproduce the very behaviours that justify the abstraction.
Sideload checklist item 13 is the only thing that closes it.

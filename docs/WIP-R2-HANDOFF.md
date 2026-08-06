# WIP handoff — second review round (R2)

**Status: partial. Items 1–5 and 7 are implemented; item 6 is untouched; most
new behaviour has no test yet.** This file is the working state for whoever
picks this up next. Delete it when R2 is closed.

Branch: `claude/ledger-excel-agent-8dqvfb`
Base of this work: `5efb5fb` ("Review P2: align the trust language …")
PR: pparekh088/Excel-plugin#1

The reviewer's verdict on `5efb5fb` was *"good architecture with ~6
production-path bugs to close"* — keep the architecture, fix the bugs. Nothing
in here is a redesign.

---

## What is DONE in this commit

### 1. Repair loop planned against the stale pre-apply workbook — FIXED
`engine/src/agent/runtime.ts`

`runAgent` read the workbook once, then after applying did
`verified = await host.refresh(changeSet)` — but the repair path still used the
original `workbook`. On the simulator those are the same object so it worked by
accident; on Office.js they are separate extractions, and the consequences were:

- the repair's snapshots claimed the *pre-apply* values as "before", so the live
  writer's own drift check saw our first write as a foreign edit and refused;
- the refused repair was merged into the master change set regardless, putting
  edits in the audit record that are not in the workbook and snapshots in the
  rollback plan for writes that never landed.

Now: the executor, `classifyRepair` and `proposeChangeSet` all take `verified`;
`host.apply(repairSet)` result is checked and a refused repair breaks the loop
without merging. `mergeChangeSets` also concatenates `compensation` (it was
dropping the repair's structural inverses).

Tests: `engine/test/agent/repairLive.test.ts` (4 tests) — a `LiveLikeHost` whose
reads are clones and whose `apply` runs its own snapshot drift check, i.e. the
two properties that make Office.js differ from the simulator.

### 2. `/ai/batch` response contract mismatch — client half FIXED, contract test MISSING
`addin/src/functions/functions.ts`

Server returns `{ok, value, error}` per item; the client declared
`Array<string|number|boolean>` and passed it straight to `AiCoordinator`, which
caches and returns it to Excel. TypeScript missed it because the JSON was
force-cast. In production a cell would render `[object Object]`.

Now: `unwrapAiResult()` maps each envelope → scalar, `#AI_ERROR!` on `ok:false`
or a non-scalar value; length mismatch throws. **Still to do: the end-to-end
contract test the reviewer asked for** (server response → client → coordinator).

### 3. `AI.FORECAST` unauthenticated — FIXED
`addin/src/api/backend.ts` (new), `addin/src/functions/functions.ts`

`aiForecast` had its own bare `fetch` with only `Content-Type`, so under
`auth_mode=entra` it was the one AI.* function still returning 401 after auth
landed everywhere else. All backend traffic now goes through
`postToBackend()` — one place to attach the token, nothing to forget.

### 4. `checkLiveDrift` throws on a sheet the change set creates — FIXED
`addin/src/excel/writer.ts`

Drift ran *before* structural edits and used `worksheets.getItem(snap.sheet)`.
For the very common `sheet.create("DCF")` + write into `DCF!A1:F40`, the
snapshots name a sheet that correctly does not exist yet. Now a shared
`existingSheets()` probe uses `getItemOrNullObject`, and *absent sheet + absent
snapshot* is the clean expected state; *absent sheet + snapshot with content* is
real drift (someone deleted it).

### 5. Structural rollback overwrote human edits — FIXED
`addin/src/excel/writer.ts`, `engine/src/changeset/structural.ts`

- `restoreName` (Office.js) deleted whatever was there and restored the old
  definition without comparing `expectRefersTo` — the deterministic engine got
  this right, the live path did not. Now guarded, and honours `force`.
- `deleteTable` had no evidence at all that the table was still ours.
  `CompensatingOp.deleteTable` gained `expectRange`; both executors refuse when
  the live table spans a different range (someone resized it — deleting the
  object would break their structured references).

### 7. Cost scoping not wired from the client — FIXED both ends
`addin/src/api/backend.ts`, `addin/src/taskpane/App.tsx`,
`server/ledger_server/routers/ai.py`

- Client: the task pane publishes the server-issued session via
  `setActiveSessionId()`; `callBackend` sends it. Module-level under the shared
  runtime, mirrored to `localStorage` for hosts without one.
- Server: `validated_session_id()` looks the session up and checks
  `principal_subject` before metering against it. An unknown or foreign session
  degrades to the caller's own bucket rather than letting a caller manufacture
  meter buckets or bill someone else's session.

### Cleanups
- `readTouchedCells` decided `absent` with `typeof rawFormula !== "string"`;
  Excel returns `""` for a blank cell's formula, so blanks read as *present* and
  produced false rollback conflicts. Now: blank = no value AND no `"="` formula.
- Repair approval passed the **original** `changeSet` to `approver()`, so a UI
  rendering the structured diff showed the wrong changes. Now the repair
  `ChangeSet` is proposed first and its own preview/diff is what gets approved.
- `server/tests/test_routing.py` import order (ruff) — this was failing CI and
  skipping the whole server test job.

---

## What is NOT done

### Item 6 — OpenAI-only deployment is still broken (untouched)

`server/ledger_server/llm.py`. Three separate defects, all confirmed:

1. **Server will not start with only `OPENAI_API_KEY`.** `build_gateway()` adds
   providers by key presence → `{mock, openai}`. `forced` is empty and
   `len(providers) == 2`, so it skips the mock path and calls
   `_routes_from_env()`, which returns the **Anthropic** `DEFAULT_ROUTES`.
   `LlmGateway.__init__` then raises `RoutingError: Route for 'planner' names
   provider 'anthropic', which is not configured`.
2. **`LEDGER_LLM_PROVIDER=openai` does nothing.** Only `"mock"` is special-cased
   (`llm.py`, in `build_gateway`).
3. **Pricing follows the old provider.** `_routes_from_env()` swaps
   `provider`/`model` but keeps the base route's `input_cost_per_mtok`,
   `output_cost_per_mtok` and `reasoning_effort`, so cost accounting silently
   uses Anthropic prices for OpenAI calls.

Suggested shape (the reviewer's words: *provider/model/pricing/reasoning need to
be a single coherent route configuration*):

- a per-provider default route table, e.g. `DEFAULT_ROUTES_BY_PROVIDER:
  dict[ProviderName, dict[AgentRole, ModelRoute]]`, each entry carrying its own
  correct pricing;
- `build_gateway()` picks the table by `LEDGER_LLM_PROVIDER` when set, else by
  which key is present (exactly one real provider → that provider's table; both
  → keep Anthropic as today; none → mock);
- `_routes_from_env()` overrides must carry pricing too — either look the model
  up in a price table or require the full route in the env value. Do not inherit
  pricing across a provider switch.

Reviewer explicitly asked for **tests covering OpenAI-only, Anthropic-only,
mixed, and forced-provider startup**. `server/tests/test_routing.py` is the
place.

### Missing tests for work that IS done

The implementations landed; these are unproven and the reviewer will look for
them. Roughly in priority order:

- **Batch contract, end to end** (item 2): server `BatchResponse` shape →
  `callBackend` → `AiCoordinator` → scalar in the cell. Include `ok:false` →
  `#AI_ERROR!`, and a length mismatch. `addin/test/` — note `auth.test.ts`
  fixtures were updated to the real `{ok, value}` envelope and are a usable
  starting point.
- **`AI.FORECAST` sends the bearer token** (item 3). Trivial with the existing
  `postToBackend` seam.
- **create-sheet + populate drift** (item 4): a change set with `createSheet` +
  cells on it must produce *no* drift when the sheet is absent, and must report
  drift when a snapshot with content finds its sheet gone. `addin/test/writer.test.ts`
  — the Office.js fake there already keeps a sheet registry.
- **Structural rollback guards on the live path** (item 5): name repointed since
  → not restored; table resized since → not deleted; `force` overrides both.
  Engine side has equivalents in `engine/test/changeset/structural.test.ts` to
  mirror.
- **Session scoping** (item 7): `/ai/batch` with another principal's session ID
  must meter into the caller's own bucket, not theirs. `server/tests/`.

### Docs not yet updated
No `DECISIONS.md` entry for any of the above. When item 6 is done, one entry
covering the R2 round is probably right (D-031), plus a note in
`PLATFORM_QUIRKS.md` about Excel's `formulas` grid returning `""` for blanks and
echoing values for non-formula cells — that is the root of two separate bugs now
(the `absent` misdetection here, and the earlier `isNullObject` snapshot issue in
Q-013).

---

## Verification state as of this commit

| Check | Result |
| --- | --- |
| `engine` tsc | clean |
| `engine` vitest | 1166 passed |
| `addin` tsc | clean |
| `addin` vitest | 65 passed |
| `server` ruff | clean (was the CI blocker) |
| `server` pytest | 92 passed, 2 skipped |
| `eval:audit` | precision 100%, recall 100% — PASS |
| `eval:edit` | 100% (10/10), 0 cells destroyed — PASS |
| `eval:aifn` | budget, cache and breaker — PASS |

`eval:edit` exercises the repair path directly, so it is the gate most likely to
catch a regression from the runtime change. Re-run all three before pushing.

## The four scenarios the reviewer will check next

Stated verbatim in the R2 review as what a final pass will focus on:

1. create a new sheet and populate it,
2. failed verification followed by a repair,
3. human edit followed by rollback,
4. authenticated AI.* functions under **both** OpenAI-only and Anthropic-only
   deployments.

(1)–(3) are covered by the fixes here but only (2) has real test coverage.
(4) is blocked on item 6.

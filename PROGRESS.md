# Progress

All six phases (0–5) are implemented and their gates measured. Every gate that
can be measured without Excel passes — and "without Excel" is doing real work
in that sentence: the gates run against the headless simulator, which is our
model of Excel rather than Excel itself (D-029). Sideload item 15 is what
converts them into claims about the real host. The items that genuinely require a live
Excel host are listed in **§ Sideload checklist** below and in each phase's
gate report — they are the honest remainder, not hidden.

## Test and gate summary

| Package | Tests | Typecheck | Lint |
| --- | --- | --- | --- |
| `engine/` | **1162** | clean | — |
| `addin/` | **65** | clean | — |
| `server/` | **92** | — | ruff clean |

| Eval | Gate | Result |
| --- | --- | --- |
| `eval:audit` | precision ≥ 95% | **100%** precision, **100%** recall, 0 findings on 4 clean baselines |
| `eval:edit` | success ≥ 85%, cells destroyed = 0 | **100%** (10/10), **0** destroyed |
| `eval:aifn` | 5k drag under budget, re-run cache ≥ 60% | 5000 cells in **150** calls, **100%** cache hits on re-run |
| `perf` (500k formulas) | < 30s, < 500MB | **19.1s**, **221MB** retained, 100× collapse |

## Phase status

| Phase | Deliverable | Gate | State |
| --- | --- | --- | --- |
| 0 | Skeleton: add-in + FastAPI + validated tool envelope | works on desktop AND web | **code complete**, sideload verification open |
| 1 | Formula parser, WIL, dependency graph, eval scaffold | §5 numbers; 500 parser tests | **pass** (713 parser tests) |
| 2 | Audit engine AUD-001..011, read-only UI | ≥95% precision; broken 15-sheet DCF top-5; zero-LLM | **pass** (100% precision) |
| 3 | Tool surface, change sets, agent runtime | edit success ≥85%; 0 destroyed; byte-identical rollback | **pass** (100%, 0, exact on cells nobody edited after apply — see D-026) |
| 4 | AI.* functions, visualizer, cost dashboard | 5k drag under budget, cache ≥60% | **pass** |
| 5 | Hardening: 500k stress, co-authoring, locale, telemetry | — | **pass** |

## Sideload checklist — the honest remainder

None of these can be executed in a headless Linux container. They need Excel
on Windows/Mac plus a browser for Excel web.

1. `npx office-addin-dev-certs install`, then `npm run dev-server`.
2. Sideload `addin/manifest.xml` on **Excel desktop** (`npm run start:desktop`)
   and on **Excel web** (Insert → Add-ins → Upload My Add-in).
3. Task pane opens from the Home-tab **Ledger** button on both hosts.
4. **Audit tab** on a real workbook: extraction completes, findings render,
   **Trace** highlights the dependency chain *and* restores the original fill
   exactly on Clear.
5. **Tools tab** with the backend running: `range.read` round trip returns
   values and formulas from a live sheet.
6. Backend stopped → offline banner, audit still works (client-side).
7. **Change set**: propose → preview → apply on a live workbook; verify
   calculation suspension, batching, and `_AI_Log` sheet creation.
8. **Drift**: edit a cell from a second session between preview and apply;
   confirm the apply is refused and nothing is written.
8b. **Rollback conflict**: apply a change set, edit one of the written cells by
   hand, then roll back. The hand-edited cell must keep the human value and be
   reported as a conflict; the other cells must restore (D-026). The headless
   equivalent passes against the Office.js fake; the live host is what proves
   the re-read sees an uncommitted in-cell edit.
9. **Custom functions**: `=AI.CLASSIFY(...)` registers under the `AI`
   namespace, batches, and returns `#AI_BUDGET!` past the budget.
10. `office-addin-manifest validate` against Microsoft's service (blocked by
    this container's egress proxy; XML well-formedness checked locally).
11. **Precedent agreement** vs `getDirectPrecedents` on 100 random formula
    cells (≥99%) — the headless equivalent is exact against a naive
    reference implementation, but the Excel comparison is the stated gate.
12. **Locale**: confirm `Range.formulas` is en-US on a de-DE host; record
    findings in `PLATFORM_QUIRKS.md`.
13. **Merged cells / protected sheets**: confirm whether Office.js throws or
    silently ignores; record in `PLATFORM_QUIRKS.md`.
14. Mac-specific: mixed content (https task pane → http backend), Q-005.
15. **Agent loop end to end on the real host**: drive `runAgent` with an
    `OfficeJsWorkbookHost` against a live workbook and confirm verification
    reads back what Excel actually stored. This is the item that turns every
    gate below from "proved against our simulator" into "proved against
    Excel" — specifically coercion on write, implicit intersection, and a
    table calculated column rewriting a formula we set (Q-010, D-029).
16. **Authenticated AI.* functions**: with `LEDGER_ENTRA_*` set and the backend
    on `auth_mode=entra`, confirm NAA acquires a token silently through the
    host, that the SHARED RUNTIME lets the custom-functions context reuse the
    task pane's cached token, and that `=AI.CLASSIFY(...)` succeeds. Then
    confirm a signed-out state shows the "sign in" message rather than a
    generic failure (D-030). Nothing in this path has run against a real Entra
    tenant, and NAA support varies by Office build.

## Known gaps and deferred work

- **Entra/NAA auth** is implemented end to end — task pane and custom
  functions, shared runtime, token cache, production build refusing to fall
  back to anonymous (D-030) — but needs the org's real app-registration IDs and
  has never run against a live tenant. Dev mode is guarded on both sides so it
  cannot reach production (D-006).
- **AUD-012 model-risk judgment rules** are deferred (D-016). They are the only
  LLM-assisted rules; shipping them earlier would have put a model dependency
  inside the zero-LLM wedge.
- **"Fix all safe issues"** is modelled (`safeAutoFixes`, risk tiers on every
  rule) but not wired to a button (D-017). The change-set engine it needs now
  exists, so this is a small piece of UI work.
- **Change-set records and cost meters are in-process** on the server; they
  move to Redis when the deployment target is multi-worker.
- **Eval corpus is 18 workbooks**, not 30+ (D-014). Every audit rule has
  labelled ground truth; growth should be driven by scored gaps.
- **Model routing is the handoff's initial policy**, not an eval result.
  Scoring real models per task class needs API keys; the seam is in place.
- **Nightly Office.js integration tests** on a Windows VM via Playwright are
  not set up — no such VM is available here.
- **Icons are generated placeholders** (D-010).

## Where things live

```
engine/   deterministic TypeScript: parser, graph, WIL, audit, change sets,
          agent runtime, simulator, evals, corpus  (D-011)
addin/    Office.js task pane, extraction, writer, highlighting, AI functions
server/   FastAPI: sessions, auth, LLM gateway, change-set records, telemetry
shared/   generated JSON Schemas — the cross-language tool contract
docs/gates/  one report per phase, with measured numbers and open items
```

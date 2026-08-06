# Platform quirks

Office.js behavior differences and traps, with repro notes. This file is
product IP — keep it current. Status legend:
**observed** = reproduced by us; **reported** = documented by Microsoft/community,
to be reproduced when we hit that phase.

## Q-001 · `getPrecedents()` is unusable for graph building — status: reported (handoff §2 INV-4)

Per-range API, slow at scale, throws on cells with no precedents
(`ItemNotFound` instead of empty), and has known correctness bugs on Excel
web. Consequence: the dependency graph is built by bulk-reading formulas and
parsing them ourselves (Phase 1). `getDirectPrecedents()` is used only as
spot-verification on ≤ 20 cells. Repro to capture in Phase 1: empty-precedent
throw + a web-vs-desktop diff on 3D references.

## Q-002 · ~5 MB payload limit per request on Excel web — status: reported

`RichApi.Error: The request payload size has exceeded the limit` when a
single sync moves too much data. Consequence: INV-5 — all reads/writes are
chunked at ≤ 10k cells per sync (`addin/src/excel/chunks.ts`). Desktop is more
forgiving; we keep one budget for both.

## Q-003 · `Range.formulas` vs `Range.formulasLocal` — status: reported

`formulas` is always en-US grammar (comma separators, en-US function names);
`formulasLocal` follows the user locale (`;` separators, translated names,
e.g. `SUMME` in German Excel). Everything internal (parser, graph, audit,
LLM) speaks en-US: **only ever read/write `formulas`**. `formulasLocal` may be
used for *display* in the UI, never for round-tripping. Locale test matrix is
Phase 5.

## Q-004 · `Range.formulas` returns raw values for non-formula cells — status: reported

The `formulas` grid is not `string[][]`: a constant cell yields its value
(number/boolean/string). Schemas and parser input types must accept the full
union (see D-009). Detection of "is a formula" is `typeof v === "string" &&
v.startsWith("=")`.

## Q-005 · Mixed content: https task pane → http://localhost backend — status: reported

Excel web and Windows desktop (WebView2/Chromium) treat `http://localhost` as
a potentially-trustworthy origin, so the dev backend on plain http works. Mac
desktop uses WKWebView (Safari rules) which blocks it. Mac dev needs the
backend behind https (mkcert/caddy) or a tunnel. Capture actual behavior per
host during Phase 0 sideload verification.

## Q-006 · `suspendApiCalculationUntilNextSync()` scope — status: reported

Suspension only lasts until the next `context.sync()`, so multi-sync write
batches must re-suspend per sync cycle. Applies from Phase 3 (write executor).
Also note `application.calculationMode = "Manual"` is workbook-global and
user-visible — prefer suspension, always restore in a `finally`.

## Q-007 · Tracked-object leaks degrade long sessions — status: reported

Proxy objects accumulate per `Excel.run` unless released; long-lived loops
over many ranges must call `untrack()` on ranges they are done with (we do,
per chunk, in `rangeRead.ts`) or memory + sync latency grow unboundedly on
web.

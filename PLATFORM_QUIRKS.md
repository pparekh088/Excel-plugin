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

## Q-008 · Merged cells: only the anchor is writable — status: reported

Writing to a non-anchor cell of a merged area is **silently ignored** on some
hosts rather than throwing. That is the worst failure mode available: the
change set reports success and nothing happened. Consequence: hazard detection
blocks any write to a non-anchor cell before the change set is applied
(`engine/src/changeset/hazards.ts`). Repro to capture on a host: merge A1:D1,
write to C1 via `range.values`, observe whether the value lands or is dropped.

## Q-009 · Protected sheets: never unprotect on the user's behalf — status: reported

Writes to a protected sheet throw. Office.js exposes `worksheet.protection`,
so we detect it and refuse with a message telling the user to unprotect it. We
deliberately do NOT call `protection.unprotect()` even when we could: silently
removing a control someone put there is precisely the helpful-but-wrong action
that destroys trust in an audit tool.

## Q-010 · Table calculated columns fight direct writes — status: reported

Writing a formula into one cell of a table's calculated column may propagate
to the whole column or be reverted, depending on host and version. Treated as
a warning rather than a block, with the user told to check the result. Repro:
create a table with a formula column, write a different formula into one body
cell, observe.

## Q-011 · V8 heap readings overstate memory without a forced GC — status: observed

Measuring `process.memoryUsage().heapUsed` right after building the graph over
a 500k-formula workbook reported 885MB; forcing a GC first showed 213MB
actually retained. The difference is transient parse trees awaiting
collection. Any memory acceptance number must state whether it means retained
or peak — we report both (`engine/scripts/perf-harness.ts`), and the budget is
about retained. Reporting peak as retained would have produced a false failure.

## Q-012 · Post-apply state has to be re-read, not assumed — status: designed around

Rollback compares the live cell against what we wrote (D-026), which means we
need to know what Excel actually stored — not what we sent it. Excel coerces on
write: a string that parses as a date becomes a serial, `=SUM(A:A)` may come
back with implicit-intersection markers, and a formula written into a table's
calculated column may be rewritten entirely (Q-010). Assuming the edit's payload
*is* the post-apply state would make every one of those look like a human edit
and block a legitimate rollback.

So `applyChangeSet` re-reads the touched cells after the write and stores that
(`captureLiveAppliedState`). Costs one extra batched sync per apply. The
simulator has no coercion and records the state inline in `applyToWorkbook`.

Not yet verified on a live host: whether the re-read observes a cell the user is
*currently* editing in the formula bar but has not committed. If Excel reports
the pre-edit value there, the user's in-flight edit would be restored over on
rollback. On the sideload checklist as item 8b.

## Q-013 · `isNullObject` is a snapshot, not a live view — status: designed around

`getItemOrNullObject(...)` returns a placeholder whose `isNullObject` resolves
at `load()`/`sync()` time and then stays put. Code that adds the sheet and
re-checks the same handle sees the value from BEFORE the add — which is what
`writeAiLog` relies on to decide whether to write headers.

Recorded because the test double got this wrong first: a live getter made
`isNullObject` flip after `worksheets.add`, and the _AI_Log header row silently
stopped being written. The fake now snapshots on `load()`, matching the host.

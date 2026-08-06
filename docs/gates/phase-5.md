# Phase 5 gate report — Hardening

Date: 2026-08-06 · Branch: `claude/ledger-excel-agent-8dqvfb`

Scope per handoff §10: giant-workbook stress (500k formulas), co-authoring
chaos tests, locale (formula separators, localized function names), merged
cells, protected sheets (detect, ask for unprotect, never fail silently),
telemetry.

## Results

| Item | Target | Actual | Status |
| --- | --- | --- | --- |
| 500k-formula stress | build within budget | **19.1s**, 100× run collapse | **pass** |
| Memory at 500k formulas | < 500MB | **221MB retained** (888MB peak, see below) | **pass** |
| Co-authoring chaos | drift never silently overwritten | 6 scenarios: concurrent edit, delete, insert, interleaved rollback, recalc-is-not-drift, multi-cell drift | **pass** |
| Locale normalization | `;` separators, localized names | de-DE and fr-FR round-trip, string literals and quoted sheet names untouched | **pass** |
| Merged cells | never fail silently | non-anchor writes blocked with an explanation | **pass** |
| Protected sheets | detect, ask, never fail silently | blocked with "unprotect it and run again"; never attempts to remove protection itself | **pass** |
| Telemetry | required | event schema with enforced redaction | **pass** |

Full suite at the end of Phase 5: engine **1090** tests, add-in **36**, server
**56**. Typecheck clean in both TypeScript packages, ruff clean, all four
evals passing.

## The memory number was wrong, and the fix was measuring properly

The first 500k run reported **885MB against a 500MB budget** — a fail. Two
rounds of work followed, and only the second was the real answer.

**Round 1 — genuine optimizations, negligible effect.** I converted the hot
cell maps from `"row,col"` string keys to packed numeric keys (Excel is 16384
columns wide, so `row * 16384 + col` is lossless), and stopped retaining a
parse tree per formula cell — the graph now parses once for the R1C1
signature, discards the tree, then re-parses only the few thousand run
*anchors*. Both are correct improvements and both stayed in. Memory moved
885MB → 906MB. Essentially nothing.

**Round 2 — measure instead of guess.** Forcing a GC before reading the heap
showed the truth:

```
workbook model only:   129MB retained
workbook + graph:      213MB retained
```

The 900MB figure was **uncollected garbage**, dominated by transient parse
trees, not retained state. The harness now reports both numbers because they
answer different questions: *retained* is what a long session holds and is
what the budget is about; *peak* matters on a memory-constrained host but V8
collects it under pressure. Reporting peak as if it were retained would have
been a false failure, and quietly relaxing the budget to make it pass would
have been worse.

The harness now runs under `--expose-gc` and says "unknown" rather than
reporting peak as retained when it cannot measure properly.

## Never fail silently

The theme of this phase. Each hazard produces a message a user can act on:

- **Protected sheet** → blocked. *"Sheet X is protected, so nothing can be
  written to it. Unprotect it and run this again — I will not attempt to
  remove the protection myself."* Silently unprotecting someone's sheet is
  exactly the kind of helpful-but-wrong action that destroys trust.
- **Merged cell, non-anchor** → blocked, because writing there is *silently
  ignored* by the host. This is the worst failure mode available: the change
  set reports success and nothing happened.
- **Missing sheet** → blocked, noting it may have been renamed or deleted.
- **Table calculated column** → warning, not a block: Excel may propagate or
  revert the edit, so the user is told to check the result.

Hazards run before the change-set preview, so a blocked plan never reaches a
write. `blocked-hazard` is a distinct run outcome, not an error.

## Locale (Q-003)

Everything internal speaks en-US; Office.js gives that for free via
`Range.formulas`, and `formulasLocal` is never round-tripped. The normalizer
handles the two places locale still leaks in — a user typing a formula in
their own locale, and displaying one back — with character-wise conversion
that is aware of string literals and quoted sheet names, because a `;` inside
`"a;b"` is the user's text, not syntax.

A bug found here is worth recording: the escaped-quote logic (`""` inside a
string) was inverted, so `=WENN(A1;"a;b";"c,d")` converted the separator
*inside* the second literal. Caught by the test that asserts string contents
survive conversion.

Function-name tables are a representative subset (de-DE, fr-FR), not
exhaustive — thousands of entries per locale. An unrecognized name passes
through unchanged rather than being corrupted, which is the safe failure.

## Telemetry: enforced, not trusted

`FORBIDDEN_FIELDS` refuses any property that could carry workbook content —
formulas, values, sheet names, addresses, intents, explanations, prompts,
diffs, file names — and refuses strings over 200 characters, since long
strings are how content leaks. Nested payloads are checked recursively.
Session ids are hashed, never stored. Disabled by default; `LEDGER_TELEMETRY=on`
enables it, and "off" means no events are constructed at all.

Eleven tests assert the negative property, which is the one that matters: a
financial model is the most confidential thing most users own, and an audit
tool that leaks it is worthless however good the audit is.

## Not verifiable in this environment

Everything on the sideload checklist from Phases 0–4 remains open, and Phase 5
adds:

- **Real locale hosts.** The normalizer is unit-tested, but Excel running in
  de-DE/fr-FR is the only way to confirm `formulas` really is en-US on those
  hosts and that our display conversion matches what Excel shows.
- **Real merged cells and protected sheets.** Hazard detection reads our model;
  whether Office.js throws or silently ignores in each case needs a host, and
  the answers belong in PLATFORM_QUIRKS.
- **Co-authoring with an actual second user.** The chaos tests simulate
  concurrent edits against the model. Real co-authoring adds timing and
  event-ordering behaviour we cannot reproduce headlessly.

# Phase 0 gate report — Skeleton

Date: 2026-08-06 · Branch: `claude/ledger-excel-agent-8dqvfb`

Gate per handoff §10: Yeoman-equivalent add-in (React+TS) sideloaded on
desktop + web; FastAPI backend with MSAL/NAA auth pattern; session service;
hello-world tool round trip (`range.read` → backend → response).
**Gate condition: works on Excel desktop AND web.**

## Verified in this environment (headless Linux container, no Excel)

| Check | Result |
| --- | --- |
| Server unit/integration tests (`pytest`) — envelope, INV-1 rejects (unknown tool 404, bad params/result 422 with JSON-pointer errors), double-resolution 409, session privacy, Redis store | **12/12 pass** (Redis tests ran against a live `redis-server`) |
| Add-in unit tests (`vitest`) — chunk planner double-coverage/gap proofs incl. 16384-wide rows and 200k-cell case; Zod schema accept/reject incl. A1 grammar and the formulas-hold-raw-values union | **25/25 pass** |
| TypeScript strict typecheck (`tsc --noEmit`, `strict` + `noUncheckedIndexedAccess`) | clean |
| Webpack production build → `dist/` with taskpane/commands bundles, assets, manifest with dev→prod URL transform applied | builds (3 size warnings from Fluent UI bundle weight — acceptable for a task pane, revisit at Phase 4 polish) |
| Zod → JSON Schema export (`npm run schemas`) → `shared/schemas/` (draft 2020-12, strict, checked in) | works; server loads it at startup |
| Live round trip against running server (Redis-backed): healthz → create session → `range.read` call validated → executor-shaped result validated + logged → session log readback | **pass** (see transcript below) |
| Manifest XML well-formedness + GUID | pass |
| Server ruff lint | clean |

Round-trip transcript (abridged):

```
healthz: {status: ok, session_store: redis, tools: [range.read]}
session: ses_8cda406190e642b4803e
invalid params rejected: [{path: /a1, message: "'nope!' does not match ..."}]
accepted: tc_fc607a670dc348e8bbfc read validated
ack: {tool_call_id: tc_..., status: completed, summary: "range.read completed"}
session log OK
```

## NOT verifiable in this environment — required to close the gate

These need a machine with Excel (Windows/Mac desktop + a browser for Excel
web) and were **not** run here. The gate is **open** until they pass:

1. `npx office-addin-dev-certs install` then `npm run dev-server`; sideload
   `manifest.xml` on **Excel desktop** (`npm run start:desktop`) and **Excel
   web** (Insert → Add-ins → Upload My Add-in).
2. Task pane opens from the Home-tab **Ledger** button on both hosts.
3. With `uvicorn ledger_server.main:app --port 8000` running: task pane shows
   backend **online**, session id, and **Run range.read** returns values +
   formulas from a real sheet on both hosts (exercises the Office.js executor
   leg that cannot run headless).
4. Backend stopped → task pane shows the offline banner and stays usable
   (graceful degradation per §3).
5. `office-addin-manifest validate` against Microsoft's online service
   (blocked by this container's egress proxy; XML checked locally instead).
6. Mac-specific: confirm Q-005 (https task pane → http://localhost backend)
   behavior and record it in PLATFORM_QUIRKS.md.

## Deviations logged

D-005 (Python 3.11 floor), D-006 (dev-auth guard, Entra IDs pending),
D-008/D-009 (schema naming/typing), D-010 (placeholder icons) — see
`DECISIONS.md`.

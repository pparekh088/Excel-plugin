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

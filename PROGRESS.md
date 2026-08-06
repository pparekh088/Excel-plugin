# Progress

## Phase 0 — Skeleton (current)

Goal: sideloadable add-in (desktop + web) + FastAPI backend + hello-world
tool round trip (`range.read`) through the validated envelope.
Gate report: `docs/gates/phase-0.md`.

| Item | Status |
| --- | --- |
| Repo layout (`addin/` + `server/` + `shared/schemas/`) | done |
| FastAPI backend: sessions (Redis + in-memory), dev/Entra auth seam, healthz | done |
| Tool-call envelope: validate params/result against generated JSON Schema (INV-1) | done |
| Zod tool schemas source of truth + `npm run schemas` export | done |
| `range.read` executor with chunked reads (INV-5) + unit-tested chunk planner | done |
| Task pane UI: backend status, session, round-trip runner, offline degradation | done |
| Manifest (desktop + web sideload) + placeholder icons | done |
| Server tests (envelope, validation rejects, Redis store) | done |
| In-container verification (pytest, vitest, tsc, webpack build, live HTTP round trip) | done — see gate report |
| Sideload verification on real Excel desktop + web | **pending — needs a machine with Excel** |
| Entra/NAA wiring with real app-registration IDs | pending (seam in place, D-006) |

## Phase 1 — WIL + formula parser + eval scaffold (next)

Not started. First moves: formula parser TDD harness + grammar corpus
(500 cases), chunked bulk extraction, range-run collapsed DAG, WIL serializer,
eval corpus v1 + headless runner.

## Phase 2 — Audit engine + read-only UI

Not started.

## Phase 3 — Agent + change sets

Not started. (Envelope already designed for server-minted calls — D-004.)

## Phase 4 — AI custom functions + polish

Not started.

## Phase 5 — Hardening

Not started. `PLATFORM_QUIRKS.md` seeded with known traps to verify.

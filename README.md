# Ledger — autonomous Excel engineering agent

Ledger is an Excel AI add-in built as an **autonomous spreadsheet engineer**,
not a chat sidebar: it understands entire workbooks (Workbook Intelligence
Layer), plans and applies changes through reviewed, reversible change sets,
audits financial models deterministically (zero LLM calls), and verifies its
own work.

Core invariants (see the handoff + `DECISIONS.md`): the LLM only ever emits
**schema-validated typed tool calls** (no freeform code against the workbook),
**no silent writes** (propose → preview → approve → apply → log), snapshots
before every write, chunked/budgeted I/O, and a deterministic audit engine
that runs entirely without the LLM.

## Repo layout

```
addin/            Office.js add-in — React 18 + Fluent UI v9, TypeScript strict
  src/tools/      Zod tool schemas (SOURCE OF TRUTH for the typed tool surface)
  src/excel/      Office.js executors (chunked I/O per INV-5)
  src/api/        Backend client + tool round-trip orchestration
  src/taskpane/   Task pane UI
  manifest.xml    Sideloadable add-in manifest (dev URLs: https://localhost:3000)
shared/schemas/   Generated JSON Schemas + tool manifest (checked in; `npm run schemas`)
server/           Agent API — FastAPI, Python >= 3.11
  ledger_server/  config, auth (dev/Entra), sessions (Redis/in-memory),
                  schema registry (INV-1 enforcement), tool-call envelope
docs/gates/       Per-phase acceptance gate reports
DECISIONS.md      Deviations + consequential choices, with reasoning
PROGRESS.md       Live phase/status tracker
PLATFORM_QUIRKS.md Office.js cross-platform behavior notes (product IP)
```

## Quickstart (dev)

Backend:

```bash
cd server
uv venv .venv && uv pip install -p .venv/bin/python -e ".[dev]"
.venv/bin/uvicorn ledger_server.main:app --reload --port 8000
# tests:
.venv/bin/pytest
```

Add-in:

```bash
cd addin
npm install
npm test && npm run typecheck   # unit tests (chunk planner, schemas)
npm run dev-server               # serves https://localhost:3000
# First run on a new machine: npx office-addin-dev-certs install
```

Sideload:

- **Excel desktop (Win/Mac):** `npm run start:desktop` (registers the manifest
  and launches Excel), or add `manifest.xml` via Insert → Add-ins → sideload.
- **Excel web:** open a workbook → Insert → Add-ins → Upload My Add-in →
  `addin/manifest.xml` (dev server must be running with trusted certs).

Open the **Ledger** button on the Home tab → task pane shows backend status,
creates a session, and runs the `range.read` round trip: local Zod parse →
server-side JSON Schema validation (INV-1) → chunked Office.js read (INV-5) →
result validated and logged server-side.

Regenerating tool schemas after editing `addin/src/tools/schemas.ts`:

```bash
cd addin && npm run schemas   # writes shared/schemas/, commit the diff
```

## Status

Phase 0 (skeleton) — see `PROGRESS.md` and `docs/gates/phase-0.md`.

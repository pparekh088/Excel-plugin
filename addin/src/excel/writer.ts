/**
 * Change-set writer: applies a change set to the live workbook (INV-2, INV-3,
 * INV-5, INV-8).
 *
 * Ordering matters and is not negotiable:
 *   1. re-read the touched cells and diff against the snapshot (drift check)
 *   2. suspend calculation
 *   3. write in batches, re-suspending per sync (Q-006)
 *   4. restore calculation, recalculate
 *   5. re-read to confirm what actually landed
 *
 * If the drift check fails, nothing is written at all. If a write throws
 * partway, we roll back from the snapshot rather than leaving the workbook
 * half-changed.
 */

import type {
  CellSnapshot,
  ChangeSet,
  Edit,
  RollbackConflict,
  RollbackOptions,
} from "ledger-engine";
import { isCellEdit } from "ledger-engine";

export interface ApplyProgress {
  written: number;
  total: number;
}

export interface ApplyOptions {
  onProgress?: (progress: ApplyProgress) => void;
  /** Cells written per sync — keeps each request well under the payload cap. */
  batchSize?: number;
}

export interface ApplyResult {
  ok: boolean;
  cellsWritten: number;
  /** Set when the apply was refused or reverted, with the reason. */
  failure?: string;
  /** True when a partial write had to be rolled back. */
  rolledBack?: boolean;
}

export interface LiveDriftEntry {
  address: string;
  expectedFormula?: string;
  actualFormula?: string;
  expectedValue: unknown;
  actualValue: unknown;
}

const DEFAULT_BATCH = 500;

/** Re-read every snapshotted cell from the live workbook and compare (INV-8). */
export async function checkLiveDrift(changeSet: ChangeSet): Promise<LiveDriftEntry[]> {
  const drift: LiveDriftEntry[] = [];
  await Excel.run(async (context) => {
    const loaded: Array<{ snap: CellSnapshot; range: Excel.Range }> = [];
    for (const snap of changeSet.snapshots) {
      const sheet = context.workbook.worksheets.getItem(snap.sheet);
      const range = sheet.getRangeByIndexes(snap.row, snap.col, 1, 1);
      range.load(["values", "formulas"]);
      loaded.push({ snap, range });
    }
    await context.sync();

    for (const { snap, range } of loaded) {
      const value = (range.values as unknown[][])[0]?.[0] ?? null;
      const rawFormula = (range.formulas as unknown[][])[0]?.[0];
      const formula =
        typeof rawFormula === "string" && rawFormula.startsWith("=") ? rawFormula : undefined;

      // A formula cell drifts only if its FORMULA changed; its value moves on
      // every recalculation and that is not someone else's edit.
      const formulaChanged = formula !== snap.formula;
      const valueChanged =
        snap.formula === undefined && !Object.is(normalize(value), normalize(snap.value));

      if (formulaChanged || valueChanged) {
        drift.push({
          address: `${snap.sheet}!${snap.row},${snap.col}`,
          expectedFormula: snap.formula,
          actualFormula: formula,
          expectedValue: snap.value,
          actualValue: value,
        });
      }
      range.untrack();
    }
  });
  return drift;
}

function normalize(value: unknown): unknown {
  return value === "" ? null : value;
}

export async function applyChangeSet(
  changeSet: ChangeSet,
  options: ApplyOptions = {}
): Promise<ApplyResult> {
  const drift = await checkLiveDrift(changeSet);
  if (drift.length > 0) {
    return {
      ok: false,
      cellsWritten: 0,
      failure:
        `Aborted before writing anything: ${drift.length} cell(s) changed since this plan was ` +
        `made (${drift.slice(0, 5).map((entry) => entry.address).join(", ")}). ` +
        `Someone else may be editing this workbook.`,
    };
  }

  const batchSize = options.batchSize ?? DEFAULT_BATCH;
  const cellEdits = changeSet.edits.filter(isCellEdit);
  let written = 0;

  try {
    await Excel.run(async (context) => {
      // Structural edits first: later cell writes may target new sheets.
      for (const edit of changeSet.edits) {
        if (isCellEdit(edit)) continue;
        applyStructural(context, edit);
      }
      await context.sync();

      for (let index = 0; index < cellEdits.length; index += batchSize) {
        const batch = cellEdits.slice(index, index + batchSize);
        // Suspension lasts only until the next sync, so re-arm every batch.
        context.application.suspendApiCalculationUntilNextSync();

        for (const edit of batch) {
          const sheet = context.workbook.worksheets.getItem(edit.sheet);
          const range = sheet.getRangeByIndexes(edit.row, edit.col, 1, 1);
          switch (edit.kind) {
            case "setFormula":
              range.formulas = [[edit.formula ?? ""]];
              break;
            case "setValue":
              range.values = [[edit.value ?? null]];
              break;
            case "setNumberFormat":
              range.numberFormat = [[edit.numberFormat ?? "General"]];
              break;
            case "clear":
              range.clear(Excel.ClearApplyTo.contents);
              break;
          }
        }
        await context.sync();
        written += batch.length;
        options.onProgress?.({ written, total: cellEdits.length });
      }

      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    // Record what actually landed. Rollback compares the live cell against
    // this to tell our write from a later human edit; without it, rollback
    // has to refuse every cell.
    changeSet.appliedState = await captureLiveAppliedState(changeSet);
  } catch (error) {
    // A partial write is worse than no write: put it back.
    const message = error instanceof Error ? error.message : String(error);
    try {
      await restoreSnapshots(changeSet.snapshots);
      return {
        ok: false,
        cellsWritten: written,
        rolledBack: true,
        failure: `Write failed after ${written} cell(s) (${message}). Restored from snapshot.`,
      };
    } catch (restoreError) {
      const restoreMessage =
        restoreError instanceof Error ? restoreError.message : String(restoreError);
      return {
        ok: false,
        cellsWritten: written,
        rolledBack: false,
        failure:
          `Write failed after ${written} cell(s) (${message}) AND the restore also failed ` +
          `(${restoreMessage}). The workbook may be in a partially changed state — ` +
          `use Excel's undo.`,
      };
    }
  }

  return { ok: true, cellsWritten: written };
}

function applyStructural(context: Excel.RequestContext, edit: Edit): void {
  if (isCellEdit(edit)) return;
  switch (edit.kind) {
    case "createSheet":
      if (edit.name) context.workbook.worksheets.add(edit.name);
      break;
    case "renameSheet":
      if (edit.sheet && edit.newName) {
        context.workbook.worksheets.getItem(edit.sheet).name = edit.newName;
      }
      break;
    case "defineName":
      if (edit.name && edit.refersTo) {
        context.workbook.names.add(edit.name, edit.refersTo);
      }
      break;
    case "createTable":
      if (edit.sheet && edit.range) {
        const table = context.workbook.tables.add(
          `${edit.sheet}!${edit.range}`,
          edit.hasHeaders ?? true
        );
        if (edit.name) table.name = edit.name;
      }
      break;
  }
}

/** One live cell as Excel currently reports it. */
interface LiveCell {
  value: unknown;
  formula?: string;
  numberFormat?: string;
  absent: boolean;
}

/** Re-read the touched cells from the live workbook in one batched pass. */
async function readTouchedCells(changeSet: ChangeSet): Promise<Map<string, LiveCell>> {
  const live = new Map<string, LiveCell>();
  await Excel.run(async (context) => {
    const loaded: Array<{ snap: CellSnapshot; range: Excel.Range }> = [];
    for (const snap of changeSet.snapshots) {
      const sheet = context.workbook.worksheets.getItem(snap.sheet);
      const range = sheet.getRangeByIndexes(snap.row, snap.col, 1, 1);
      range.load(["values", "formulas", "numberFormat"]);
      loaded.push({ snap, range });
    }
    await context.sync();

    for (const { snap, range } of loaded) {
      const value = (range.values as unknown[][])[0]?.[0] ?? null;
      const rawFormula = (range.formulas as unknown[][])[0]?.[0];
      const numberFormat = (range.numberFormat as unknown[][])[0]?.[0];
      live.set(cellKey(snap.sheet, snap.row, snap.col), {
        value: normalize(value),
        ...(typeof rawFormula === "string" && rawFormula.startsWith("=")
          ? { formula: rawFormula }
          : {}),
        ...(typeof numberFormat === "string" ? { numberFormat } : {}),
        // Excel has no "absent" — an empty cell reads back as "".
        absent: normalize(value) === null && typeof rawFormula !== "string",
      });
      range.untrack();
    }
  });
  return live;
}

function cellKey(sheet: string, row: number, col: number): string {
  return `${sheet.toUpperCase()}!${row},${col}`;
}

function liveAddress(sheet: string, row: number, col: number): string {
  return `${sheet}!R${row + 1}C${col + 1}`;
}

/**
 * Capture what the live workbook holds right after a successful apply.
 *
 * The simulator records this inside `applyToWorkbook`; the Office.js path
 * writes through the real API, so it has to re-read to find out what Excel
 * actually stored (coercions, autocorrect, implicit intersection).
 */
export async function captureLiveAppliedState(changeSet: ChangeSet): Promise<CellSnapshot[]> {
  const live = await readTouchedCells(changeSet);
  return changeSet.snapshots.map((snap) => {
    const cell = live.get(cellKey(snap.sheet, snap.row, snap.col));
    return {
      sheet: snap.sheet,
      row: snap.row,
      col: snap.col,
      value: (cell?.value ?? null) as CellSnapshot["value"],
      ...(cell?.formula !== undefined ? { formula: cell.formula } : {}),
      ...(cell?.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
      absent: cell?.absent ?? true,
    };
  });
}

export interface LiveRollbackReport {
  restoredCells: number;
  conflicts: RollbackConflict[];
  unrestorable: string[];
  ok: boolean;
}

/** Same rule as the engine: formulas for formula cells, values for constants. */
function changedSinceApply(applied: CellSnapshot, current: LiveCell | undefined): boolean {
  if (applied.absent !== (current?.absent ?? true)) return true;
  if (applied.formula !== current?.formula) return true;
  if (applied.formula === undefined && !Object.is(current?.value ?? null, normalize(applied.value))) {
    return true;
  }
  return false;
}

/**
 * User-initiated rollback against the live workbook (INV-3).
 *
 * Cells edited by a human after we applied are LEFT ALONE and reported as
 * conflicts. Reverting somebody's newer work because our verification failed
 * would be the most damaging thing this add-in could do. `force` exists only
 * to serve an explicit decision the user made with the conflict list in front
 * of them.
 */
export async function rollbackChangeSet(
  changeSet: ChangeSet,
  options: RollbackOptions = {}
): Promise<LiveRollbackReport> {
  const live = await readTouchedCells(changeSet);
  const appliedByCell = new Map<string, CellSnapshot>();
  for (const state of changeSet.appliedState ?? []) {
    appliedByCell.set(cellKey(state.sheet, state.row, state.col), state);
  }

  const conflicts: RollbackConflict[] = [];
  const restorable: CellSnapshot[] = [];

  for (const snap of changeSet.snapshots) {
    const key = cellKey(snap.sheet, snap.row, snap.col);
    const current = live.get(key);
    const applied = appliedByCell.get(key);
    const currentState = {
      value: (current?.value ?? null) as RollbackConflict["current"]["value"],
      ...(current?.formula !== undefined ? { formula: current.formula } : {}),
    };

    // No post-apply state means we cannot prove the cell is still ours; the
    // safe reading of missing evidence is "do not touch it".
    if (!applied) {
      conflicts.push({
        address: liveAddress(snap.sheet, snap.row, snap.col),
        reason:
          "No post-apply state was recorded for this cell, so I cannot tell whether " +
          "someone has edited it since. Left as-is.",
        applied: { value: null },
        current: currentState,
      });
      continue;
    }

    if (!options.force && changedSinceApply(applied, current)) {
      conflicts.push({
        address: liveAddress(snap.sheet, snap.row, snap.col),
        reason: "Cell was edited after the AI change; your edit was kept.",
        applied: {
          value: applied.value,
          ...(applied.formula !== undefined ? { formula: applied.formula } : {}),
        },
        current: currentState,
      });
      continue;
    }
    restorable.push(snap);
  }

  const restoredCells = restorable.length > 0 ? await restoreSnapshots(restorable) : 0;

  const unrestorable: string[] = [];
  for (const edit of changeSet.edits) {
    if (isCellEdit(edit)) continue;
    unrestorable.push(
      `Structural change (${edit.kind}${edit.name ? ` "${edit.name}"` : ""}) is not reversed ` +
        `by rollback; undo it by hand if unwanted.`
    );
  }

  return { restoredCells, conflicts, unrestorable, ok: true };
}

/**
 * Restore cells from snapshots UNCONDITIONALLY — the live-workbook half of
 * rollback (INV-3).
 *
 * This does no conflict checking, so call it only where nobody can have edited
 * in between: the immediate recovery from a failed partial write. Every
 * user-initiated rollback must go through `rollbackChangeSet`.
 */
export async function restoreSnapshots(snapshots: CellSnapshot[]): Promise<number> {
  let restored = 0;
  await Excel.run(async (context) => {
    for (let index = 0; index < snapshots.length; index += DEFAULT_BATCH) {
      const batch = snapshots.slice(index, index + DEFAULT_BATCH);
      context.application.suspendApiCalculationUntilNextSync();

      for (const snap of batch) {
        const sheet = context.workbook.worksheets.getItem(snap.sheet);
        const range = sheet.getRangeByIndexes(snap.row, snap.col, 1, 1);
        if (snap.absent) {
          range.clear(Excel.ClearApplyTo.contents);
        } else if (snap.formula !== undefined) {
          range.formulas = [[snap.formula]];
        } else {
          range.values = [[snap.value as string | number | boolean]];
        }
        if (snap.numberFormat !== undefined) {
          range.numberFormat = [[snap.numberFormat]];
        }
        restored++;
      }
      await context.sync();
    }
    context.workbook.application.calculate(Excel.CalculationType.full);
    await context.sync();
  });
  return restored;
}

const LOG_SHEET = "_AI_Log";
const LOG_HEADERS = ["Timestamp", "Change set", "Intent", "Risk", "Summary", "Detail"];

/**
 * Append the change set's explanation to the _AI_Log sheet (INV-10).
 * The sheet is created on first use and is user-deletable — it is a record,
 * not machine state, so nothing breaks if someone removes it.
 */
export async function writeAiLog(changeSet: ChangeSet, explanation: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const existing = sheets.getItemOrNullObject(LOG_SHEET);
    existing.load("isNullObject");
    await context.sync();

    const sheet = existing.isNullObject ? sheets.add(LOG_SHEET) : existing;
    if (existing.isNullObject) {
      sheet.getRangeByIndexes(0, 0, 1, LOG_HEADERS.length).values = [LOG_HEADERS];
      sheet.getRangeByIndexes(0, 0, 1, LOG_HEADERS.length).format.font.bold = true;
      await context.sync();
    }

    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["isNullObject", "rowIndex", "rowCount"]);
    await context.sync();

    const nextRow = used.isNullObject ? 1 : used.rowIndex + used.rowCount;
    sheet.getRangeByIndexes(nextRow, 0, 1, LOG_HEADERS.length).values = [
      [
        new Date().toISOString(),
        changeSet.id,
        changeSet.intent,
        changeSet.risk,
        changeSet.summary,
        explanation,
      ],
    ];
    await context.sync();
  });
}

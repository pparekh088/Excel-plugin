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
  CompensatingOp,
  Edit,
  RollbackConflict,
  RollbackOptions,
  StructuralEdit,
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
  /** Structural changes reversed as part of that rollback, in plain language. */
  structuralReversed?: string[];
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
  const structuralEdits = changeSet.edits.filter(
    (edit): edit is StructuralEdit => !isCellEdit(edit)
  );
  // planCompensation emits one op per structural edit, in order, so the two
  // arrays line up by index.
  const compensation = changeSet.compensation ?? [];
  const landed: CompensatingOp[] = [];
  let written = 0;

  try {
    // Structural edits first (later cell writes may target new sheets), and
    // one sync each: batching them would leave us unable to say which of them
    // landed if the sync throws, and an inverse we cannot aim is useless.
    for (let index = 0; index < structuralEdits.length; index++) {
      await Excel.run(async (context) => {
        applyStructural(context, structuralEdits[index]!);
        await context.sync();
      });
      const op = compensation[index];
      if (op) landed.push(op);
    }

    await Excel.run(async (context) => {
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
      // Cells first, then the structural inverses — a sheet we created can
      // only be removed once its contents are back where they belong. Nobody
      // can have edited between our write and this recovery, so the guards
      // are bypassed: leaving half a transaction behind is the worse outcome.
      const structural = await undoStructural(landed, { force: true });
      return {
        ok: false,
        cellsWritten: written,
        rolledBack: true,
        structuralReversed: structural.reversed,
        failure:
          `Write failed after ${written} cell(s) (${message}). Restored from snapshot` +
          (structural.reversed.length > 0
            ? ` and reversed ${structural.reversed.length} structural change(s).`
            : ".") +
          (structural.unreversed.length > 0
            ? ` Could NOT reverse: ${structural.unreversed.join(" ")}`
            : ""),
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
  /** Structural changes actually reversed. */
  reversedStructural: string[];
  unrestorable: string[];
  ok: boolean;
}

export interface StructuralUndoResult {
  reversed: string[];
  unreversed: string[];
}

/**
 * Execute compensating operations against the live workbook, in reverse order.
 *
 * The engine plans these (`planCompensation`) — it knows what the inverse of
 * each structural edit is and what has to be true for the inverse to be safe.
 * This function is only the Office.js executor for that plan, which is why the
 * guards read live state here rather than being decided in the engine.
 *
 * `force` skips the guards and is used in exactly two places: recovering from
 * a failed apply (nobody can have edited in between) and an explicit user
 * decision made with the reasons in front of them.
 */
export async function undoStructural(
  ops: CompensatingOp[],
  options: { force?: boolean } = {}
): Promise<StructuralUndoResult> {
  const reversed: string[] = [];
  const unreversed: string[] = [];

  for (const op of [...ops].reverse()) {
    try {
      switch (op.kind) {
        case "none":
          unreversed.push(`${op.describes}: ${op.reason}`);
          break;

        case "deleteSheet": {
          const outcome = await undoCreateSheet(op.sheet, options.force ?? false);
          (outcome.ok ? reversed : unreversed).push(outcome.message);
          break;
        }

        case "renameSheet": {
          const outcome = await undoRename(op.from, op.to, options.force ?? false);
          (outcome.ok ? reversed : unreversed).push(outcome.message);
          break;
        }

        case "deleteName": {
          const outcome = await undoDefineName(op.name, op.expectRefersTo, options.force ?? false);
          (outcome.ok ? reversed : unreversed).push(outcome.message);
          break;
        }

        case "restoreName": {
          await Excel.run(async (context) => {
            const existing = context.workbook.names.getItemOrNullObject(op.name);
            existing.load(["isNullObject", "formula"]);
            await context.sync();
            if (!existing.isNullObject) existing.delete();
            context.workbook.names.add(op.name, op.refersTo);
            await context.sync();
          });
          reversed.push(`Restored the defined name "${op.name}" to ${op.refersTo}.`);
          break;
        }

        case "deleteTable": {
          const outcome = await undoCreateTable(op.name);
          (outcome.ok ? reversed : unreversed).push(outcome.message);
          break;
        }
      }
    } catch (error) {
      // One inverse failing must not strand the rest: keep going and say so.
      const message = error instanceof Error ? error.message : String(error);
      unreversed.push(`Undoing ${op.kind} failed: ${message}`);
    }
  }

  return { reversed, unreversed };
}

interface UndoOutcome {
  ok: boolean;
  message: string;
}

async function undoCreateSheet(name: string, force: boolean): Promise<UndoOutcome> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) {
      return { ok: false, message: `Sheet "${name}" is already gone; nothing to delete.` };
    }

    // The guard: the sheet must be empty. Our own cells were restored (and so
    // cleared, since they did not exist before), so anything left is somebody
    // else's work and the sheet stays.
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["isNullObject", "address"]);
    await context.sync();

    if (!used.isNullObject && !force) {
      return {
        ok: false,
        message:
          `Sheet "${name}" was created by this change but is no longer empty (${used.address}), ` +
          `so it was NOT deleted. Delete it by hand if you want it gone.`,
      };
    }
    sheet.delete();
    await context.sync();
    return { ok: true, message: `Deleted the sheet "${name}" this change created.` };
  });
}

async function undoRename(from: string, to: string, force: boolean): Promise<UndoOutcome> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(from);
    const clash = context.workbook.worksheets.getItemOrNullObject(to);
    sheet.load("isNullObject");
    clash.load("isNullObject");
    await context.sync();

    if (sheet.isNullObject) {
      return {
        ok: false,
        message: `Sheet "${from}" no longer exists, so the rename back to "${to}" was skipped.`,
      };
    }
    if (!clash.isNullObject && !force) {
      return {
        ok: false,
        message: `Cannot rename "${from}" back to "${to}": a sheet by that name exists again.`,
      };
    }
    sheet.name = to;
    await context.sync();
    return { ok: true, message: `Renamed the sheet "${from}" back to "${to}".` };
  });
}

async function undoDefineName(
  name: string,
  expectRefersTo: string,
  force: boolean
): Promise<UndoOutcome> {
  return Excel.run(async (context) => {
    const item = context.workbook.names.getItemOrNullObject(name);
    item.load(["isNullObject", "formula"]);
    await context.sync();

    if (item.isNullObject) {
      return { ok: false, message: `Defined name "${name}" is already gone; nothing to remove.` };
    }
    const current = typeof item.formula === "string" ? item.formula : undefined;
    if (current !== undefined && current !== expectRefersTo && !force) {
      return {
        ok: false,
        message:
          `Defined name "${name}" now points at ${current} rather than the ${expectRefersTo} ` +
          `this change set set, so somebody has edited it. Left alone.`,
      };
    }
    item.delete();
    await context.sync();
    return { ok: true, message: `Removed the defined name "${name}" this change created.` };
  });
}

async function undoCreateTable(name: string): Promise<UndoOutcome> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItemOrNullObject(name);
    table.load("isNullObject");
    await context.sync();
    if (table.isNullObject) {
      return { ok: false, message: `Table "${name}" is already gone; nothing to remove.` };
    }
    // delete() removes the table object; the cells stay, and the cell rollback
    // has already put their prior contents back.
    table.delete();
    await context.sync();
    return { ok: true, message: `Removed the table "${name}" this change created.` };
  });
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
  const structuralEdits = changeSet.edits.filter((edit) => !isCellEdit(edit));
  if (structuralEdits.length > 0 && changeSet.compensation === undefined) {
    unrestorable.push(
      `${structuralEdits.length} structural change(s) have no recorded inverse, so they are ` +
        `not reversed: ${structuralEdits.map((edit) => edit.kind).join(", ")}.`
    );
  }

  // After the cells, so a sheet we created is empty by the time we try to
  // remove it. Guards stay on: a human may have edited since we applied.
  const structural = await undoStructural(changeSet.compensation ?? [], {
    ...(options.force !== undefined ? { force: options.force } : {}),
  });
  unrestorable.push(...structural.unreversed);

  return {
    restoredCells,
    conflicts,
    reversedStructural: structural.reversed,
    unrestorable,
    ok: true,
  };
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

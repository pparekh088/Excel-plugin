/**
 * Compensating operations for structural edits (INV-2, INV-3).
 *
 * Cell edits are reversible because we snapshot the cell. Structural edits are
 * not: there is no "prior state" of a sheet that did not exist. Reversing them
 * needs an explicit inverse — delete the sheet we created, rename the sheet
 * back, restore the defined name we overwrote — planned BEFORE we apply, while
 * the pre-state is still observable.
 *
 * This module is pure. It plans the inverses and states the guard each one
 * must satisfy to run; executing them belongs to whichever host applied the
 * originals (the simulator here, Office.js in the add-in). That split is what
 * lets the same reasoning cover both without the deterministic layer knowing
 * anything about Excel.
 *
 * Every inverse is guarded by the same rule cell rollback uses: reverse it only
 * if we can still prove it is ours. A sheet we created that now holds somebody
 * else's data is not deleted, and we say so rather than pretending the
 * rollback was complete.
 */

import { Workbook, a1Range } from "../model/workbook";
import { Edit, StructuralEdit, isCellEdit } from "./types";

export type CompensatingOp =
  | {
      kind: "deleteSheet";
      sheet: string;
      /**
       * The cells this change set wrote on that sheet. Anything else on it
       * when we come to undo is somebody's work, and the sheet stays.
       */
      ourCells: number[];
    }
  | { kind: "renameSheet"; from: string; to: string }
  | { kind: "deleteName"; name: string; scope: string | null; expectRefersTo: string }
  | {
      kind: "restoreName";
      name: string;
      scope: string | null;
      refersTo: string;
      comment?: string;
      expectRefersTo: string;
    }
  | {
      kind: "deleteTable";
      name: string;
      sheet: string;
      /**
       * The range the table was created over. If the live table now spans a
       * different range, somebody has resized or moved it since — deleting the
       * table object would break their structured references, so it stays.
       */
      expectRange?: string;
    }
  | { kind: "none"; describes: string; reason: string };

/** Packed (row, col) so a sheet's cell set is cheap to compare. */
const COLUMN_STRIDE = 16_384;

function packedCellsOn(sheet: string, edits: Edit[]): number[] {
  const packed = new Set<number>();
  const target = sheet.toUpperCase();
  for (const edit of edits) {
    if (!isCellEdit(edit) || edit.sheet.toUpperCase() !== target) continue;
    packed.add(edit.row * COLUMN_STRIDE + edit.col);
  }
  return [...packed];
}

/**
 * Plan the inverse of every structural edit, against the workbook AS IT IS
 * BEFORE the change set is applied. Order matches `edits`; rollback runs them
 * in reverse so a create-then-rename undoes in the right sequence.
 */
export function planCompensation(workbook: Workbook, edits: Edit[]): CompensatingOp[] {
  const ops: CompensatingOp[] = [];
  // Track what earlier edits in this same change set already created, so a
  // later edit does not read "already exists" from our own work.
  const createdSheets = new Set<string>();
  const createdNames = new Set<string>();
  const createdTables = new Set<string>();

  for (const edit of edits) {
    if (isCellEdit(edit)) continue;
    ops.push(planOne(workbook, edit, edits, { createdSheets, createdNames, createdTables }));
  }
  return ops;
}

interface PlanState {
  createdSheets: Set<string>;
  createdNames: Set<string>;
  createdTables: Set<string>;
}

function nameKey(name: string, scope: string | null): string {
  return `${(scope ?? "").toUpperCase()}::${name.toUpperCase()}`;
}

function planOne(
  workbook: Workbook,
  edit: StructuralEdit,
  allEdits: Edit[],
  state: PlanState
): CompensatingOp {
  switch (edit.kind) {
    case "createSheet": {
      if (!edit.name) {
        return { kind: "none", describes: "createSheet", reason: "No sheet name was given." };
      }
      const key = edit.name.toUpperCase();
      if (workbook.sheet(edit.name) !== undefined && !state.createdSheets.has(key)) {
        // addSheet is idempotent, so this applied as a no-op. Deleting the
        // sheet on rollback would destroy a sheet we did not create.
        return {
          kind: "none",
          describes: `createSheet "${edit.name}"`,
          reason: `Sheet "${edit.name}" already existed, so nothing was created.`,
        };
      }
      state.createdSheets.add(key);
      return { kind: "deleteSheet", sheet: edit.name, ourCells: packedCellsOn(edit.name, allEdits) };
    }

    case "renameSheet": {
      if (!edit.sheet || !edit.newName) {
        return { kind: "none", describes: "renameSheet", reason: "Rename was incompletely specified." };
      }
      // Undo by renaming back — guarded on the sheet still carrying the name
      // we gave it.
      return { kind: "renameSheet", from: edit.newName, to: edit.sheet };
    }

    case "defineName": {
      if (!edit.name || !edit.refersTo) {
        return { kind: "none", describes: "defineName", reason: "Name was incompletely specified." };
      }
      const scope = edit.sheet ?? null;
      const key = nameKey(edit.name, scope);
      const existing = state.createdNames.has(key)
        ? undefined
        : workbook.definedName(edit.name, scope);
      state.createdNames.add(key);

      if (existing === undefined) {
        return {
          kind: "deleteName",
          name: edit.name,
          scope,
          expectRefersTo: edit.refersTo,
        };
      }
      // We overwrote a definition somebody was relying on: put it back.
      return {
        kind: "restoreName",
        name: existing.name,
        scope: existing.scope,
        refersTo: existing.refersTo,
        ...(existing.comment !== undefined ? { comment: existing.comment } : {}),
        expectRefersTo: edit.refersTo,
      };
    }

    case "createTable": {
      if (!edit.name) {
        return { kind: "none", describes: "createTable", reason: "No table name was given." };
      }
      const key = edit.name.toUpperCase();
      if (workbook.table(edit.name) !== undefined && !state.createdTables.has(key)) {
        return {
          kind: "none",
          describes: `createTable "${edit.name}"`,
          reason: `Table "${edit.name}" already existed, so nothing was created.`,
        };
      }
      state.createdTables.add(key);
      return {
        kind: "deleteTable",
        name: edit.name,
        sheet: edit.sheet ?? "",
        ...(edit.range ? { expectRange: edit.range.toUpperCase() } : {}),
      };
    }
  }
}

export interface CompensationResult {
  /** Inverses that ran, in plain language. */
  reversed: string[];
  /** Inverses that did NOT run, and why — feeds the rollback report verbatim. */
  unreversed: string[];
}

/**
 * Execute the inverses against the in-memory model (simulator path).
 *
 * Runs in reverse order. Each guard answers one question: is this still ours
 * to undo? When the answer is no, the change stays and the reason is recorded
 * rather than swallowed.
 *
 * `force` skips the guards. It exists only to serve an explicit user decision
 * made with the reasons in front of them.
 */
export function applyCompensation(
  workbook: Workbook,
  ops: CompensatingOp[],
  options: { force?: boolean } = {}
): CompensationResult {
  const reversed: string[] = [];
  const unreversed: string[] = [];

  for (const op of [...ops].reverse()) {
    switch (op.kind) {
      case "none":
        unreversed.push(`${op.describes}: ${op.reason}`);
        break;

      case "deleteSheet": {
        const sheet = workbook.sheet(op.sheet);
        if (!sheet) {
          unreversed.push(`Sheet "${op.sheet}" is already gone; nothing to delete.`);
          break;
        }
        const ours = new Set(op.ourCells);
        const foreign = [...sheet.cells.keys()].filter((key) => !ours.has(key));
        if (foreign.length > 0 && !options.force) {
          unreversed.push(
            `Sheet "${op.sheet}" was created by this change but now holds ${foreign.length} ` +
              `cell(s) somebody else added, so it was NOT deleted. Delete it by hand if you ` +
              `want it gone.`
          );
          break;
        }
        workbook.removeSheet(op.sheet);
        reversed.push(`Deleted the sheet "${op.sheet}" this change created.`);
        break;
      }

      case "renameSheet": {
        const sheet = workbook.sheet(op.from);
        if (!sheet) {
          unreversed.push(
            `Sheet "${op.from}" no longer exists, so the rename back to "${op.to}" was skipped.`
          );
          break;
        }
        if (workbook.sheet(op.to) !== undefined && !options.force) {
          unreversed.push(
            `Cannot rename "${op.from}" back to "${op.to}": a sheet by that name exists again.`
          );
          break;
        }
        workbook.renameSheet(op.from, op.to);
        reversed.push(`Renamed the sheet "${op.from}" back to "${op.to}".`);
        break;
      }

      case "deleteName": {
        const existing = workbook.definedName(op.name, op.scope);
        if (!existing) {
          unreversed.push(`Defined name "${op.name}" is already gone; nothing to remove.`);
          break;
        }
        if (existing.refersTo !== op.expectRefersTo && !options.force) {
          unreversed.push(
            `Defined name "${op.name}" now points at ${existing.refersTo} rather than the ` +
              `${op.expectRefersTo} this change set set, so somebody has edited it. Left alone.`
          );
          break;
        }
        workbook.removeName(op.name, op.scope);
        reversed.push(`Removed the defined name "${op.name}" this change created.`);
        break;
      }

      case "restoreName": {
        const existing = workbook.definedName(op.name, op.scope);
        if (existing && existing.refersTo !== op.expectRefersTo && !options.force) {
          unreversed.push(
            `Defined name "${op.name}" has been edited since; its prior definition ` +
              `(${op.refersTo}) was NOT restored over that.`
          );
          break;
        }
        workbook.removeName(op.name, op.scope);
        workbook.names.push({
          name: op.name,
          scope: op.scope,
          refersTo: op.refersTo,
          ...(op.comment !== undefined ? { comment: op.comment } : {}),
        });
        reversed.push(`Restored the defined name "${op.name}" to ${op.refersTo}.`);
        break;
      }

      case "deleteTable": {
        const existing = workbook.table(op.name);
        if (!existing) {
          unreversed.push(`Table "${op.name}" is already gone; nothing to remove.`);
          break;
        }
        const currentRange = a1Range(
          existing.startRow,
          existing.startCol,
          existing.endRow,
          existing.endCol
        ).toUpperCase();
        if (op.expectRange !== undefined && currentRange !== op.expectRange && !options.force) {
          unreversed.push(
            `Table "${op.name}" now spans ${currentRange}, not the ${op.expectRange} this ` +
              `change created it over — somebody has resized it. Removing the table would ` +
              `break their structured references, so it was left alone.`
          );
          break;
        }
        workbook.removeTable(op.name);
        reversed.push(`Removed the table "${op.name}" this change created.`);
        break;
      }
    }
  }

  return { reversed, unreversed };
}

/**
 * Change-set engine (INV-2, INV-3, INV-8).
 *
 * Pure and deterministic: it computes snapshots, diffs, impact and rollback
 * plans against the in-memory workbook model. The add-in applies the resulting
 * edits through Office.js; the simulator applies them directly in evals. Both
 * paths use the same change set, so what the user previews is what gets
 * applied.
 */

import { DependencyGraph } from "../graph/graph";
import { Cell, Workbook, a1, fullAddress } from "../model/workbook";
import { applyCompensation, planCompensation } from "./structural";
import {
  CellEdit,
  CellSnapshot,
  ChangeSet,
  RollbackConflict,
  RollbackOptions,
  DiffEntry,
  DriftCheck,
  DriftEntry,
  Edit,
  ImpactSummary,
  RiskTier,
  RollbackReport,
  isCellEdit,
} from "./types";

let counter = 0;

export function newChangeSetId(): string {
  counter += 1;
  return `cs_${Date.now().toString(36)}_${counter.toString(36)}`;
}

/**
 * Risk tiering per handoff §4:
 *   LOW    formatting / number formats
 *   MEDIUM new sheets, formulas written into empty cells
 *   HIGH   overwriting non-empty cells, deleting anything, changing existing formulas
 */
/**
 * Prior state to judge risk against, keyed "SHEET!row,col".
 *
 * Risk is a statement about what an edit DESTROYS, so it has to be measured
 * against the state the user was shown. Once a change set is applied, judging
 * a follow-up edit against the live workbook measures it against our own
 * write — which makes every correction to our own formula look like
 * "overwriting existing logic".
 */
export type PriorState = Map<string, CellSnapshot>;

export function priorStateKey(sheet: string, row: number, col: number): string {
  return `${sheet.toUpperCase()}!${row},${col}`;
}

export function riskOf(edit: Edit, workbook: Workbook, prior?: PriorState): RiskTier {
  if (!isCellEdit(edit)) {
    // Structural: creating things is medium, renaming existing things is high.
    return edit.kind === "renameSheet" ? "high" : "medium";
  }
  if (edit.kind === "setNumberFormat") return "low";
  if (edit.kind === "clear") return "high";

  const snapshot = prior?.get(priorStateKey(edit.sheet, edit.row, edit.col));
  const existing = snapshot
    ? snapshot.absent
      ? undefined
      : {
          row: edit.row,
          col: edit.col,
          value: snapshot.value,
          ...(snapshot.formula !== undefined ? { formula: snapshot.formula } : {}),
        }
    : workbook.sheet(edit.sheet)?.get(edit.row, edit.col);
  const occupied = existing !== undefined && existing.value !== null && existing.value !== "";
  if (existing?.formula !== undefined) return "high"; // changing existing logic
  if (occupied) return "high"; // overwriting data
  return "medium"; // writing into an empty cell
}

export function overallRisk(edits: Edit[], workbook: Workbook, prior?: PriorState): RiskTier {
  let risk: RiskTier = "low";
  for (const edit of edits) {
    const tier = riskOf(edit, workbook, prior);
    if (tier === "high") return "high";
    if (tier === "medium") risk = "medium";
  }
  return risk;
}

/** INV-3: capture prior state of every touched cell before anything is written. */
export function snapshot(workbook: Workbook, edits: Edit[]): CellSnapshot[] {
  const seen = new Set<string>();
  const snapshots: CellSnapshot[] = [];
  for (const edit of edits) {
    if (!isCellEdit(edit)) continue;
    const key = `${edit.sheet.toUpperCase()}!${edit.row},${edit.col}`;
    if (seen.has(key)) continue;
    seen.add(key);

    const cell = workbook.sheet(edit.sheet)?.get(edit.row, edit.col);
    snapshots.push({
      sheet: edit.sheet,
      row: edit.row,
      col: edit.col,
      value: cell?.value ?? null,
      ...(cell?.formula !== undefined ? { formula: cell.formula } : {}),
      ...(cell?.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
      absent: cell === undefined,
    });
  }
  return snapshots;
}

export function buildDiff(workbook: Workbook, edits: Edit[]): DiffEntry[] {
  const entries: DiffEntry[] = [];
  for (const edit of edits) {
    if (!isCellEdit(edit)) continue;
    const cell = workbook.sheet(edit.sheet)?.get(edit.row, edit.col);
    entries.push({
      address: fullAddress(edit.sheet, edit.row, edit.col),
      sheet: edit.sheet,
      row: edit.row,
      col: edit.col,
      before: {
        value: cell?.value ?? null,
        ...(cell?.formula !== undefined ? { formula: cell.formula } : {}),
        ...(cell?.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
        absent: cell === undefined,
      },
      after: {
        ...(edit.kind === "setValue" ? { value: edit.value } : {}),
        ...(edit.kind === "setFormula" ? { formula: edit.formula } : {}),
        ...(edit.kind === "setNumberFormat" ? { numberFormat: edit.numberFormat } : {}),
        cleared: edit.kind === "clear",
      },
      overwritesFormula:
        cell?.formula !== undefined &&
        (edit.kind === "clear" || edit.formula !== cell.formula),
    });
  }
  return entries;
}

/**
 * Impact analysis through the dependency graph. When any downstream path runs
 * through an opaque reference the count is reported as a LOWER BOUND rather
 * than presented as complete (INV-4 honesty).
 */
export function analyzeImpact(
  workbook: Workbook,
  graph: DependencyGraph,
  edits: Edit[]
): ImpactSummary {
  const affectedNodeIds = new Set<number>();
  for (const edit of edits) {
    if (!isCellEdit(edit)) continue;
    for (const id of graph.impactOfArea(edit.sheet, edit.row, edit.col, edit.row, edit.col)) {
      affectedNodeIds.add(id);
    }
  }

  let affectedCells = 0;
  let opaqueDownstream = 0;
  const affectedOutputs: string[] = [];
  for (const id of affectedNodeIds) {
    const node = graph.node(id);
    if (!node) continue;
    affectedCells += node.cellCount;
    if (node.opaque) opaqueDownstream++;
    // A formula nothing else reads is an output worth naming in the preview.
    if (node.kind === "formula" && graph.dependents(node.id).length === 0) {
      affectedOutputs.push(graph.nodeAddress(node));
    }
  }

  const touchedSheets = new Set(
    edits.filter(isCellEdit).map((edit) => edit.sheet.toUpperCase())
  );
  const affectedCharts = workbook.charts
    .filter((chart) => touchedSheets.has(chart.sheet.toUpperCase()))
    .map((chart) => chart.name);
  const affectedPivots = workbook.pivots
    .filter((pivot) => touchedSheets.has(pivot.sheet.toUpperCase()))
    .map((pivot) => pivot.name);

  return {
    affectedCells,
    affectedNodes: affectedNodeIds.size,
    affectedOutputs: affectedOutputs.slice(0, 20),
    affectedCharts,
    affectedPivots,
    opaqueDownstream,
  };
}

export interface ProposeOptions {
  intent: string;
  summary: string;
  edits: Edit[];
  graph?: DependencyGraph;
}

export function proposeChangeSet(workbook: Workbook, options: ProposeOptions): ChangeSet {
  const graph = options.graph ?? DependencyGraph.build(workbook);
  return {
    id: newChangeSetId(),
    createdAt: new Date().toISOString(),
    intent: options.intent,
    summary: options.summary,
    edits: options.edits,
    risk: overallRisk(options.edits, workbook),
    status: "proposed",
    snapshots: snapshot(workbook, options.edits),
    // Planned now, against the pre-apply workbook: once the sheet exists we can
    // no longer tell whether we were the ones who created it.
    compensation: planCompensation(workbook, options.edits),
    diff: buildDiff(workbook, options.edits),
    impact: analyzeImpact(workbook, graph, options.edits),
  };
}

/**
 * INV-8: re-read the touched cells and compare against the snapshot taken at
 * propose time. Any difference means someone else edited the workbook and the
 * change set must not be applied.
 */
export function checkDrift(workbook: Workbook, changeSet: ChangeSet): DriftCheck {
  const entries: DriftEntry[] = [];
  for (const snap of changeSet.snapshots) {
    const cell = workbook.sheet(snap.sheet)?.get(snap.row, snap.col);
    const currentValue = cell?.value ?? null;
    const currentFormula = cell?.formula;
    const wasAbsent = snap.absent;
    const isAbsent = cell === undefined;

    const sameFormula = currentFormula === snap.formula;
    // Values recalculate on their own; only a formula change (or a value
    // change on a constant cell) counts as someone else's edit.
    const valueMatters = snap.formula === undefined;
    const sameValue = !valueMatters || Object.is(currentValue, snap.value);

    if (wasAbsent !== isAbsent || !sameFormula || !sameValue) {
      entries.push({
        address: fullAddress(snap.sheet, snap.row, snap.col),
        expected: { value: snap.value, ...(snap.formula ? { formula: snap.formula } : {}) },
        actual: {
          value: currentValue,
          ...(currentFormula ? { formula: currentFormula } : {}),
        },
      });
    }
  }
  return { clean: entries.length === 0, entries };
}

/**
 * Apply the change set's edits to the in-memory model (simulator path).
 *
 * Records `appliedState` as part of applying, so rollback can always tell our
 * write from a later human edit. Recording it here rather than leaving it to
 * the caller means the two can never drift apart.
 */
export function applyToWorkbook(workbook: Workbook, changeSet: ChangeSet): void {
  for (const edit of changeSet.edits) {
    if (isCellEdit(edit)) {
      applyCellEdit(workbook, edit);
      continue;
    }
    switch (edit.kind) {
      case "createSheet":
        if (edit.name) workbook.addSheet(edit.name);
        break;
      case "defineName":
        if (edit.name && edit.refersTo) {
          // Defining a name that exists REPLACES it, as Excel's names.add
          // does. Pushing a second entry left two definitions of one name,
          // and the lookup would then answer with whichever came first.
          workbook.removeName(edit.name, edit.sheet ?? null);
          workbook.names.push({
            name: edit.name,
            scope: edit.sheet ?? null,
            refersTo: edit.refersTo,
          });
        }
        break;
      default:
        // renameSheet / createTable are applied host-side only.
        break;
    }
  }
  changeSet.appliedState = captureAppliedState(workbook, changeSet);
}

function applyCellEdit(workbook: Workbook, edit: CellEdit): void {
  const sheet = workbook.addSheet(edit.sheet);
  const existing = sheet.get(edit.row, edit.col);

  switch (edit.kind) {
    case "clear":
      sheet.delete(edit.row, edit.col);
      return;
    case "setFormula":
      sheet.set({
        row: edit.row,
        col: edit.col,
        value: existing?.value ?? null,
        formula: edit.formula,
        ...(existing?.numberFormat !== undefined
          ? { numberFormat: existing.numberFormat }
          : {}),
      });
      return;
    case "setValue":
      sheet.set({
        row: edit.row,
        col: edit.col,
        value: edit.value ?? null,
        ...(existing?.numberFormat !== undefined
          ? { numberFormat: existing.numberFormat }
          : {}),
      });
      return;
    case "setNumberFormat":
      sheet.set({
        row: edit.row,
        col: edit.col,
        value: existing?.value ?? null,
        ...(existing?.formula !== undefined ? { formula: existing.formula } : {}),
        numberFormat: edit.numberFormat,
      });
      return;
  }
}

/**
 * Record what the cells hold immediately after applying. Rollback needs this to
 * tell "nobody has touched this since" from "a human has edited it".
 *
 * `applyToWorkbook` calls this itself. Hosts that apply edits out of band —
 * the Office.js writer, which pushes through the real Excel API — must call it
 * on the re-read workbook once their write completes.
 *
 * Safe to call either side of a recalculation: the comparison in
 * `changedSinceApply` looks at formulas for formula cells and values only for
 * constants, and recalculation moves neither.
 */
export function captureAppliedState(workbook: Workbook, changeSet: ChangeSet): CellSnapshot[] {
  return changeSet.snapshots.map((snap) => {
    const cell = workbook.sheet(snap.sheet)?.get(snap.row, snap.col);
    return {
      sheet: snap.sheet,
      row: snap.row,
      col: snap.col,
      value: cell?.value ?? null,
      ...(cell?.formula !== undefined ? { formula: cell.formula } : {}),
      ...(cell?.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
      absent: cell === undefined,
    };
  });
}

/**
 * Has this cell changed since we applied? Uses the same rule as drift
 * detection: a formula cell's VALUE moves on every recalculation, so only its
 * formula counts; a constant cell's value is the thing to watch.
 */
function changedSinceApply(applied: CellSnapshot, current: Cell | undefined): boolean {
  const isAbsent = current === undefined;
  if (applied.absent !== isAbsent) return true;
  if (applied.formula !== current?.formula) return true;
  if (applied.formula === undefined && !Object.is(current?.value ?? null, applied.value)) {
    return true;
  }
  return false;
}

/**
 * Rollback from snapshots. Best-effort by design and honest about it:
 * structural edits and anything with host-side state we did not capture are
 * listed in `unrestorable` rather than silently skipped.
 *
 * Cells a human edited after we applied are NOT restored. Reverting somebody's
 * newer work because our verification failed would be the most damaging thing
 * this system could do, so those come back as conflicts for the user to decide
 * on. `force` overrides, and exists only to serve an explicit user decision
 * made with the conflict list in front of them.
 */
export function rollback(
  workbook: Workbook,
  changeSet: ChangeSet,
  options: RollbackOptions = {}
): RollbackReport {
  const unrestorable: string[] = [];
  const conflicts: RollbackConflict[] = [];
  let restoredCells = 0;

  // Index the post-apply state so each cell can be checked in O(1).
  const appliedByCell = new Map<string, CellSnapshot>();
  for (const state of changeSet.appliedState ?? []) {
    appliedByCell.set(`${state.sheet.toUpperCase()}!${state.row},${state.col}`, state);
  }

  const structuralEdits = changeSet.edits.filter((edit) => !isCellEdit(edit));
  if (structuralEdits.length > 0 && changeSet.compensation === undefined) {
    // A change set that predates compensation planning, or one that came back
    // from storage without it. We cannot invent the pre-state after the fact.
    unrestorable.push(
      `${structuralEdits.length} structural change(s) have no recorded inverse, so they are ` +
        `not reversed: ${structuralEdits.map((edit) => edit.kind).join(", ")}.`
    );
  }

  for (const snap of changeSet.snapshots) {
    const sheet = workbook.sheet(snap.sheet);
    if (!sheet) {
      unrestorable.push(`Sheet "${snap.sheet}" no longer exists; ${a1(snap.row, snap.col)} skipped.`);
      continue;
    }

    const current = sheet.get(snap.row, snap.col);
    const applied = appliedByCell.get(`${snap.sheet.toUpperCase()}!${snap.row},${snap.col}`);

    // No post-apply state recorded means we cannot prove the cell is still
    // ours. Restoring anyway could silently revert somebody's work, so the
    // safe reading of missing evidence is "do not touch it".
    if (!applied) {
      conflicts.push({
        address: fullAddress(snap.sheet, snap.row, snap.col),
        reason:
          "No post-apply state was recorded for this cell, so I cannot tell whether " +
          "someone has edited it since. Left as-is.",
        applied: { value: null },
        current: {
          value: current?.value ?? null,
          ...(current?.formula !== undefined ? { formula: current.formula } : {}),
        },
      });
      continue;
    }

    if (!options.force && changedSinceApply(applied, current)) {
      conflicts.push({
        address: fullAddress(snap.sheet, snap.row, snap.col),
        reason: "Cell was edited after the AI change; your edit was kept.",
        applied: {
          value: applied.value,
          ...(applied.formula !== undefined ? { formula: applied.formula } : {}),
        },
        current: {
          value: current?.value ?? null,
          ...(current?.formula !== undefined ? { formula: current.formula } : {}),
        },
      });
      continue;
    }

    if (snap.absent) {
      sheet.delete(snap.row, snap.col);
      restoredCells++;
      continue;
    }
    sheet.set({
      row: snap.row,
      col: snap.col,
      value: snap.value,
      ...(snap.formula !== undefined ? { formula: snap.formula } : {}),
      ...(snap.numberFormat !== undefined ? { numberFormat: snap.numberFormat } : {}),
    });
    restoredCells++;
  }

  // Structural inverses run AFTER the cells are restored: a sheet we created
  // can only be deleted once we have put back whatever we wrote on it.
  const structural = applyCompensation(workbook, changeSet.compensation ?? [], {
    ...(options.force !== undefined ? { force: options.force } : {}),
  });
  unrestorable.push(...structural.unreversed);

  if (workbook.pivots.length > 0) {
    unrestorable.push(
      `${workbook.pivots.length} pivot table(s) present: their caches are not restored by ` +
        `rollback and may need a manual refresh.`
    );
  }

  return {
    changeSetId: changeSet.id,
    restoredCells,
    reversedStructural: structural.reversed,
    unrestorable,
    conflicts,
    // "ok" means the rollback completed as designed — conflicts are a correct
    // outcome, not a failure. The caller reports them to the user.
    ok: true,
  };
}

/** Plain-language explanation written to chat and _AI_Log (INV-10). */
export function explainChangeSet(changeSet: ChangeSet): string {
  const lines: string[] = [];
  lines.push(`${changeSet.summary}`);
  lines.push("");
  lines.push(`Intent: ${changeSet.intent}`);
  lines.push(`Risk: ${changeSet.risk}`);

  const formulaEdits = changeSet.diff.filter((entry) => entry.after.formula !== undefined);
  const valueEdits = changeSet.diff.filter((entry) => entry.after.value !== undefined);
  const clears = changeSet.diff.filter((entry) => entry.after.cleared);
  const overwrites = changeSet.diff.filter((entry) => entry.overwritesFormula);

  lines.push("");
  lines.push("What changed:");
  if (formulaEdits.length > 0) {
    lines.push(`  ${formulaEdits.length} formula(s) set`);
    for (const entry of formulaEdits.slice(0, 8)) {
      lines.push(
        `    ${entry.address}: ${entry.before.formula ?? String(entry.before.value ?? "(empty)")}` +
          ` -> ${entry.after.formula}`
      );
    }
    if (formulaEdits.length > 8) lines.push(`    ... and ${formulaEdits.length - 8} more`);
  }
  if (valueEdits.length > 0) lines.push(`  ${valueEdits.length} value(s) written`);
  if (clears.length > 0) lines.push(`  ${clears.length} cell(s) cleared`);
  if (overwrites.length > 0) {
    lines.push(`  ${overwrites.length} existing formula(s) replaced — these are the risky ones`);
  }

  lines.push("");
  lines.push("What it affects downstream:");
  lines.push(
    `  ${changeSet.impact.affectedCells} cell(s) across ${changeSet.impact.affectedNodes} block(s) recalculate`
  );
  if (changeSet.impact.affectedOutputs.length > 0) {
    lines.push(`  outputs touched: ${changeSet.impact.affectedOutputs.slice(0, 6).join(", ")}`);
  }
  if (changeSet.impact.affectedCharts.length > 0) {
    lines.push(`  charts on touched sheets: ${changeSet.impact.affectedCharts.join(", ")}`);
  }
  if (changeSet.impact.affectedPivots.length > 0) {
    lines.push(
      `  pivots on touched sheets: ${changeSet.impact.affectedPivots.join(", ")} ` +
        `(may need manual refresh)`
    );
  }
  if (changeSet.impact.opaqueDownstream > 0) {
    lines.push(
      `  NOTE: ${changeSet.impact.opaqueDownstream} downstream block(s) use INDIRECT/OFFSET or ` +
        `external links, so the impact figures above are a LOWER BOUND.`
    );
  }
  return lines.join("\n");
}

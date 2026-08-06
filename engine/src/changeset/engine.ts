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
import { Workbook, a1, fullAddress } from "../model/workbook";
import {
  CellEdit,
  CellSnapshot,
  ChangeSet,
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
export function riskOf(edit: Edit, workbook: Workbook): RiskTier {
  if (!isCellEdit(edit)) {
    // Structural: creating things is medium, renaming existing things is high.
    return edit.kind === "renameSheet" ? "high" : "medium";
  }
  if (edit.kind === "setNumberFormat") return "low";
  if (edit.kind === "clear") return "high";

  const existing = workbook.sheet(edit.sheet)?.get(edit.row, edit.col);
  const occupied = existing !== undefined && existing.value !== null && existing.value !== "";
  if (existing?.formula !== undefined) return "high"; // changing existing logic
  if (occupied) return "high"; // overwriting data
  return "medium"; // writing into an empty cell
}

export function overallRisk(edits: Edit[], workbook: Workbook): RiskTier {
  let risk: RiskTier = "low";
  for (const edit of edits) {
    const tier = riskOf(edit, workbook);
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

/** Apply the change set's edits to the in-memory model (simulator path). */
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
          workbook.names.push({ name: edit.name, scope: null, refersTo: edit.refersTo });
        }
        break;
      default:
        // renameSheet / createTable are applied host-side only.
        break;
    }
  }
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
 * Rollback from snapshots. Best-effort by design and honest about it:
 * structural edits and anything with host-side state we did not capture are
 * listed in `unrestorable` rather than silently skipped.
 */
export function rollback(workbook: Workbook, changeSet: ChangeSet): RollbackReport {
  const unrestorable: string[] = [];
  const skippedDueToDrift: string[] = [];
  let restoredCells = 0;

  for (const edit of changeSet.edits) {
    if (isCellEdit(edit)) continue;
    switch (edit.kind) {
      case "createSheet":
        unrestorable.push(
          `Sheet "${edit.name}" was created; removing it is not part of rollback ` +
            `(delete it manually if unwanted).`
        );
        break;
      case "renameSheet":
        unrestorable.push(`Sheet rename ${edit.name} -> ${edit.newName} is not reversed.`);
        break;
      case "createTable":
        unrestorable.push(`Table "${edit.name}" was created and is not removed by rollback.`);
        break;
      case "defineName":
        unrestorable.push(`Defined name "${edit.name}" is not removed by rollback.`);
        break;
    }
  }

  for (const snap of changeSet.snapshots) {
    const sheet = workbook.sheet(snap.sheet);
    if (!sheet) {
      unrestorable.push(`Sheet "${snap.sheet}" no longer exists; ${a1(snap.row, snap.col)} skipped.`);
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

  if (workbook.pivots.length > 0) {
    unrestorable.push(
      `${workbook.pivots.length} pivot table(s) present: their caches are not restored by ` +
        `rollback and may need a manual refresh.`
    );
  }

  return {
    changeSetId: changeSet.id,
    restoredCells,
    unrestorable,
    skippedDueToDrift,
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

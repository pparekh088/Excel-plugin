/**
 * Pre-flight hazard checks (handoff §10 Phase 5).
 *
 * Things that make a write fail, or worse, silently do the wrong thing:
 * protected sheets, merged cells, spill ranges, tables with calculated
 * columns, and cells inside an array formula.
 *
 * The rule throughout: never fail silently. A hazard either blocks the change
 * set with a message a user can act on, or warns and proceeds — and which of
 * those it is depends on whether proceeding could damage something.
 */

import { Workbook, a1, fullAddress } from "../model/workbook";
import { Edit, isCellEdit } from "./types";

export type HazardSeverity = "blocking" | "warning";

export interface Hazard {
  kind:
    | "protected-sheet"
    | "merged-cell"
    | "spill-range"
    | "table-calculated-column"
    | "missing-sheet";
  severity: HazardSeverity;
  address?: string;
  sheet: string;
  /** What the user should do about it. */
  message: string;
}

export interface HazardReport {
  hazards: Hazard[];
  /** True when at least one hazard blocks the change set. */
  blocked: boolean;
}

/** Is (row, col) inside any merged area on this sheet? */
function mergedAreaAt(
  workbook: Workbook,
  sheetName: string,
  row: number,
  col: number
): [number, number, number, number] | null {
  const sheet = workbook.sheet(sheetName);
  if (!sheet) return null;
  for (const area of sheet.merged) {
    const [startRow, startCol, endRow, endCol] = area;
    if (row >= startRow && row <= endRow && col >= startCol && col <= endCol) return area;
  }
  return null;
}

export function checkHazards(workbook: Workbook, edits: Edit[]): HazardReport {
  const hazards: Hazard[] = [];
  const seenSheets = new Set<string>();

  for (const edit of edits) {
    if (!isCellEdit(edit)) {
      continue;
    }
    const sheet = workbook.sheet(edit.sheet);
    if (!sheet) {
      if (!seenSheets.has(edit.sheet)) {
        seenSheets.add(edit.sheet);
        hazards.push({
          kind: "missing-sheet",
          severity: "blocking",
          sheet: edit.sheet,
          message:
            `Sheet "${edit.sheet}" does not exist. It may have been renamed or deleted since ` +
            `this plan was made.`,
        });
      }
      continue;
    }

    // Protected sheets: writing throws, so ask rather than fail mid-apply.
    if (sheet.protectedSheet && !seenSheets.has(`prot:${sheet.name}`)) {
      seenSheets.add(`prot:${sheet.name}`);
      hazards.push({
        kind: "protected-sheet",
        severity: "blocking",
        sheet: sheet.name,
        message:
          `Sheet "${sheet.name}" is protected, so nothing can be written to it. Unprotect it ` +
          `and run this again — I will not attempt to remove the protection myself.`,
      });
    }

    // Merged cells: only the top-left cell of a merge is writable. Writing to
    // any other cell in the area silently does nothing in some hosts.
    const merged = mergedAreaAt(workbook, sheet.name, edit.row, edit.col);
    if (merged) {
      const [startRow, startCol] = merged;
      const isAnchor = edit.row === startRow && edit.col === startCol;
      if (!isAnchor) {
        hazards.push({
          kind: "merged-cell",
          severity: "blocking",
          sheet: sheet.name,
          address: fullAddress(sheet.name, edit.row, edit.col),
          message:
            `${a1(edit.row, edit.col)} is inside a merged area anchored at ` +
            `${a1(startRow, startCol)}. Only the anchor cell is writable; writing here would ` +
            `be silently ignored.`,
        });
      }
    }

    // Writing into a table's calculated column fights the table's own fill.
    for (const table of workbook.tables) {
      if (table.sheet.toUpperCase() !== sheet.name.toUpperCase()) continue;
      const inBody =
        edit.row > table.headerRow &&
        edit.row <= table.endRow &&
        edit.col >= table.startCol &&
        edit.col <= table.endCol;
      if (!inBody) continue;
      const columnHasFormulas = columnIsCalculated(workbook, table, edit.col);
      if (columnHasFormulas && edit.kind !== "setNumberFormat") {
        hazards.push({
          kind: "table-calculated-column",
          severity: "warning",
          sheet: sheet.name,
          address: fullAddress(sheet.name, edit.row, edit.col),
          message:
            `${a1(edit.row, edit.col)} is in a calculated column of table "${table.name}". ` +
            `Excel may propagate this edit to the whole column, or revert it — check the ` +
            `result after applying.`,
        });
      }
    }
  }

  return {
    hazards,
    blocked: hazards.some((hazard) => hazard.severity === "blocking"),
  };
}

function columnIsCalculated(
  workbook: Workbook,
  table: { sheet: string; headerRow: number; endRow: number },
  col: number
): boolean {
  const sheet = workbook.sheet(table.sheet);
  if (!sheet) return false;
  let formulaCells = 0;
  let bodyCells = 0;
  for (let row = table.headerRow + 1; row <= table.endRow; row++) {
    const cell = sheet.get(row, col);
    if (!cell) continue;
    bodyCells++;
    if (cell.formula !== undefined) formulaCells++;
  }
  // A calculated column is one where essentially every body cell is a formula.
  return bodyCells >= 2 && formulaCells >= bodyCells - 1;
}

/** Human-readable summary for the change-set preview. */
export function describeHazards(report: HazardReport): string {
  if (report.hazards.length === 0) return "";
  const blocking = report.hazards.filter((hazard) => hazard.severity === "blocking");
  const warnings = report.hazards.filter((hazard) => hazard.severity === "warning");

  const lines: string[] = [];
  if (blocking.length > 0) {
    lines.push(`Cannot apply — ${blocking.length} blocking issue(s):`);
    for (const hazard of blocking.slice(0, 8)) lines.push(`  · ${hazard.message}`);
    if (blocking.length > 8) lines.push(`  ... and ${blocking.length - 8} more`);
  }
  if (warnings.length > 0) {
    lines.push(`${warnings.length} warning(s):`);
    for (const hazard of warnings.slice(0, 8)) lines.push(`  · ${hazard.message}`);
    if (warnings.length > 8) lines.push(`  ... and ${warnings.length - 8} more`);
  }
  return lines.join("\n");
}

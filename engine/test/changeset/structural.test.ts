/**
 * Structural atomicity (P0-3).
 *
 * A change set that creates a sheet and writes to it is one transaction. The
 * cells were always reversible; the sheet was not, so "rollback" left a
 * half-undone workbook behind while reporting success. These tests pin the
 * inverse of every structural edit, and — more importantly — pin when the
 * inverse must REFUSE to run because the thing it would destroy is no longer
 * ours.
 */

import { describe, expect, it } from "vitest";
import {
  applyToWorkbook,
  proposeChangeSet,
  rollback,
} from "../../src/changeset/engine";
import { applyCompensation, planCompensation } from "../../src/changeset/structural";
import { Edit } from "../../src/changeset/types";
import { Workbook } from "../../src/model/workbook";
import { workbookOf } from "../helpers/build";

function propose(workbook: Workbook, edits: Edit[]) {
  return proposeChangeSet(workbook, { intent: "x", summary: "x", edits });
}

describe("planCompensation", () => {
  it("plans a delete for a sheet that did not exist", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const ops = planCompensation(workbook, [{ kind: "createSheet", name: "New" }]);
    expect(ops).toEqual([{ kind: "deleteSheet", sheet: "New", ourCells: [] }]);
  });

  it("plans NOTHING for a sheet that already existed", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const ops = planCompensation(workbook, [{ kind: "createSheet", name: "S" }]);
    expect(ops[0]!.kind).toBe("none");
    expect(ops[0]).toMatchObject({ reason: expect.stringContaining("already existed") });
  });

  it("carries the cells the change set writes on a new sheet", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const ops = planCompensation(workbook, [
      { kind: "createSheet", name: "New" },
      { kind: "setValue", sheet: "New", row: 0, col: 0, value: 1 },
      { kind: "setValue", sheet: "New", row: 3, col: 2, value: 2 },
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 9 },
    ]);
    expect(ops).toHaveLength(1);
    expect(ops[0]).toMatchObject({ kind: "deleteSheet", sheet: "New" });
    // Two cells on the new sheet; the one on S does not belong to it.
    expect((ops[0] as { ourCells: number[] }).ourCells).toHaveLength(2);
  });

  it("plans a delete for a brand-new defined name", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const ops = planCompensation(workbook, [
      { kind: "defineName", name: "Revenue", refersTo: "=S!$A$1" },
    ]);
    expect(ops[0]).toEqual({
      kind: "deleteName",
      name: "Revenue",
      scope: null,
      expectRefersTo: "=S!$A$1",
    });
  });

  it("plans a RESTORE when the defined name already existed", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.names.push({
      name: "Revenue",
      scope: null,
      refersTo: "=S!$B$9",
      comment: "the real one",
    });
    const ops = planCompensation(workbook, [
      { kind: "defineName", name: "Revenue", refersTo: "=S!$A$1" },
    ]);
    expect(ops[0]).toEqual({
      kind: "restoreName",
      name: "Revenue",
      scope: null,
      refersTo: "=S!$B$9",
      comment: "the real one",
      expectRefersTo: "=S!$A$1",
    });
  });

  it("does not read its own earlier creation as pre-existing", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    // Two edits creating the same sheet: the first creates, the second is a
    // no-op, and neither may be planned as "the sheet was already there".
    const ops = planCompensation(workbook, [
      { kind: "createSheet", name: "New" },
      { kind: "createSheet", name: "New" },
    ]);
    expect(ops[0]!.kind).toBe("deleteSheet");
    expect(ops[1]!.kind).toBe("deleteSheet");
  });

  it("plans the inverse rename", () => {
    const workbook = workbookOf({ Old: { A1: 1 } });
    const ops = planCompensation(workbook, [
      { kind: "renameSheet", sheet: "Old", newName: "New" },
    ]);
    expect(ops[0]).toEqual({ kind: "renameSheet", from: "New", to: "Old" });
  });
});

describe("rollback reverses structural edits", () => {
  it("deletes a sheet the change set created", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [
      { kind: "createSheet", name: "Summary" },
      { kind: "setFormula", sheet: "Summary", row: 0, col: 0, formula: "=S!A1" },
    ]);
    applyToWorkbook(workbook, changeSet);
    expect(workbook.sheet("Summary")).toBeDefined();

    const report = rollback(workbook, changeSet);
    expect(workbook.sheet("Summary")).toBeUndefined();
    expect(report.reversedStructural.join(" ")).toContain("Summary");
    expect(report.unrestorable).toHaveLength(0);
  });

  it("removes a defined name the change set created", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [
      { kind: "defineName", name: "Rev", refersTo: "=S!$A$1" },
    ]);
    applyToWorkbook(workbook, changeSet);
    expect(workbook.definedName("Rev")).toBeDefined();

    rollback(workbook, changeSet);
    expect(workbook.definedName("Rev")).toBeUndefined();
  });

  it("restores the prior definition of a name it overwrote", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    workbook.names.push({ name: "Rev", scope: null, refersTo: "=S!$Z$99" });
    const changeSet = propose(workbook, [
      { kind: "defineName", name: "Rev", refersTo: "=S!$A$1" },
    ]);
    applyToWorkbook(workbook, changeSet);

    rollback(workbook, changeSet);
    expect(workbook.definedName("Rev")?.refersTo).toBe("=S!$Z$99");
    // Exactly one definition, not two.
    expect(workbook.names.filter((name) => name.name === "Rev")).toHaveLength(1);
  });

  it("leaves a pre-existing sheet alone", () => {
    const workbook = workbookOf({ S: { A1: 1 }, Summary: { A1: "important" } });
    const changeSet = propose(workbook, [
      { kind: "createSheet", name: "Summary" },
      { kind: "setValue", sheet: "Summary", row: 5, col: 0, value: 1 },
    ]);
    applyToWorkbook(workbook, changeSet);

    const report = rollback(workbook, changeSet);
    expect(workbook.sheet("Summary")).toBeDefined();
    expect(workbook.sheet("Summary")!.get(0, 0)!.value).toBe("important");
    expect(report.unrestorable.join(" ")).toContain("already existed");
  });

  it("does NOT delete a created sheet somebody has since put data on", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [
      { kind: "createSheet", name: "Scratch" },
      { kind: "setValue", sheet: "Scratch", row: 0, col: 0, value: 1 },
    ]);
    applyToWorkbook(workbook, changeSet);
    // A colleague adds their own notes to the sheet we made.
    workbook.sheet("Scratch")!.set({ row: 20, col: 0, value: "my notes" });

    const report = rollback(workbook, changeSet);
    expect(workbook.sheet("Scratch")).toBeDefined();
    expect(workbook.sheet("Scratch")!.get(20, 0)!.value).toBe("my notes");
    expect(report.unrestorable.join(" ")).toContain("NOT deleted");
    // Our own cell on it was still reverted.
    expect(workbook.sheet("Scratch")!.get(0, 0)).toBeUndefined();
  });

  it("deletes it anyway when the user explicitly forces the rollback", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [{ kind: "createSheet", name: "Scratch" }]);
    applyToWorkbook(workbook, changeSet);
    workbook.sheet("Scratch")!.set({ row: 20, col: 0, value: "my notes" });

    rollback(workbook, changeSet, { force: true });
    expect(workbook.sheet("Scratch")).toBeUndefined();
  });

  it("does NOT remove a defined name somebody has since repointed", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [
      { kind: "defineName", name: "Rev", refersTo: "=S!$A$1" },
    ]);
    applyToWorkbook(workbook, changeSet);
    workbook.definedName("Rev")!.refersTo = "=S!$C$3";

    const report = rollback(workbook, changeSet);
    expect(workbook.definedName("Rev")?.refersTo).toBe("=S!$C$3");
    expect(report.unrestorable.join(" ")).toContain("somebody has edited it");
  });

  it("says so plainly when a change set has no recorded inverse", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [{ kind: "createSheet", name: "New" }]);
    applyToWorkbook(workbook, changeSet);
    // A change set from before compensation planning existed, or one that came
    // back from storage without it.
    delete changeSet.compensation;

    const report = rollback(workbook, changeSet);
    expect(workbook.sheet("New")).toBeDefined();
    expect(report.unrestorable.join(" ")).toContain("no recorded inverse");
  });

  it("undoes a create-then-rename in the right order", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [
      { kind: "createSheet", name: "Draft" },
      { kind: "renameSheet", sheet: "Draft", newName: "Final" },
    ]);
    applyToWorkbook(workbook, changeSet);
    // The simulator does not apply renames (host-side only), so emulate it.
    workbook.renameSheet("Draft", "Final");

    rollback(workbook, changeSet);
    // Renamed back to Draft first, then deleted — neither sheet survives.
    expect(workbook.sheet("Final")).toBeUndefined();
    expect(workbook.sheet("Draft")).toBeUndefined();
  });

  it("restores the cells before removing the sheet that held them", () => {
    const workbook = workbookOf({ S: { A1: 1 } });
    const changeSet = propose(workbook, [
      { kind: "createSheet", name: "New" },
      { kind: "setValue", sheet: "New", row: 0, col: 0, value: 5 },
      { kind: "setValue", sheet: "S", row: 0, col: 0, value: 42 },
    ]);
    applyToWorkbook(workbook, changeSet);

    const report = rollback(workbook, changeSet);
    expect(report.restoredCells).toBe(2);
    expect(workbook.sheet("S")!.get(0, 0)!.value).toBe(1);
    expect(workbook.sheet("New")).toBeUndefined();
  });
});

describe("applyCompensation on its own", () => {
  it("runs the inverses in reverse order", () => {
    const workbook = new Workbook("w");
    workbook.addSheet("A");
    workbook.addSheet("B");
    const result = applyCompensation(workbook, [
      { kind: "deleteSheet", sheet: "A", ourCells: [] },
      { kind: "deleteSheet", sheet: "B", ourCells: [] },
    ]);
    expect(result.reversed).toEqual([
      'Deleted the sheet "B" this change created.',
      'Deleted the sheet "A" this change created.',
    ]);
  });

  it("keeps going after one inverse refuses", () => {
    const workbook = new Workbook("w");
    workbook.addSheet("A").set({ row: 0, col: 0, value: "theirs" });
    workbook.addSheet("B");
    const result = applyCompensation(workbook, [
      { kind: "deleteSheet", sheet: "A", ourCells: [] },
      { kind: "deleteSheet", sheet: "B", ourCells: [] },
    ]);
    expect(workbook.sheet("A")).toBeDefined();
    expect(workbook.sheet("B")).toBeUndefined();
    expect(result.reversed).toHaveLength(1);
    expect(result.unreversed).toHaveLength(1);
  });
});

describe("Workbook structural primitives", () => {
  it("removeSheet drops names, tables, charts and pivots scoped to it", () => {
    const workbook = new Workbook("w");
    workbook.addSheet("Keep");
    workbook.addSheet("Drop");
    workbook.names.push({ name: "Local", scope: "Drop", refersTo: "=Drop!$A$1" });
    workbook.names.push({ name: "Global", scope: null, refersTo: "=Keep!$A$1" });
    workbook.tables.push({
      name: "T",
      sheet: "Drop",
      headerRow: 0,
      startRow: 0,
      endRow: 1,
      startCol: 0,
      endCol: 1,
      columns: [],
      hasTotals: false,
    });
    workbook.charts.push({ name: "C", sheet: "Drop", sourceRanges: [] });
    workbook.pivots.push({ name: "P", sheet: "Drop", sourceRange: "", olap: false });

    expect(workbook.removeSheet("Drop")).toBe(true);
    expect(workbook.sheet("Drop")).toBeUndefined();
    expect(workbook.names.map((name) => name.name)).toEqual(["Global"]);
    expect(workbook.tables).toHaveLength(0);
    expect(workbook.charts).toHaveLength(0);
    expect(workbook.pivots).toHaveLength(0);
  });

  it("renameSheet refuses to collide with an existing sheet", () => {
    const workbook = new Workbook("w");
    workbook.addSheet("A");
    workbook.addSheet("B");
    expect(workbook.renameSheet("A", "B")).toBe(false);
    expect(workbook.sheet("A")).toBeDefined();
  });

  it("renameSheet follows sheet-scoped names and tables", () => {
    const workbook = new Workbook("w");
    workbook.addSheet("Old");
    workbook.names.push({ name: "Local", scope: "Old", refersTo: "=Old!$A$1" });
    expect(workbook.renameSheet("Old", "New")).toBe(true);
    expect(workbook.sheet("New")).toBeDefined();
    expect(workbook.names[0]!.scope).toBe("New");
  });
});

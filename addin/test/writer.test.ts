/**
 * Writer tests with a minimal Office.js fake.
 *
 * Office.js cannot run headlessly, so these exercise the writer's ORDERING
 * and failure handling — the parts that decide whether a workbook survives a
 * bad apply — against a stub that records the calls it receives. What this
 * cannot verify (real payload limits, real calc suspension semantics) is on
 * the sideload checklist.
 */

import { beforeEach, describe, expect, it, vi } from "vitest";
import { proposeChangeSet, Workbook } from "ledger-engine";

interface RecordedWrite {
  sheet: string;
  row: number;
  col: number;
  kind: "formulas" | "values" | "numberFormat" | "clear";
  payload?: unknown;
}

interface FakeState {
  writes: RecordedWrite[];
  syncs: number;
  suspensions: number;
  calculations: number;
  /** Values the fake reports when a range is read (drift simulation). */
  reads: Map<string, { value: unknown; formula?: string }>;
  /** Throw on the Nth write, to exercise the partial-failure path. */
  throwOnWrite?: number;
}

const state: FakeState = {
  writes: [],
  syncs: 0,
  suspensions: 0,
  calculations: 0,
  reads: new Map(),
};

function makeRange(sheet: string, row: number, col: number) {
  const key = `${sheet}!${row},${col}`;
  const range: Record<string, unknown> = {
    values: [[null]],
    formulas: [[""]],
    numberFormat: [["General"]],
    format: { fill: { color: "", clear: vi.fn() }, font: { bold: false } },
    load: vi.fn(() => {
      const stored = state.reads.get(key);
      range.values = [[stored?.value ?? null]];
      range.formulas = [[stored?.formula ?? stored?.value ?? ""]];
    }),
    untrack: vi.fn(),
    clear: vi.fn(() => {
      state.writes.push({ sheet, row, col, kind: "clear" });
      state.reads.delete(key);
    }),
    select: vi.fn(),
  };

  return new Proxy(range, {
    set(target, prop, value) {
      if (state.throwOnWrite !== undefined && state.writes.length >= state.throwOnWrite) {
        // Fire once: the restore path must be allowed to run.
        state.throwOnWrite = undefined;
        throw new Error("simulated Office.js failure");
      }
      if (prop === "formulas" || prop === "values" || prop === "numberFormat") {
        state.writes.push({
          sheet,
          row,
          col,
          kind: prop as RecordedWrite["kind"],
          payload: value,
        });
        // A write is visible to the next read, as it would be in Excel.
        const written = (value as unknown[][])[0]?.[0];
        if (prop === "values") {
          state.reads.set(key, { value: written });
        } else if (prop === "formulas") {
          state.reads.set(
            key,
            typeof written === "string" && written.startsWith("=")
              ? { value: state.reads.get(key)?.value ?? null, formula: written }
              : { value: written }
          );
        }
      }
      target[prop as string] = value;
      return true;
    },
  });
}

function makeSheet(name: string) {
  return {
    name,
    getRangeByIndexes: (row: number, col: number) => makeRange(name, row, col),
    getUsedRangeOrNullObject: () => ({
      isNullObject: true,
      load: vi.fn(),
      rowIndex: 0,
      rowCount: 0,
    }),
    activate: vi.fn(),
    protection: { protected: false, load: vi.fn() },
  };
}

function installOfficeFake(): void {
  const worksheets = {
    getItem: (name: string) => makeSheet(name),
    getItemOrNullObject: (name: string) => ({ ...makeSheet(name), isNullObject: true, load: vi.fn() }),
    add: vi.fn((name: string) => makeSheet(name)),
    load: vi.fn(),
    items: [],
  };

  const context = {
    workbook: {
      worksheets,
      names: { add: vi.fn() },
      tables: { add: vi.fn(() => ({ name: "" })) },
      application: {
        calculate: vi.fn(() => {
          state.calculations++;
        }),
      },
    },
    application: {
      suspendApiCalculationUntilNextSync: vi.fn(() => {
        state.suspensions++;
      }),
    },
    sync: vi.fn(async () => {
      state.syncs++;
    }),
  };

  (globalThis as Record<string, unknown>).Excel = {
    run: async (callback: (context: unknown) => Promise<unknown>) => callback(context),
    ClearApplyTo: { contents: "contents" },
    CalculationType: { full: "full" },
    SheetVisibility: { visible: "Visible" },
  };
}

beforeEach(() => {
  state.writes = [];
  state.syncs = 0;
  state.suspensions = 0;
  state.calculations = 0;
  state.reads = new Map();
  state.throwOnWrite = undefined;
  installOfficeFake();
});

function buildWorkbook(): Workbook {
  const workbook = new Workbook("test");
  const sheet = workbook.addSheet("S");
  sheet.set({ row: 0, col: 0, value: 5 });
  sheet.set({ row: 0, col: 1, value: 10, formula: "=A1*2" });
  return workbook;
}

describe("applyChangeSet", () => {
  it("writes formulas, suspends calculation, and recalculates once at the end", async () => {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = buildWorkbook();
    // The live workbook matches the snapshot, so no drift.
    state.reads.set("S!0,1", { value: 10, formula: "=A1*2" });

    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });

    const result = await applyChangeSet(changeSet);
    expect(result.ok).toBe(true);
    expect(result.cellsWritten).toBe(1);
    expect(state.writes).toContainEqual(
      expect.objectContaining({ sheet: "S", row: 0, col: 1, kind: "formulas" })
    );
    expect(state.suspensions).toBeGreaterThanOrEqual(1);
    expect(state.calculations).toBe(1);
  });

  it("refuses to write anything when the workbook drifted (INV-8)", async () => {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = buildWorkbook();
    // Someone else changed the formula since the snapshot.
    state.reads.set("S!0,1", { value: 99, formula: "=A1*999" });

    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });

    const result = await applyChangeSet(changeSet);
    expect(result.ok).toBe(false);
    expect(result.cellsWritten).toBe(0);
    expect(state.writes).toHaveLength(0);
    expect(result.failure).toContain("Someone else may be editing");
  });

  it("does not treat a recalculated value as drift", async () => {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = buildWorkbook();
    // Same formula, different value — a recalculation, not an edit.
    state.reads.set("S!0,1", { value: 12345, formula: "=A1*2" });

    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    expect((await applyChangeSet(changeSet)).ok).toBe(true);
  });

  it("batches large change sets across syncs (INV-5)", async () => {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = new Workbook("big");
    workbook.addSheet("S");
    const edits = Array.from({ length: 1200 }, (_, index) => ({
      kind: "setValue" as const,
      sheet: "S",
      row: index,
      col: 0,
      value: index,
    }));
    const changeSet = proposeChangeSet(workbook, {
      intent: "bulk",
      summary: "bulk",
      edits,
    });

    const result = await applyChangeSet(changeSet, { batchSize: 500 });
    expect(result.ok).toBe(true);
    expect(result.cellsWritten).toBe(1200);
    // 3 write batches, each re-arming suspension (Q-006).
    expect(state.suspensions).toBe(3);
  });

  it("reports progress as it writes", async () => {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = new Workbook("big");
    workbook.addSheet("S");
    const edits = Array.from({ length: 30 }, (_, index) => ({
      kind: "setValue" as const,
      sheet: "S",
      row: index,
      col: 0,
      value: index,
    }));
    const changeSet = proposeChangeSet(workbook, { intent: "x", summary: "x", edits });

    const progress: number[] = [];
    await applyChangeSet(changeSet, {
      batchSize: 10,
      onProgress: (update) => progress.push(update.written),
    });
    expect(progress).toEqual([10, 20, 30]);
  });

  it("rolls back from the snapshot when a write fails partway", async () => {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = new Workbook("S");
    const sheet = workbook.addSheet("S");
    for (let row = 0; row < 6; row++) {
      sheet.set({ row, col: 0, value: row, formula: `=${row}+1` });
      state.reads.set(`S!${row},0`, { value: row, formula: `=${row}+1` });
    }

    const changeSet = proposeChangeSet(workbook, {
      intent: "test",
      summary: "test",
      edits: Array.from({ length: 6 }, (_, row) => ({
        kind: "setFormula" as const,
        sheet: "S",
        row,
        col: 0,
        formula: "=999",
      })),
    });

    state.throwOnWrite = 3; // fail after three writes land
    const result = await applyChangeSet(changeSet);

    expect(result.ok).toBe(false);
    expect(result.rolledBack).toBe(true);
    expect(result.failure).toContain("Restored from snapshot");
    // The restore wrote the original formulas back.
    const restored = state.writes.filter((write) => write.payload && String(JSON.stringify(write.payload)).includes("+1"));
    expect(restored.length).toBeGreaterThan(0);
  });
});

describe("live rollback does not destroy human edits (P0-2)", () => {
  /** before = 10, we apply 20, a human changes it to 30. */
  async function applyThenHumanEdits(humanValue: unknown) {
    const { applyChangeSet } = await import("../src/excel/writer");
    const workbook = new Workbook("test");
    workbook.addSheet("S").set({ row: 26, col: 5, value: 10 });
    state.reads.set("S!26,5", { value: 10 });

    const changeSet = proposeChangeSet(workbook, {
      intent: "bump F27",
      summary: "bump F27",
      edits: [{ kind: "setValue", sheet: "S", row: 26, col: 5, value: 20 }],
    });
    // The fake makes writes visible to the next read, so the post-apply
    // re-read sees the 20 we just wrote.
    const applied = await applyChangeSet(changeSet);
    expect(applied.ok).toBe(true);
    expect(changeSet.appliedState?.[0]?.value).toBe(20);

    // Now a human edits the same cell.
    state.reads.set("S!26,5", { value: humanValue });
    state.writes = [];
    return changeSet;
  }

  it("leaves the human's value alone and reports a conflict", async () => {
    const { rollbackChangeSet } = await import("../src/excel/writer");
    const changeSet = await applyThenHumanEdits(30);

    const report = await rollbackChangeSet(changeSet);
    expect(report.restoredCells).toBe(0);
    expect(state.writes).toHaveLength(0); // nothing written to the workbook
    expect(report.conflicts).toHaveLength(1);
    expect(report.conflicts[0]?.applied.value).toBe(20);
    expect(report.conflicts[0]?.current.value).toBe(30);
    expect(report.ok).toBe(true);
  });

  it("restores when the cell still holds what we wrote", async () => {
    const { rollbackChangeSet } = await import("../src/excel/writer");
    const changeSet = await applyThenHumanEdits(20); // unchanged since apply

    const report = await rollbackChangeSet(changeSet);
    expect(report.conflicts).toHaveLength(0);
    expect(report.restoredCells).toBe(1);
    expect(state.writes).toContainEqual(
      expect.objectContaining({ sheet: "S", row: 26, col: 5, payload: [[10]] })
    );
  });

  it("restores over a human edit only when forced", async () => {
    const { rollbackChangeSet } = await import("../src/excel/writer");
    const changeSet = await applyThenHumanEdits(30);

    const report = await rollbackChangeSet(changeSet, { force: true });
    expect(report.conflicts).toHaveLength(0);
    expect(report.restoredCells).toBe(1);
    expect(state.writes).toContainEqual(
      expect.objectContaining({ sheet: "S", row: 26, col: 5, payload: [[10]] })
    );
  });

  it("refuses every cell when no post-apply state was recorded", async () => {
    const { rollbackChangeSet } = await import("../src/excel/writer");
    const changeSet = await applyThenHumanEdits(20);
    delete changeSet.appliedState;

    const report = await rollbackChangeSet(changeSet);
    expect(report.restoredCells).toBe(0);
    expect(state.writes).toHaveLength(0);
    expect(report.conflicts[0]?.reason).toContain("cannot tell");
  });
});

describe("_AI_Log (INV-10)", () => {
  it("creates the log sheet with headers on first write", async () => {
    const { writeAiLog } = await import("../src/excel/writer");
    const workbook = buildWorkbook();
    const changeSet = proposeChangeSet(workbook, {
      intent: "Change growth to 5.2%",
      summary: "Set FY2025 growth",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 0.052 }],
    });

    await writeAiLog(changeSet, "what changed and why");
    const headerWrite = state.writes.find(
      (write) => Array.isArray(write.payload) && String(write.payload).includes("Timestamp")
    );
    expect(headerWrite).toBeDefined();
    const entryWrite = state.writes.find(
      (write) => Array.isArray(write.payload) && String(write.payload).includes("Change growth")
    );
    expect(entryWrite).toBeDefined();
  });
});

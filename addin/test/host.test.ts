/**
 * OfficeJsWorkbookHost against the Office.js fake (P1-5).
 *
 * Same contract as engine/test/host/host.test.ts, exercised through the real
 * production code path — extraction, the writer, live drift, live rollback.
 *
 * What this proves: the production host implements the interface the agent
 * loop is written against, and the loop can therefore run unmodified in Excel.
 * What it does NOT prove: that Excel behaves like the fake. Coercion on write,
 * implicit intersection and calculated-column rewrites are precisely what the
 * fake cannot reproduce, which is why sideload item 13 exists.
 */

import { beforeEach, describe, expect, it, vi } from "vitest";
import { proposeChangeSet, Workbook } from "ledger-engine";

interface FakeState {
  reads: Map<string, { value: unknown; formula?: string }>;
  sheets: Set<string>;
  writes: number;
}

const state: FakeState = { reads: new Map(), sheets: new Set(), writes: 0 };

function makeRange(sheet: string, row: number, col: number) {
  const key = `${sheet}!${row},${col}`;
  const range: Record<string, unknown> = {
    values: [[null]],
    formulas: [[""]],
    numberFormat: [["General"]],
    load: vi.fn(() => {
      const stored = state.reads.get(key);
      range.values = [[stored?.value ?? null]];
      range.formulas = [[stored?.formula ?? stored?.value ?? ""]];
    }),
    untrack: vi.fn(),
    clear: vi.fn(() => state.reads.delete(key)),
  };
  return new Proxy(range, {
    set(target, prop, value) {
      if (prop === "values" || prop === "formulas") {
        state.writes++;
        const written = (value as unknown[][])[0]?.[0];
        state.reads.set(
          key,
          prop === "formulas" && typeof written === "string" && written.startsWith("=")
            ? { value: state.reads.get(key)?.value ?? null, formula: written }
            : { value: written }
        );
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
      isNullObject: false,
      address: "A1:B1",
      rowIndex: 0,
      rowCount: 1,
      columnIndex: 0,
      columnCount: 2,
      load: vi.fn(),
    }),
    delete: vi.fn(() => state.sheets.delete(name.toUpperCase())),
    protection: { protected: false, load: vi.fn() },
    isNullObject: false,
    load: vi.fn(),
  };
}

beforeEach(() => {
  state.reads = new Map([
    ["S!0,0", { value: 10 }],
    ["S!0,1", { value: 20, formula: "=A1*2" }],
  ]);
  state.sheets = new Set(["S"]);
  state.writes = 0;

  const context = {
    workbook: {
      worksheets: {
        getItem: (name: string) => makeSheet(name),
        getItemOrNullObject: (name: string) => ({
          ...makeSheet(name),
          isNullObject: !state.sheets.has(name.toUpperCase()),
        }),
        add: vi.fn((name: string) => {
          state.sheets.add(name.toUpperCase());
          return makeSheet(name);
        }),
        load: vi.fn(),
        items: [...state.sheets].map((name) => makeSheet(name)),
      },
      names: {
        add: vi.fn(),
        getItemOrNullObject: () => ({ isNullObject: true, load: vi.fn(), delete: vi.fn() }),
        load: vi.fn(),
        items: [],
      },
      tables: {
        add: vi.fn(() => ({ name: "" })),
        getItemOrNullObject: () => ({ isNullObject: true, load: vi.fn(), delete: vi.fn() }),
        load: vi.fn(),
        items: [],
      },
      application: { calculate: vi.fn() },
    },
    application: { suspendApiCalculationUntilNextSync: vi.fn() },
    sync: vi.fn(async () => undefined),
  };

  (globalThis as Record<string, unknown>).Excel = {
    run: async (callback: (context: unknown) => Promise<unknown>) => callback(context),
    ClearApplyTo: { contents: "contents" },
    CalculationType: { full: "full" },
    SheetVisibility: { visible: "Visible" },
  };
});

function workbookWithCells(): Workbook {
  const workbook = new Workbook("test");
  const sheet = workbook.addSheet("S");
  sheet.set({ row: 0, col: 0, value: 10 });
  sheet.set({ row: 0, col: 1, value: 20, formula: "=A1*2" });
  return workbook;
}

describe("OfficeJsWorkbookHost", () => {
  it("declares the capabilities the simulator lacks", async () => {
    const { OfficeJsWorkbookHost } = await import("../src/excel/host");
    const host = new OfficeJsWorkbookHost();
    expect(host.kind).toBe("office-js");
    // The reason the abstraction exists: this host sees things ours cannot.
    expect(host.capabilities.realRecalculation).toBe(true);
    expect(host.capabilities.coercesOnWrite).toBe(true);
    expect(host.capabilities.structuralEdits).toBe(true);
    expect(host.capabilities.observesHazards).toBe(true);
  });

  it("reports drift in the engine's shape", async () => {
    const { OfficeJsWorkbookHost } = await import("../src/excel/host");
    const host = new OfficeJsWorkbookHost();
    const changeSet = proposeChangeSet(workbookWithCells(), {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });

    expect(await host.checkDrift(changeSet)).toHaveLength(0);

    // Somebody else changes the formula.
    state.reads.set("S!0,1", { value: 99, formula: "=A1*99" });
    const drift = await host.checkDrift(changeSet);
    expect(drift).toHaveLength(1);
    expect(drift[0]!.expected.formula).toBe("=A1*2");
    expect(drift[0]!.actual.formula).toBe("=A1*99");
  });

  it("applies through the writer and records post-apply state", async () => {
    const { OfficeJsWorkbookHost } = await import("../src/excel/host");
    const host = new OfficeJsWorkbookHost();
    const changeSet = proposeChangeSet(workbookWithCells(), {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });

    const outcome = await host.apply(changeSet);
    expect(outcome.ok).toBe(true);
    expect(outcome.cellsWritten).toBe(1);
    // The evidence rollback needs, read back from the host (D-026).
    expect(changeSet.appliedState?.[0]?.formula).toBe("=A1*3");
  });

  it("surfaces a refused write instead of pretending it landed", async () => {
    const { OfficeJsWorkbookHost } = await import("../src/excel/host");
    const host = new OfficeJsWorkbookHost();
    const changeSet = proposeChangeSet(workbookWithCells(), {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setFormula", sheet: "S", row: 0, col: 1, formula: "=A1*3" }],
    });
    // Drift appears between propose and apply.
    state.reads.set("S!0,1", { value: 1, formula: "=SOMETHING()" });

    const outcome = await host.apply(changeSet);
    expect(outcome.ok).toBe(false);
    expect(outcome.cellsWritten).toBe(0);
    expect(outcome.failure).toContain("Someone else may be editing");
  });

  it("rolls back through the host in the engine's report shape", async () => {
    const { OfficeJsWorkbookHost } = await import("../src/excel/host");
    const host = new OfficeJsWorkbookHost();
    const changeSet = proposeChangeSet(workbookWithCells(), {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 42 }],
    });
    await host.apply(changeSet);

    const report = await host.rollback(changeSet);
    expect(report.changeSetId).toBe(changeSet.id);
    expect(report.restoredCells).toBe(1);
    expect(report.conflicts).toHaveLength(0);
    expect(report.reversedStructural).toEqual([]);
    expect(state.reads.get("S!0,0")?.value).toBe(10);
  });

  it("refuses to roll back over a human edit", async () => {
    const { OfficeJsWorkbookHost } = await import("../src/excel/host");
    const host = new OfficeJsWorkbookHost();
    const changeSet = proposeChangeSet(workbookWithCells(), {
      intent: "x",
      summary: "x",
      edits: [{ kind: "setValue", sheet: "S", row: 0, col: 0, value: 42 }],
    });
    await host.apply(changeSet);
    state.reads.set("S!0,0", { value: 999 }); // a human edits it

    const report = await host.rollback(changeSet);
    expect(report.restoredCells).toBe(0);
    expect(report.conflicts).toHaveLength(1);
    expect(state.reads.get("S!0,0")?.value).toBe(999);
  });
});

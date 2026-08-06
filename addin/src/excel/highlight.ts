/**
 * Trace highlighting: paint the dependency chain of a finding into the grid,
 * and take it away again cleanly.
 *
 * Highlighting mutates the workbook (fill colour), so it must be perfectly
 * reversible — we snapshot each touched cell's existing fill before painting
 * and restore exactly that on clear. Never leave a user's formatting altered
 * because they clicked "Trace".
 */

import type { CellRef } from "ledger-engine";

interface SavedFill {
  sheet: string;
  row: number;
  col: number;
  color: string;
}

const HIGHLIGHT_PRIMARY = "#FFF1C2"; // the finding itself
const HIGHLIGHT_CHAIN = "#DCEEFB"; // cells in its dependency chain

let saved: SavedFill[] = [];

/** Group refs by sheet so each sheet is one batch of range operations. */
function bySheet(refs: CellRef[]): Map<string, CellRef[]> {
  const out = new Map<string, CellRef[]>();
  for (const ref of refs) {
    let list = out.get(ref.sheet);
    if (!list) out.set(ref.sheet, (list = []));
    list.push(ref);
  }
  return out;
}

export async function highlightTrace(primary: CellRef, chain: CellRef[]): Promise<void> {
  await clearHighlight();

  await Excel.run(async (context) => {
    const next: SavedFill[] = [];
    // Deduplicate; the primary cell wins if it also appears in the chain.
    const seen = new Set<string>([`${primary.sheet}!${primary.row},${primary.col}`]);
    const chainRefs = chain.filter((ref) => {
      const key = `${ref.sheet}!${ref.row},${ref.col}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    const paint = async (refs: CellRef[], color: string): Promise<void> => {
      for (const [sheetName, cells] of bySheet(refs)) {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        // Read existing fills first so the clear is exact.
        const ranges = cells.map((ref) => {
          const range = sheet.getRangeByIndexes(ref.row, ref.col, 1, 1);
          range.format.fill.load("color");
          return { ref, range };
        });
        await context.sync();

        for (const { ref, range } of ranges) {
          next.push({
            sheet: sheetName,
            row: ref.row,
            col: ref.col,
            color: range.format.fill.color,
          });
          range.format.fill.color = color;
        }
        await context.sync();
        for (const { range } of ranges) range.untrack();
      }
    };

    await paint(chainRefs, HIGHLIGHT_CHAIN);
    await paint([primary], HIGHLIGHT_PRIMARY);
    saved = next;
  });
}

export async function clearHighlight(): Promise<void> {
  if (saved.length === 0) return;
  const toRestore = saved;
  saved = [];

  await Excel.run(async (context) => {
    for (const [sheetName, cells] of bySheet(
      toRestore.map((item) => ({ sheet: item.sheet, row: item.row, col: item.col }))
    )) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      for (const cell of cells) {
        const original = toRestore.find(
          (item) => item.sheet === cell.sheet && item.row === cell.row && item.col === cell.col
        );
        if (!original) continue;
        const range = sheet.getRangeByIndexes(cell.row, cell.col, 1, 1);
        // An empty colour string means "no fill" — clearing is the honest
        // restore, since setting "" would throw.
        if (original.color && original.color !== "#FFFFFF") {
          range.format.fill.color = original.color;
        } else {
          range.format.fill.clear();
        }
      }
      await context.sync();
    }
  });
}

export function hasActiveHighlight(): boolean {
  return saved.length > 0;
}

/** Scroll to and select a cell, so "Trace" also navigates there. */
export async function navigateTo(ref: CellRef): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(ref.sheet);
    sheet.activate();
    const range = sheet.getRangeByIndexes(ref.row, ref.col, 1, 1);
    range.select();
    await context.sync();
  });
}

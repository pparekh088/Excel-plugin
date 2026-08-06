/**
 * Broken-variant generators: take a clean model and inject a specific,
 * labelled defect. Each returns the expected (rule, address) pair, which is
 * the ground truth the audit evals score precision and recall against.
 *
 * Injections are deliberately realistic — the mistakes analysts actually make
 * (a plugged number over a formula, a copy that missed the last column, a
 * deleted row leaving #REF!) rather than synthetic noise.
 */

import { Workbook, a1, parseA1Range } from "../model/workbook";
import { CorpusWorkbook, InjectedDefect } from "./models";

function write(
  workbook: Workbook,
  sheetName: string,
  address: string,
  value: string | number | boolean | null
): void {
  const parsed = parseA1Range(address);
  if (!parsed) throw new Error(`bad address ${address}`);
  const sheet = workbook.addSheet(sheetName);
  const isFormula = typeof value === "string" && value.startsWith("=");
  sheet.set({
    row: parsed.startRow,
    col: parsed.startCol,
    value: isFormula ? 0 : value,
    ...(isFormula ? { formula: value } : {}),
  });
}

export type Injector = (workbook: Workbook) => InjectedDefect[];

/** AUD-001: overwrite a mid-run formula cell so it breaks the fill pattern. */
export function injectFormulaInconsistency(
  sheet: string,
  address: string,
  replacement: string
): Injector {
  return (workbook) => {
    write(workbook, sheet, address, replacement);
    return [
      {
        rule: "AUD-001",
        address: `${sheet}!${address}`,
        note: `formula replaced with '${replacement}', breaking the surrounding run`,
      },
    ];
  };
}

/** AUD-002: bake a magic number into an otherwise clean formula. */
export function injectHardcodedInFormula(
  sheet: string,
  address: string,
  replacement: string
): Injector {
  return (workbook) => {
    write(workbook, sheet, address, replacement);
    return [
      {
        rule: "AUD-002",
        address: `${sheet}!${address}`,
        note: `magic number embedded in formula '${replacement}'`,
      },
    ];
  };
}

/** AUD-003: replace a formula with a plugged constant that still has dependents. */
export function injectPluggedConstant(
  sheet: string,
  address: string,
  value: number
): Injector {
  return (workbook) => {
    write(workbook, sheet, address, value);
    return [
      {
        rule: "AUD-003",
        address: `${sheet}!${address}`,
        note: `formula replaced by hardcoded ${value} inside a calculation chain`,
      },
    ];
  };
}

/** AUD-004: simulate a deleted precedent, leaving #REF! behind. */
export function injectRefError(sheet: string, address: string): Injector {
  return (workbook) => {
    write(workbook, sheet, address, "=#REF!*2");
    const target = workbook.addSheet(sheet);
    const parsed = parseA1Range(address)!;
    const cell = target.get(parsed.startRow, parsed.startCol);
    if (cell) cell.value = "#REF!";
    return [
      {
        rule: "AUD-004",
        address: `${sheet}!${address}`,
        note: "precedent deleted, formula now yields #REF!",
      },
    ];
  };
}

/**
 * AUD-005: close a loop between two cells. The engine reports one finding per
 * cycle GROUP, anchored at whichever member it reaches first, so both members
 * are acceptable addresses for the match.
 */
export function injectCircular(sheet: string, first: string, second: string): Injector {
  return (workbook) => {
    write(workbook, sheet, first, `=${second}+1`);
    write(workbook, sheet, second, `=${first}*2`);
    return [
      {
        rule: "AUD-005",
        address: `${sheet}!${first}`,
        alsoAcceptable: [`${sheet}!${second}`],
        note: `circular reference ${first} <-> ${second}`,
      },
    ];
  };
}

/** AUD-006: point a formula at an empty cell. */
export function injectEmptyRef(
  sheet: string,
  address: string,
  emptyTarget: string
): Injector {
  return (workbook) => {
    write(workbook, sheet, address, `=${emptyTarget}*2`);
    return [
      {
        rule: "AUD-006",
        address: `${sheet}!${address}`,
        note: `formula references empty cell ${emptyTarget}`,
      },
    ];
  };
}

/**
 * AUD-008: break the balance-sheet tie-out. The assertion surfaces at the
 * CHECK cell, not at the edit that broke it, so `checkCells` carries the
 * addresses where the finding is expected to appear.
 */
export function injectBalanceBreak(
  sheet: string,
  address: string,
  delta: number,
  checkCells: string[]
): Injector {
  return (workbook) => {
    const parsed = parseA1Range(address)!;
    const target = workbook.addSheet(sheet);
    const cell = target.get(parsed.startRow, parsed.startCol);
    const existing = cell?.formula ?? "=0";
    write(workbook, sheet, address, `${existing}+${delta}`);
    const [primary, ...rest] = checkCells;
    return [
      {
        rule: "AUD-008",
        address: `${sheet}!${primary}`,
        alsoAcceptable: rest.map((cellAddress) => `${sheet}!${cellAddress}`),
        // The edited cell now carries a literal and breaks its row's pattern,
        // so AUD-001/AUD-002 firing there is correct, not a false positive.
        collateral: [`${sheet}!${address}`],
        note: `balance assertion broken by ${delta} at ${sheet}!${address}`,
      },
    ];
  };
}

/** AUD-009: sprinkle volatile functions through a model. */
export function injectVolatiles(sheet: string, addresses: string[]): Injector {
  return (workbook) => {
    return addresses.map((address) => {
      write(workbook, sheet, address, "=OFFSET($A$1,1,1)*NOW()");
      return {
        rule: "AUD-009",
        address: `${sheet}!${address}`,
        note: "volatile functions (OFFSET, NOW) introduced",
      };
    });
  };
}

/** AUD-011: add an external workbook link. */
export function injectExternalLink(sheet: string, address: string): Injector {
  return (workbook) => {
    write(workbook, sheet, address, "=[Legacy Model.xlsx]Summary!$B$4*1.1");
    return [
      {
        rule: "AUD-011",
        address: `${sheet}!${address}`,
        note: "external workbook reference",
      },
    ];
  };
}

/** Build a broken variant from a clean generator plus injectors. */
export function brokenVariant(
  base: CorpusWorkbook,
  id: string,
  title: string,
  injectors: Injector[]
): CorpusWorkbook {
  const defects: InjectedDefect[] = [];
  for (const inject of injectors) defects.push(...inject(base.workbook));
  return { id, title, workbook: base.workbook, defects };
}

/** Column letter for a 0-based index — corpus builders work in letters. */
export function columnLetter(index: number): string {
  return a1(0, index).replace(/[0-9]+$/, "");
}

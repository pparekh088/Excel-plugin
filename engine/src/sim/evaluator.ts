/**
 * Formula evaluator for the headless Excel simulator (handoff §9).
 *
 * Scope is deliberately the function subset the eval corpus uses, plus the
 * operators. It exists so CI can run the whole tool runtime — audit, change
 * sets, agent loop — against an in-memory workbook with no Excel present.
 *
 * It is NOT a spreadsheet engine and must never be presented as one: an
 * unsupported function evaluates to #NAME? and is reported, so a corpus that
 * outgrows the evaluator fails loudly instead of scoring against nonsense.
 */

import {
  ArrayLit,
  Node,
  NO_SHEET,
  SheetSpec,
} from "../parser/ast";
import { parseFormula } from "../parser/parser";
import { CellValue, Workbook } from "../model/workbook";
import { resolveExtracted } from "../graph/resolve";
import { extractRefs } from "../parser/refs";

export type ErrorText =
  | "#REF!"
  | "#VALUE!"
  | "#DIV/0!"
  | "#NAME?"
  | "#N/A"
  | "#NUM!"
  | "#NULL!"
  | "#CIRCULAR!";

export const ERRORS: Record<string, ErrorText> = {
  ref: "#REF!",
  value: "#VALUE!",
  div0: "#DIV/0!",
  name: "#NAME?",
  na: "#N/A",
  num: "#NUM!",
  nul: "#NULL!",
  circular: "#CIRCULAR!",
};

const ERROR_SET = new Set<string>(Object.values(ERRORS));

export function isError(value: unknown): value is ErrorText {
  return typeof value === "string" && ERROR_SET.has(value);
}

/** A scalar or a 2D block (ranges and array literals evaluate to blocks). */
export type EvalValue = CellValue | CellValue[][];

export interface EvalContext {
  workbook: Workbook;
  sheet: string;
  row: number;
  col: number;
  /** Resolves a cell to its current value; drives recalculation order. */
  read: (sheet: string, row: number, col: number) => CellValue;
  /** LET/LAMBDA bindings in scope. */
  locals: Map<string, EvalValue>;
}

// -------------------------------------------------------------- coercion

function flatten(value: EvalValue): CellValue[] {
  if (Array.isArray(value)) return value.flat();
  return [value];
}

/** Excel numeric coercion: "" -> 0, "12" -> 12, TRUE -> 1, text -> #VALUE!. */
export function toNumber(value: EvalValue): number | ErrorText {
  if (Array.isArray(value)) {
    const first = value[0]?.[0];
    return first === undefined ? ERRORS.value! : toNumber(first);
  }
  if (isError(value)) return value;
  if (value === null || value === "") return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const trimmed = value.trim();
  if (trimmed === "") return 0;
  const parsed = Number(trimmed);
  return Number.isNaN(parsed) ? ERRORS.value! : parsed;
}

export function toText(value: EvalValue): string | ErrorText {
  if (Array.isArray(value)) {
    const first = value[0]?.[0];
    return first === undefined ? "" : toText(first);
  }
  if (isError(value)) return value;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

export function toBoolean(value: EvalValue): boolean | ErrorText {
  if (Array.isArray(value)) {
    const first = value[0]?.[0];
    return first === undefined ? false : toBoolean(first);
  }
  if (isError(value)) return value;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (value === null || value === "") return false;
  const upper = value.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return ERRORS.value!;
}

/** Numbers inside a range, skipping text/blanks the way SUM/AVERAGE do. */
function numericsOf(values: EvalValue[]): number[] | ErrorText {
  const out: number[] = [];
  for (const value of values) {
    for (const item of flatten(value)) {
      if (isError(item)) return item;
      if (typeof item === "number") out.push(item);
      else if (typeof item === "boolean") continue; // booleans are skipped in ranges
      else if (item === null || item === "") continue;
      else if (typeof item === "string") {
        const parsed = Number(item);
        if (!Number.isNaN(parsed) && item.trim() !== "") out.push(parsed);
      }
    }
  }
  return out;
}

// --------------------------------------------------------------- criteria

/** COUNTIF/SUMIF criteria: ">5", "<>x", "abc", 42. */
function matchesCriterion(value: CellValue, criterion: EvalValue): boolean {
  const raw = Array.isArray(criterion) ? criterion[0]?.[0] ?? null : criterion;
  if (typeof raw === "string") {
    const operator = /^(<=|>=|<>|<|>|=)/.exec(raw);
    if (operator) {
      const rest = raw.slice(operator[0].length);
      const target: CellValue = rest.trim() === "" ? "" : (Number.isNaN(Number(rest)) ? rest : Number(rest));
      const comparison = compareValues(value, target);
      if (comparison === null) return operator[0] === "<>";
      switch (operator[0]) {
        case "<":
          return comparison < 0;
        case "<=":
          return comparison <= 0;
        case ">":
          return comparison > 0;
        case ">=":
          return comparison >= 0;
        case "=":
          return comparison === 0;
        case "<>":
          return comparison !== 0;
      }
    }
  }
  const comparison = compareValues(value, Array.isArray(criterion) ? null : criterion);
  return comparison === 0;
}

/** Excel ordering: numbers < text; blanks coerce. null when incomparable. */
function compareValues(left: CellValue, right: CellValue): number | null {
  const l = left === null ? "" : left;
  const r = right === null ? "" : right;
  if (typeof l === "number" && typeof r === "number") return l === r ? 0 : l < r ? -1 : 1;
  if (typeof l === "boolean" || typeof r === "boolean") {
    const ln = typeof l === "boolean" ? (l ? 1 : 0) : l;
    const rn = typeof r === "boolean" ? (r ? 1 : 0) : r;
    return ln === rn ? 0 : ln < rn ? -1 : 1;
  }
  if (typeof l === "number" && typeof r === "string") {
    if (r === "") return l === 0 ? 0 : l < 0 ? -1 : 1;
    return -1;
  }
  if (typeof l === "string" && typeof r === "number") {
    if (l === "") return r === 0 ? 0 : r > 0 ? -1 : 1;
    return 1;
  }
  const ls = String(l).toUpperCase();
  const rs = String(r).toUpperCase();
  return ls === rs ? 0 : ls < rs ? -1 : 1;
}

// -------------------------------------------------------------- functions

type FnArgs = Array<() => EvalValue>;

export const SUPPORTED_FUNCTIONS = new Set([
  "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "ROUND", "ROUNDUP",
  "ROUNDDOWN", "ABS", "SQRT", "POWER", "EXP", "LN", "LOG", "LOG10", "MOD",
  "INT", "IF", "IFERROR", "IFNA", "AND", "OR", "NOT", "TRUE", "FALSE",
  "CONCATENATE", "CONCAT", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER",
  "LOWER", "TEXT", "VALUE", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
  "AVERAGEIF", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP", "CHOOSE",
  "SUMPRODUCT", "NPV", "PMT", "SIGN", "MEDIAN", "STDEV", "VAR", "SMALL",
  "LARGE", "RANK", "ISERROR", "ISNUMBER", "ISBLANK", "ISTEXT", "NA",
  "SUBTOTAL", "LET", "N", "T",
]);

function evalFunction(name: string, args: FnArgs, context: EvalContext): EvalValue {
  const all = (): EvalValue[] => args.map((arg) => arg());
  const first = (): EvalValue => (args[0] ? args[0]() : null);

  const numericArgs = (): number[] | ErrorText => numericsOf(all());
  const scalarNumber = (index: number): number | ErrorText => {
    const arg = args[index];
    return arg ? toNumber(arg()) : 0;
  };

  switch (name) {
    case "SUM": {
      const numbers = numericArgs();
      return isError(numbers) ? numbers : numbers.reduce((a, b) => a + b, 0);
    }
    case "AVERAGE": {
      const numbers = numericArgs();
      if (isError(numbers)) return numbers;
      if (numbers.length === 0) return ERRORS.div0!;
      return numbers.reduce((a, b) => a + b, 0) / numbers.length;
    }
    case "MIN": {
      const numbers = numericArgs();
      if (isError(numbers)) return numbers;
      return numbers.length === 0 ? 0 : Math.min(...numbers);
    }
    case "MAX": {
      const numbers = numericArgs();
      if (isError(numbers)) return numbers;
      return numbers.length === 0 ? 0 : Math.max(...numbers);
    }
    case "MEDIAN": {
      const numbers = numericArgs();
      if (isError(numbers)) return numbers;
      if (numbers.length === 0) return ERRORS.num!;
      const sorted = [...numbers].sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 === 0 ? (sorted[mid - 1]! + sorted[mid]!) / 2 : sorted[mid]!;
    }
    case "COUNT": {
      const numbers = numericArgs();
      return isError(numbers) ? numbers : numbers.length;
    }
    case "COUNTA": {
      let count = 0;
      for (const value of all()) {
        for (const item of flatten(value)) {
          if (item !== null && item !== "") count++;
        }
      }
      return count;
    }
    case "ROUND":
    case "ROUNDUP":
    case "ROUNDDOWN": {
      const value = scalarNumber(0);
      const digits = args.length > 1 ? scalarNumber(1) : 0;
      if (isError(value)) return value;
      if (isError(digits)) return digits;
      const factor = 10 ** digits;
      const scaled = value * factor;
      const rounded =
        name === "ROUND"
          ? Math.sign(scaled) * Math.round(Math.abs(scaled))
          : name === "ROUNDUP"
            ? Math.sign(scaled) * Math.ceil(Math.abs(scaled))
            : Math.sign(scaled) * Math.floor(Math.abs(scaled));
      return rounded / factor;
    }
    case "ABS": {
      const value = scalarNumber(0);
      return isError(value) ? value : Math.abs(value);
    }
    case "SIGN": {
      const value = scalarNumber(0);
      return isError(value) ? value : Math.sign(value);
    }
    case "SQRT": {
      const value = scalarNumber(0);
      if (isError(value)) return value;
      return value < 0 ? ERRORS.num! : Math.sqrt(value);
    }
    case "POWER": {
      const base = scalarNumber(0);
      const exponent = scalarNumber(1);
      if (isError(base)) return base;
      if (isError(exponent)) return exponent;
      return base ** exponent;
    }
    case "EXP": {
      const value = scalarNumber(0);
      return isError(value) ? value : Math.exp(value);
    }
    case "LN": {
      const value = scalarNumber(0);
      if (isError(value)) return value;
      return value <= 0 ? ERRORS.num! : Math.log(value);
    }
    case "LOG10": {
      const value = scalarNumber(0);
      if (isError(value)) return value;
      return value <= 0 ? ERRORS.num! : Math.log10(value);
    }
    case "LOG": {
      const value = scalarNumber(0);
      const base = args.length > 1 ? scalarNumber(1) : 10;
      if (isError(value)) return value;
      if (isError(base)) return base;
      return value <= 0 ? ERRORS.num! : Math.log(value) / Math.log(base);
    }
    case "MOD": {
      const value = scalarNumber(0);
      const divisor = scalarNumber(1);
      if (isError(value)) return value;
      if (isError(divisor)) return divisor;
      if (divisor === 0) return ERRORS.div0!;
      return value - divisor * Math.floor(value / divisor);
    }
    case "INT": {
      const value = scalarNumber(0);
      return isError(value) ? value : Math.floor(value);
    }
    case "IF": {
      const condition = toBoolean(first());
      if (isError(condition)) return condition;
      if (condition) return args[1] ? args[1]() : true;
      return args[2] ? args[2]() : false;
    }
    case "IFERROR": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      return isError(scalar) ? (args[1] ? args[1]() : "") : value;
    }
    case "IFNA": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      return scalar === ERRORS.na ? (args[1] ? args[1]() : "") : value;
    }
    case "AND": {
      for (const value of all()) {
        for (const item of flatten(value)) {
          if (isError(item)) return item;
          if (item === null || item === "") continue;
          const bool = toBoolean(item);
          if (isError(bool)) return bool;
          if (!bool) return false;
        }
      }
      return true;
    }
    case "OR": {
      let sawAny = false;
      for (const value of all()) {
        for (const item of flatten(value)) {
          if (isError(item)) return item;
          if (item === null || item === "") continue;
          const bool = toBoolean(item);
          if (isError(bool)) return bool;
          sawAny = true;
          if (bool) return true;
        }
      }
      return sawAny ? false : false;
    }
    case "NOT": {
      const bool = toBoolean(first());
      return isError(bool) ? bool : !bool;
    }
    case "TRUE":
      return true;
    case "FALSE":
      return false;
    case "NA":
      return ERRORS.na!;
    case "CONCATENATE":
    case "CONCAT": {
      let out = "";
      for (const value of all()) {
        for (const item of flatten(value)) {
          const text = toText(item);
          if (isError(text)) return text;
          out += text;
        }
      }
      return out;
    }
    case "LEFT":
    case "RIGHT": {
      const text = toText(first());
      const count = args.length > 1 ? scalarNumber(1) : 1;
      if (isError(text)) return text;
      if (isError(count)) return count;
      return name === "LEFT" ? text.slice(0, count) : text.slice(Math.max(0, text.length - count));
    }
    case "MID": {
      const text = toText(first());
      const start = scalarNumber(1);
      const count = scalarNumber(2);
      if (isError(text)) return text;
      if (isError(start)) return start;
      if (isError(count)) return count;
      return text.slice(start - 1, start - 1 + count);
    }
    case "LEN": {
      const text = toText(first());
      return isError(text) ? text : text.length;
    }
    case "TRIM": {
      const text = toText(first());
      return isError(text) ? text : text.trim().replace(/\s+/g, " ");
    }
    case "UPPER": {
      const text = toText(first());
      return isError(text) ? text : text.toUpperCase();
    }
    case "LOWER": {
      const text = toText(first());
      return isError(text) ? text : text.toLowerCase();
    }
    case "VALUE": {
      const value = toNumber(first());
      return value;
    }
    case "N": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      if (isError(scalar)) return scalar;
      if (typeof scalar === "number") return scalar;
      if (typeof scalar === "boolean") return scalar ? 1 : 0;
      return 0;
    }
    case "T": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      if (isError(scalar)) return scalar;
      return typeof scalar === "string" ? scalar : "";
    }
    case "TEXT": {
      // Enough of the format grammar for the corpus: 0.00, 0%, #,##0.
      const value = toNumber(first());
      const format = toText(args[1] ? args[1]() : "");
      if (isError(value)) return value;
      if (isError(format)) return format;
      return formatNumber(value, format);
    }
    case "ISERROR": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      return isError(scalar);
    }
    case "ISNUMBER": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      return typeof scalar === "number";
    }
    case "ISTEXT": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      return typeof scalar === "string" && !isError(scalar);
    }
    case "ISBLANK": {
      const value = first();
      const scalar = Array.isArray(value) ? value[0]?.[0] ?? null : value;
      return scalar === null || scalar === "";
    }
    case "SUMIF":
    case "AVERAGEIF": {
      const range = flatten(first());
      const criterion = args[1] ? args[1]() : null;
      const sumRange = args[2] ? flatten(args[2]()) : range;
      const picked: number[] = [];
      range.forEach((item, index) => {
        if (!matchesCriterion(item, criterion)) return;
        const target = sumRange[index];
        if (typeof target === "number") picked.push(target);
      });
      if (name === "SUMIF") return picked.reduce((a, b) => a + b, 0);
      return picked.length === 0
        ? ERRORS.div0!
        : picked.reduce((a, b) => a + b, 0) / picked.length;
    }
    case "SUMIFS": {
      const sumRange = flatten(first());
      const conditions: Array<{ range: CellValue[]; criterion: EvalValue }> = [];
      for (let i = 1; i + 1 < args.length; i += 2) {
        conditions.push({ range: flatten(args[i]!()), criterion: args[i + 1]!() });
      }
      let total = 0;
      sumRange.forEach((item, index) => {
        const ok = conditions.every((condition) =>
          matchesCriterion(condition.range[index] ?? null, condition.criterion)
        );
        if (ok && typeof item === "number") total += item;
      });
      return total;
    }
    case "COUNTIF": {
      const range = flatten(first());
      const criterion = args[1] ? args[1]() : null;
      return range.filter((item) => matchesCriterion(item, criterion)).length;
    }
    case "COUNTIFS": {
      const conditions: Array<{ range: CellValue[]; criterion: EvalValue }> = [];
      for (let i = 0; i + 1 < args.length; i += 2) {
        conditions.push({ range: flatten(args[i]!()), criterion: args[i + 1]!() });
      }
      const length = conditions[0]?.range.length ?? 0;
      let count = 0;
      for (let index = 0; index < length; index++) {
        if (
          conditions.every((condition) =>
            matchesCriterion(condition.range[index] ?? null, condition.criterion)
          )
        ) {
          count++;
        }
      }
      return count;
    }
    case "SUMPRODUCT": {
      const arrays = all().map((value) => flatten(value));
      const length = arrays[0]?.length ?? 0;
      let total = 0;
      for (let index = 0; index < length; index++) {
        let product = 1;
        for (const array of arrays) {
          const item = array[index];
          const numeric =
            typeof item === "number" ? item : typeof item === "boolean" ? (item ? 1 : 0) : 0;
          product *= numeric;
        }
        total += product;
      }
      return total;
    }
    case "NPV": {
      const rate = scalarNumber(0);
      if (isError(rate)) return rate;
      const flows = numericsOf(all().slice(1));
      if (isError(flows)) return flows;
      return flows.reduce((sum, flow, index) => sum + flow / (1 + rate) ** (index + 1), 0);
    }
    case "PMT": {
      const rate = scalarNumber(0);
      const periods = scalarNumber(1);
      const present = scalarNumber(2);
      if (isError(rate)) return rate;
      if (isError(periods)) return periods;
      if (isError(present)) return present;
      if (rate === 0) return -present / periods;
      return (-present * rate) / (1 - (1 + rate) ** -periods);
    }
    case "INDEX": {
      const block = first();
      const rowIndex = args.length > 1 ? scalarNumber(1) : 1;
      const colIndex = args.length > 2 ? scalarNumber(2) : 1;
      if (isError(rowIndex)) return rowIndex;
      if (isError(colIndex)) return colIndex;
      if (!Array.isArray(block)) return rowIndex <= 1 && colIndex <= 1 ? block : ERRORS.ref!;
      // A single-row or single-column block indexes linearly.
      if (block.length === 1 && args.length === 2) {
        return block[0]![rowIndex - 1] ?? ERRORS.ref!;
      }
      if ((block[0]?.length ?? 0) === 1 && args.length === 2) {
        return block[rowIndex - 1]?.[0] ?? ERRORS.ref!;
      }
      return block[rowIndex - 1]?.[colIndex - 1] ?? ERRORS.ref!;
    }
    case "MATCH": {
      const target = first();
      const block = args[1] ? args[1]() : null;
      const mode = args.length > 2 ? scalarNumber(2) : 1;
      if (isError(mode)) return mode;
      const list = flatten(block);
      const scalar = Array.isArray(target) ? target[0]?.[0] ?? null : target;
      if (mode === 0) {
        const index = list.findIndex((item) => compareValues(item, scalar) === 0);
        return index < 0 ? ERRORS.na! : index + 1;
      }
      // Approximate match on a sorted list.
      let best = -1;
      list.forEach((item, index) => {
        const comparison = compareValues(item, scalar);
        if (comparison === null) return;
        if (mode >= 1 ? comparison <= 0 : comparison >= 0) best = index;
      });
      return best < 0 ? ERRORS.na! : best + 1;
    }
    case "VLOOKUP":
    case "HLOOKUP": {
      const target = first();
      const block = args[1] ? args[1]() : null;
      const index = args.length > 2 ? scalarNumber(2) : 1;
      const approximate = args.length > 3 ? toBoolean(args[3]!()) : true;
      if (isError(index)) return index;
      if (isError(approximate)) return approximate;
      if (!Array.isArray(block)) return ERRORS.na!;
      const scalar = Array.isArray(target) ? target[0]?.[0] ?? null : target;

      if (name === "VLOOKUP") {
        let match = -1;
        for (let row = 0; row < block.length; row++) {
          const comparison = compareValues(block[row]![0] ?? null, scalar);
          if (comparison === 0) {
            match = row;
            break;
          }
          if (approximate && comparison !== null && comparison < 0) match = row;
        }
        if (match < 0) return ERRORS.na!;
        return block[match]![index - 1] ?? ERRORS.ref!;
      }
      const headerRow = block[0] ?? [];
      let match = -1;
      for (let col = 0; col < headerRow.length; col++) {
        const comparison = compareValues(headerRow[col] ?? null, scalar);
        if (comparison === 0) {
          match = col;
          break;
        }
        if (approximate && comparison !== null && comparison < 0) match = col;
      }
      if (match < 0) return ERRORS.na!;
      return block[index - 1]?.[match] ?? ERRORS.ref!;
    }
    case "XLOOKUP": {
      const target = first();
      const lookup = flatten(args[1] ? args[1]() : null);
      const results = flatten(args[2] ? args[2]() : null);
      const scalar = Array.isArray(target) ? target[0]?.[0] ?? null : target;
      const index = lookup.findIndex((item) => compareValues(item, scalar) === 0);
      if (index < 0) return args[3] ? args[3]() : ERRORS.na!;
      return results[index] ?? ERRORS.ref!;
    }
    case "CHOOSE": {
      const index = scalarNumber(0);
      if (isError(index)) return index;
      const chosen = args[index];
      return chosen ? chosen() : ERRORS.value!;
    }
    case "SMALL":
    case "LARGE": {
      const numbers = numericsOf([first()]);
      const k = scalarNumber(1);
      if (isError(numbers)) return numbers;
      if (isError(k)) return k;
      const sorted = [...numbers].sort((a, b) => (name === "SMALL" ? a - b : b - a));
      return sorted[k - 1] ?? ERRORS.num!;
    }
    case "RANK": {
      const value = scalarNumber(0);
      const numbers = numericsOf([args[1] ? args[1]() : null]);
      if (isError(value)) return value;
      if (isError(numbers)) return numbers;
      const ascending = args.length > 2 ? toBoolean(args[2]!()) : false;
      if (isError(ascending)) return ascending;
      const sorted = [...numbers].sort((a, b) => (ascending ? a - b : b - a));
      const index = sorted.indexOf(value);
      return index < 0 ? ERRORS.na! : index + 1;
    }
    case "STDEV":
    case "VAR": {
      const numbers = numericArgs();
      if (isError(numbers)) return numbers;
      if (numbers.length < 2) return ERRORS.div0!;
      const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
      const variance =
        numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - 1);
      return name === "VAR" ? variance : Math.sqrt(variance);
    }
    case "SUBTOTAL": {
      const code = scalarNumber(0);
      if (isError(code)) return code;
      const rest = args.slice(1).map((arg) => arg());
      const numbers = numericsOf(rest);
      if (isError(numbers)) return numbers;
      const base = code > 100 ? code - 100 : code;
      switch (base) {
        case 1:
          return numbers.length === 0
            ? ERRORS.div0!
            : numbers.reduce((a, b) => a + b, 0) / numbers.length;
        case 2:
          return numbers.length;
        case 4:
          return numbers.length === 0 ? 0 : Math.max(...numbers);
        case 5:
          return numbers.length === 0 ? 0 : Math.min(...numbers);
        case 9:
          return numbers.reduce((a, b) => a + b, 0);
        default:
          return ERRORS.value!;
      }
    }
    default:
      // Unsupported function: fail loudly rather than silently returning 0.
      void context;
      return ERRORS.name!;
  }
}

function formatNumber(value: number, format: string): string {
  if (format.includes("%")) {
    const decimals = (/\.(0+)/.exec(format)?.[1] ?? "").length;
    return `${(value * 100).toFixed(decimals)}%`;
  }
  const decimals = (/\.(0+)/.exec(format)?.[1] ?? "").length;
  const fixed = value.toFixed(decimals);
  if (format.includes(",")) {
    const [whole, fraction] = fixed.split(".");
    const grouped = (whole ?? "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return fraction ? `${grouped}.${fraction}` : grouped;
  }
  return fixed;
}

// ------------------------------------------------------------- evaluation

function sheetOf(spec: SheetSpec, context: EvalContext): string | null {
  if (spec.external !== null) return null;
  if (spec.start === null) return context.sheet;
  const sheet = context.workbook.sheet(spec.start);
  return sheet ? sheet.name : null;
}

function readBlock(
  context: EvalContext,
  sheet: string,
  startRow: number,
  endRow: number,
  startCol: number,
  endCol: number
): EvalValue {
  if (startRow === endRow && startCol === endCol) {
    return context.read(sheet, startRow, endRow === startRow ? startCol : startCol);
  }
  const block: CellValue[][] = [];
  for (let row = startRow; row <= endRow; row++) {
    const line: CellValue[] = [];
    for (let col = startCol; col <= endCol; col++) line.push(context.read(sheet, row, col));
    block.push(line);
  }
  return block;
}

export function evaluate(node: Node, context: EvalContext): EvalValue {
  switch (node.kind) {
    case "number":
      return node.value;
    case "string":
      return node.value;
    case "bool":
      return node.value;
    case "error":
      return node.value;
    case "missing":
      return null;
    case "group":
      return evaluate(node.expr, context);
    case "cell": {
      const sheet = sheetOf(node.sheet, context);
      if (sheet === null) return ERRORS.ref!;
      return context.read(sheet, node.addr.row, node.addr.col);
    }
    case "range": {
      const sheet = sheetOf(node.sheet, context);
      if (sheet === null) return ERRORS.ref!;
      return readBlock(
        context,
        sheet,
        Math.min(node.start.row, node.end.row),
        Math.max(node.start.row, node.end.row),
        Math.min(node.start.col, node.end.col),
        Math.max(node.start.col, node.end.col)
      );
    }
    case "colRange":
    case "rowRange":
    case "name":
    case "structured": {
      // Resolve through the shared resolver so names/tables behave identically
      // to the graph's view of them.
      const resolution = resolveExtracted(
        extractRefs(node),
        {
          workbook: context.workbook,
          hostSheet: context.sheet,
          hostRow: context.row,
        }
      );
      if (resolution.areas.length === 0) {
        if (node.kind === "name") {
          const local = context.locals.get(node.name.toUpperCase());
          if (local !== undefined) return local;
        }
        return ERRORS.name!;
      }
      const area = resolution.areas[0]!;
      const target = context.workbook.sheet(area.sheet);
      if (!target) return ERRORS.ref!;
      const bounds = target.bounds();
      const startRow = area.startRow ?? bounds?.startRow ?? 0;
      const endRow = area.endRow ?? bounds?.endRow ?? 0;
      const startCol = area.startCol ?? bounds?.startCol ?? 0;
      const endCol = area.endCol ?? bounds?.endCol ?? 0;
      return readBlock(context, target.name, startRow, endRow, startCol, endCol);
    }
    case "array": {
      const literal = node as ArrayLit;
      return literal.rows.map((row) =>
        row.map((item) => {
          const value = evaluate(item, context);
          return Array.isArray(value) ? (value[0]?.[0] ?? null) : value;
        })
      );
    }
    case "unary": {
      const value = toNumber(evaluate(node.operand, context));
      if (isError(value)) return value;
      return node.op === "-" ? -value : value;
    }
    case "percent": {
      const value = toNumber(evaluate(node.operand, context));
      return isError(value) ? value : value / 100;
    }
    case "implicitIntersection": {
      const value = evaluate(node.operand, context);
      return Array.isArray(value) ? (value[0]?.[0] ?? null) : value;
    }
    case "spill":
      return evaluate(node.operand, context);
    case "binary":
      return evaluateBinary(node.op, node.left, node.right, context);
    case "func": {
      if (node.name === "LET") return evaluateLet(node.args, context);
      if (node.name === "LAMBDA") return ERRORS.value!; // only callable form supported
      const args: FnArgs = node.args.map((arg) => () => evaluate(arg, context));
      return evalFunction(node.name, args, context);
    }
    case "callExpr": {
      // LAMBDA(params, body)(args)
      if (node.callee.kind === "func" && node.callee.name === "LAMBDA") {
        const params = node.callee.args.slice(0, -1);
        const body = node.callee.args[node.callee.args.length - 1];
        if (!body) return ERRORS.value!;
        const locals = new Map(context.locals);
        params.forEach((param, index) => {
          if (param.kind !== "name") return;
          const arg = node.args[index];
          locals.set(param.name.toUpperCase(), arg ? evaluate(arg, context) : null);
        });
        return evaluate(body, { ...context, locals });
      }
      return ERRORS.value!;
    }
    case "bad":
      return ERRORS.name!;
  }
}

function evaluateLet(args: Node[], context: EvalContext): EvalValue {
  const locals = new Map(context.locals);
  let index = 0;
  while (index + 1 < args.length) {
    const nameNode = args[index]!;
    const valueNode = args[index + 1]!;
    const value = evaluate(valueNode, { ...context, locals });
    if (nameNode.kind === "name") locals.set(nameNode.name.toUpperCase(), value);
    index += 2;
  }
  const body = args[index];
  return body ? evaluate(body, { ...context, locals }) : ERRORS.value!;
}

function evaluateBinary(
  op: string,
  leftNode: Node,
  rightNode: Node,
  context: EvalContext
): EvalValue {
  if (op === ":" || op === "," || op === " ") {
    // Reference operators: evaluate through the resolver as a block.
    return ERRORS.value!;
  }
  const left = evaluate(leftNode, context);
  const right = evaluate(rightNode, context);

  const leftScalar = Array.isArray(left) ? (left[0]?.[0] ?? null) : left;
  const rightScalar = Array.isArray(right) ? (right[0]?.[0] ?? null) : right;
  if (isError(leftScalar)) return leftScalar;
  if (isError(rightScalar)) return rightScalar;

  if (op === "&") {
    const l = toText(leftScalar);
    const r = toText(rightScalar);
    if (isError(l)) return l;
    if (isError(r)) return r;
    return l + r;
  }

  if (["=", "<>", "<", "<=", ">", ">="].includes(op)) {
    const comparison = compareValues(leftScalar, rightScalar);
    if (comparison === null) return ERRORS.value!;
    switch (op) {
      case "=":
        return comparison === 0;
      case "<>":
        return comparison !== 0;
      case "<":
        return comparison < 0;
      case "<=":
        return comparison <= 0;
      case ">":
        return comparison > 0;
      case ">=":
        return comparison >= 0;
    }
  }

  const l = toNumber(leftScalar);
  const r = toNumber(rightScalar);
  if (isError(l)) return l;
  if (isError(r)) return r;
  switch (op) {
    case "+":
      return l + r;
    case "-":
      return l - r;
    case "*":
      return l * r;
    case "/":
      return r === 0 ? ERRORS.div0! : l / r;
    case "^":
      return l ** r;
    default:
      return ERRORS.value!;
  }
}

/** Evaluate a formula string in a context — convenience for tests. */
export function evaluateFormula(formula: string, context: EvalContext): EvalValue {
  const parse = parseFormula(formula);
  if (!parse.ok && parse.ast.kind === "bad") return ERRORS.name!;
  return evaluate(parse.ast, context);
}

export { NO_SHEET };

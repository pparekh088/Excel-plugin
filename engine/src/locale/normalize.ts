/**
 * Locale normalization (handoff §10 Phase 5, PLATFORM_QUIRKS Q-003).
 *
 * Everything internal — parser, graph, audit, LLM prompts, change sets —
 * speaks en-US. Office.js gives us that for free through `Range.formulas`;
 * `formulasLocal` is the user's locale and must NEVER be round-tripped.
 *
 * This module exists for the two places where locale still leaks in:
 *   1. a user typing a formula into our UI in their own locale
 *   2. displaying a formula back to them the way Excel shows it
 *
 * Both are conversions at the edge. Nothing downstream ever sees a localized
 * formula.
 */

import { parseFormula } from "../parser/parser";
import { toFormulaText } from "../parser/serialize";
import { Node, walk } from "../parser/ast";

export interface LocaleSpec {
  /** Argument separator: "," in en-US, ";" in most of Europe. */
  argumentSeparator: string;
  /** Decimal point: "." in en-US, "," in most of Europe. */
  decimalSeparator: string;
  /** Row separator inside array literals: ";" en-US, "\\" in some locales. */
  arrayRowSeparator: string;
  /** Localized function name -> canonical en-US name. */
  functionNames: Record<string, string>;
}

export const EN_US: LocaleSpec = {
  argumentSeparator: ",",
  decimalSeparator: ".",
  arrayRowSeparator: ";",
  functionNames: {},
};

/**
 * A representative subset of localized function names. The full table is
 * thousands of entries per locale; we ship the functions the corpus and the
 * audit rules actually reason about, and fall back to passing an unknown name
 * through unchanged rather than corrupting it.
 */
export const DE_DE: LocaleSpec = {
  argumentSeparator: ";",
  decimalSeparator: ",",
  arrayRowSeparator: "\\",
  functionNames: {
    SUMME: "SUM",
    MITTELWERT: "AVERAGE",
    ANZAHL: "COUNT",
    ANZAHL2: "COUNTA",
    WENN: "IF",
    WENNFEHLER: "IFERROR",
    UND: "AND",
    ODER: "OR",
    NICHT: "NOT",
    RUNDEN: "ROUND",
    MIN: "MIN",
    MAX: "MAX",
    SVERWEIS: "VLOOKUP",
    WVERWEIS: "HLOOKUP",
    INDEX: "INDEX",
    VERGLEICH: "MATCH",
    SUMMEWENN: "SUMIF",
    ZÄHLENWENN: "COUNTIF",
    HEUTE: "TODAY",
    JETZT: "NOW",
    SUMMENPRODUKT: "SUMPRODUCT",
  },
};

export const FR_FR: LocaleSpec = {
  argumentSeparator: ";",
  decimalSeparator: ",",
  arrayRowSeparator: "\\",
  functionNames: {
    SOMME: "SUM",
    MOYENNE: "AVERAGE",
    NB: "COUNT",
    NBVAL: "COUNTA",
    SI: "IF",
    SIERREUR: "IFERROR",
    ET: "AND",
    OU: "OR",
    NON: "NOT",
    ARRONDI: "ROUND",
    RECHERCHEV: "VLOOKUP",
    RECHERCHEH: "HLOOKUP",
    EQUIV: "MATCH",
    "SOMME.SI": "SUMIF",
    "NB.SI": "COUNTIF",
    AUJOURDHUI: "TODAY",
    MAINTENANT: "NOW",
    SOMMEPROD: "SUMPRODUCT",
  },
};

export const LOCALES: Record<string, LocaleSpec> = {
  "en-US": EN_US,
  "de-DE": DE_DE,
  "fr-FR": FR_FR,
};

/**
 * Convert a formula typed in `locale` into en-US.
 *
 * Tokenizing is done character-wise with string-literal awareness, because a
 * separator inside a quoted string is data, not syntax — replacing it would
 * silently corrupt the user's text.
 */
export function toEnUs(formula: string, locale: LocaleSpec): string {
  if (locale === EN_US) return formula;

  let out = "";
  let inString = false;
  let inQuotedSheet = false;
  let identifier = "";

  const flushIdentifier = (): void => {
    if (identifier === "") return;
    const canonical = locale.functionNames[identifier.toUpperCase()];
    out += canonical ?? identifier;
    identifier = "";
  };

  for (let index = 0; index < formula.length; index++) {
    const char = formula[index]!;

    if (inString) {
      out += char;
      if (char === '"') {
        if (formula[index + 1] === '"') {
          // Escaped quote: consume both and stay inside the string.
          out += '"';
          index++;
        } else {
          inString = false;
        }
      }
      continue;
    }
    if (inQuotedSheet) {
      out += char;
      if (char === "'") inQuotedSheet = false;
      continue;
    }
    if (char === '"') {
      flushIdentifier();
      inString = true;
      out += char;
      continue;
    }
    if (char === "'") {
      flushIdentifier();
      inQuotedSheet = true;
      out += char;
      continue;
    }

    // Identifier characters accumulate so function names translate whole.
    if (/[A-Za-z0-9_.À-ɏ]/.test(char)) {
      identifier += char;
      continue;
    }
    flushIdentifier();

    if (char === locale.argumentSeparator) {
      out += ",";
      continue;
    }
    if (char === locale.arrayRowSeparator && locale.arrayRowSeparator !== ";") {
      out += ";";
      continue;
    }
    if (char === locale.decimalSeparator && locale.decimalSeparator !== ".") {
      // Only a separator BETWEEN digits is a decimal point; elsewhere it is
      // an argument separator that happens to share the character.
      const previous = out[out.length - 1] ?? "";
      const next = formula[index + 1] ?? "";
      out += /[0-9]/.test(previous) && /[0-9]/.test(next) ? "." : char;
      continue;
    }
    out += char;
  }
  flushIdentifier();
  return out;
}

/** Convert an en-US formula for display in `locale`. Display only. */
export function fromEnUs(formula: string, locale: LocaleSpec): string {
  if (locale === EN_US) return formula;
  const reverse = new Map(
    Object.entries(locale.functionNames).map(([localized, canonical]) => [canonical, localized])
  );

  const parsed = parseFormula(formula);
  if (!parsed.ok) return formula; // never mangle what we do not understand

  const names: string[] = [];
  walk(parsed.ast, (node: Node) => {
    if (node.kind === "func") names.push(node.name);
  });

  let out = toFormulaText(parsed.ast);
  for (const name of new Set(names)) {
    const localized = reverse.get(name);
    if (!localized) continue;
    out = out.replace(new RegExp(`\\b${name}\\(`, "g"), `${localized}(`);
  }

  // Separators last, so function-name substitution cannot be confused by them.
  let result = "";
  let inString = false;
  for (let index = 0; index < out.length; index++) {
    const char = out[index]!;
    if (inString) {
      result += char;
      if (char === '"') {
        if (out[index + 1] === '"') {
          result += '"';
          index++;
        } else {
          inString = false;
        }
      }
      continue;
    }
    if (char === '"') {
      inString = true;
      result += char;
      continue;
    }
    if (char === ",") {
      result += locale.argumentSeparator;
      continue;
    }
    if (char === "." && /[0-9]/.test(out[index - 1] ?? "") && /[0-9]/.test(out[index + 1] ?? "")) {
      result += locale.decimalSeparator;
      continue;
    }
    result += char;
  }
  return result;
}

/** Detect whether a formula plausibly uses a non-en-US locale. */
export function looksLocalized(formula: string, locale: LocaleSpec): boolean {
  if (locale === EN_US) return false;
  const upper = formula.toUpperCase();
  return Object.keys(locale.functionNames).some((name) => upper.includes(`${name}(`));
}

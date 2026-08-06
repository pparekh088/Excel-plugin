/**
 * Formula tokenizer (en-US A1 grammar).
 *
 * Context-sensitive pieces (sheet prefixes, 3D spans, structured refs,
 * column/row ranges, spill #) are assembled in the parser; the tokenizer's
 * job is to classify spans without ever throwing. Whitespace is a real token
 * because ' ' is Excel's intersection operator.
 */

export type TokenType =
  | "number"
  | "string" // "..." with "" escape
  | "quoted" // '...' sheet-name quoting with '' escape
  | "error" // #REF! etc.
  | "ident" // names, function names, unqualified sheet names, bare columns
  | "cell" // $A$1 style (letters+digits)
  | "colRef" // $A / A$  ($-marked column-only piece)
  | "rowRef" // $1  ($-marked row-only piece)
  | "bracket" // [ ... ] block (structured ref part or external workbook prefix)
  | "op" // + - * / ^ & = <> <= >= < > % , ; : ( ) { } @ # !
  | "ws"
  | "unknown"
  | "eof";

export interface Token {
  type: TokenType;
  /** Raw source slice. */
  text: string;
  /** For string/quoted: unescaped value. For error: canonical error. */
  value?: string;
  start: number;
  end: number;
}

const ERROR_LITERALS = [
  "#GETTING_DATA",
  "#DIV/0!",
  "#VALUE!",
  "#NULL!",
  "#SPILL!",
  "#CALC!",
  "#BLOCKED!",
  "#CONNECT!",
  "#BUSY!",
  "#FIELD!",
  "#UNKNOWN!",
  "#EXTERNAL!",
  "#NAME?",
  "#NUM!",
  "#REF!",
  "#N/A",
];

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})/;
const NUMBER_RE = /^(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?/;

// Non-ASCII is allowed in names, but only real letters/marks/digits — not
// zero-width spaces, non-characters, or symbols, which would otherwise be
// silently absorbed into an identifier and hide a malformed formula.
const UNICODE_IDENT = /[\p{L}\p{M}\p{Nl}]/u;

function isIdentStart(ch: string): boolean {
  return /[A-Za-z_\\]/.test(ch) || (ch.charCodeAt(0) >= 0x00a0 && UNICODE_IDENT.test(ch));
}

function isIdentPart(ch: string): boolean {
  return (
    /[A-Za-z0-9_.?\\]/.test(ch) ||
    (ch.charCodeAt(0) >= 0x00a0 && (UNICODE_IDENT.test(ch) || /\p{Nd}/u.test(ch)))
  );
}

export function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = input.length;

  const push = (type: TokenType, start: number, end: number, value?: string) => {
    tokens.push({ type, text: input.slice(start, end), value, start, end });
  };

  while (i < n) {
    const ch = input[i]!;
    const start = i;

    // whitespace (space, tab, newline — all legal inside formulas)
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      while (i < n && /[ \t\n\r]/.test(input[i]!)) i++;
      push("ws", start, i);
      continue;
    }

    // string literal
    if (ch === '"') {
      i++;
      let value = "";
      let closed = false;
      while (i < n) {
        if (input[i] === '"') {
          if (input[i + 1] === '"') {
            value += '"';
            i += 2;
          } else {
            i++;
            closed = true;
            break;
          }
        } else {
          value += input[i];
          i++;
        }
      }
      push(closed ? "string" : "unknown", start, i, value);
      continue;
    }

    // quoted sheet name
    if (ch === "'") {
      i++;
      let value = "";
      let closed = false;
      while (i < n) {
        if (input[i] === "'") {
          if (input[i + 1] === "'") {
            value += "'";
            i += 2;
          } else {
            i++;
            closed = true;
            break;
          }
        } else {
          value += input[i];
          i++;
        }
      }
      push(closed ? "quoted" : "unknown", start, i, value);
      continue;
    }

    // error literal
    if (ch === "#") {
      const rest = input.slice(i);
      const hit = ERROR_LITERALS.find((e) =>
        rest.toUpperCase().startsWith(e)
      );
      if (hit) {
        i += hit.length;
        push("error", start, i, hit);
      } else {
        i++;
        push("op", start, i); // spill operator
      }
      continue;
    }

    // bracket block (structured ref part / external workbook prefix).
    // ' escapes the next char inside structured refs; nesting is tracked.
    if (ch === "[") {
      i++;
      let depth = 1;
      let inner = "";
      while (i < n && depth > 0) {
        const c = input[i]!;
        if (c === "'" && i + 1 < n) {
          inner += input[i + 1];
          i += 2;
          continue;
        }
        if (c === "[") depth++;
        if (c === "]") {
          depth--;
          if (depth === 0) {
            i++;
            break;
          }
        }
        inner += c;
        i++;
      }
      push(depth === 0 ? "bracket" : "unknown", start, i, inner);
      continue;
    }

    // number
    const numMatch = NUMBER_RE.exec(input.slice(i));
    if (numMatch && /[\d.]/.test(ch)) {
      i += numMatch[0].length;
      push("number", start, i);
      continue;
    }

    // $-prefixed cell / column / row pieces
    if (ch === "$") {
      const cellMatch = CELL_RE.exec(input.slice(i));
      if (cellMatch) {
        i += cellMatch[0].length;
        push("cell", start, i);
        continue;
      }
      const colMatch = /^\$[A-Za-z]{1,3}(?![A-Za-z0-9_.?\\])/.exec(input.slice(i));
      if (colMatch) {
        i += colMatch[0].length;
        push("colRef", start, i);
        continue;
      }
      const rowMatch = /^\$[1-9][0-9]{0,6}/.exec(input.slice(i));
      if (rowMatch) {
        i += rowMatch[0].length;
        push("rowRef", start, i);
        continue;
      }
      i++;
      push("unknown", start, i);
      continue;
    }

    // identifiers / cells / TRUE / FALSE
    if (isIdentStart(ch)) {
      const cellMatch = CELL_RE.exec(input.slice(i));
      // "A1" is a cell only if not followed by more ident chars ("A1B" is a name).
      if (cellMatch) {
        const after = input[i + cellMatch[0].length];
        if (after === undefined || !isIdentPart(after)) {
          // Mixed-abs middle form like A$1 is included by CELL_RE via groups.
          i += cellMatch[0].length;
          push("cell", start, i);
          continue;
        }
      }
      // A$1 with $ inside: letters then $digits
      const mixedMatch = /^([A-Za-z]{1,3})\$([1-9][0-9]{0,6})(?![0-9A-Za-z_.?\\])/.exec(
        input.slice(i)
      );
      if (mixedMatch) {
        i += mixedMatch[0].length;
        push("cell", start, i);
        continue;
      }
      while (i < n && isIdentPart(input[i]!)) i++;
      push("ident", start, i);
      continue;
    }

    // multi-char comparison operators
    const two = input.slice(i, i + 2);
    if (two === "<>" || two === "<=" || two === ">=") {
      i += 2;
      push("op", start, i);
      continue;
    }

    if ("+-*/^&=<>%,;:(){}@!".includes(ch)) {
      i++;
      push("op", start, i);
      continue;
    }

    i++;
    push("unknown", start, i);
  }

  push("eof", n, n);
  return tokens;
}

export function colLettersToIndex(letters: string): number {
  let value = 0;
  for (const ch of letters.toUpperCase()) {
    value = value * 26 + (ch.charCodeAt(0) - 64);
  }
  return value - 1;
}

export function colIndexToLetters(index: number): string {
  let value = index + 1;
  let out = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

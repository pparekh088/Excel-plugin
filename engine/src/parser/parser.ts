/**
 * Formula parser (en-US, A1). Pratt/precedence-climbing with Excel's operator
 * precedence (high → low):
 *
 *   reference ops  : (range)   " " (intersection)   , (union, inside groups)
 *   unary - +          (binds tighter than ^ : -2^2 = 4)
 *   % postfix
 *   ^ (left-assoc in Excel: 2^3^2 = 64)
 *   * /
 *   + -
 *   &
 *   = <> < <= > >=
 *
 * Never throws: malformed input produces `bad` nodes + diagnostics.
 */

import {
  BinaryOp,
  BoolLit,
  CellAddr,
  Diagnostic,
  Node,
  NO_SHEET,
  ParseResult,
  SheetSpec,
  StructuredItem,
  StructuredRefNode,
  isRefValued,
} from "./ast";
import { Token, colLettersToIndex, tokenize } from "./tokenizer";

const BINARY_PREC: Partial<Record<string, number>> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  "<=": 1,
  ">": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
};
const PREC_PERCENT = 6;
const PREC_UNARY = 7;
// Excel reference-operator precedence: ':' > ' ' (intersection) > ',' (union).
const PREC_ISECT = 8;
const PREC_COLON = 9;

const MAX_COL = 16_383; // XFD
const MAX_ROW = 1_048_575; // 1048576 rows, 0-based
const CELL_TOKEN_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/;

/**
 * Built-in functions whose names also match the A1 cell pattern. Everything
 * else that looks like a cell IS a cell: `A1(` is a syntax error, not a call.
 * (Custom functions are namespaced — AI.ASK — so they never collide.)
 */
const CELL_LIKE_FUNCTIONS = new Set(["LOG10", "ATAN2", "T3", "F4"]);

const STRUCTURED_ITEMS: Record<string, StructuredItem> = {
  "#ALL": "#All",
  "#DATA": "#Data",
  "#HEADERS": "#Headers",
  "#TOTALS": "#Totals",
  "#THIS ROW": "#This Row",
};

/** Internal-only pieces produced while folding X:Y column/row ranges. */
interface ColPiece {
  kind: "colPiece";
  sheet: SheetSpec;
  col: number;
  abs: boolean;
}
interface RowPiece {
  kind: "rowPiece";
  sheet: SheetSpec;
  row: number;
  abs: boolean;
}
type PNode = Node | ColPiece | RowPiece;

class Parser {
  private tokens: Token[];
  private pos = 0;
  private diagnostics: Diagnostic[] = [];

  constructor(private source: string) {
    this.tokens = tokenize(source);
  }

  // ------------------------------------------------------------- plumbing

  private raw(offset = 0): Token {
    return this.tokens[Math.min(this.pos + offset, this.tokens.length - 1)]!;
  }

  /** Next non-whitespace token (whitespace significance handled explicitly). */
  private peek(offset = 0): Token {
    let index = this.pos;
    let remaining = offset;
    for (;;) {
      const token = this.tokens[Math.min(index, this.tokens.length - 1)]!;
      if (token.type === "ws") {
        index++;
        continue;
      }
      if (remaining === 0) return token;
      remaining--;
      index++;
    }
  }

  private next(): Token {
    while (this.raw().type === "ws") this.pos++;
    const token = this.raw();
    if (token.type !== "eof") this.pos++;
    return token;
  }

  private isOp(token: Token, text: string): boolean {
    return token.type === "op" && token.text === text;
  }

  private diag(message: string, position: number): void {
    this.diagnostics.push({ message, position });
  }

  // ------------------------------------------------------------ entry

  parse(): ParseResult {
    const ast = this.parseExpr(0);
    const finalized = this.finalize(ast);
    const trailing = this.peek();
    if (trailing.type !== "eof") {
      this.diag(`unexpected trailing input '${this.source.slice(trailing.start)}'`, trailing.start);
    }
    return {
      source: this.source,
      ast: finalized,
      diagnostics: this.diagnostics,
      ok:
        this.diagnostics.length === 0 &&
        !containsBad(finalized) &&
        trailing.type === "eof",
    };
  }

  /** Convert stray col/row pieces (unfolded X: halves) into bad nodes. */
  private finalize(node: PNode): Node {
    if (node.kind === "colPiece" || node.kind === "rowPiece") {
      this.diag("dangling column/row reference piece", 0);
      return { kind: "bad", raw: "", message: "dangling column/row piece" };
    }
    return node;
  }

  // ------------------------------------------------------------ expressions

  private parseExpr(minPrec: number): PNode {
    let left = this.parseUnary();

    for (;;) {
      // Reference operators bind tightest and are handled first.
      const rawNext = this.raw();
      if (rawNext.type === "ws") {
        // Possible intersection: WS between two ref-valued operands.
        const after = this.peek();
        if (
          PREC_ISECT >= minPrec &&
          isRefStart(after) &&
          isPRefValued(left)
        ) {
          // Consume as intersection only if the right side really parses ref-valued;
          // otherwise fall through (plain spacing before an operator/operand).
          const save = this.pos;
          const savedDiags = this.diagnostics.length;
          this.pos++; // consume ws
          const right = this.parseExpr(PREC_ISECT + 1);
          if (isPRefValued(right)) {
            left = this.foldRef(" ", left, right, rawNext.start);
            continue;
          }
          this.pos = save;
          this.diagnostics.length = savedDiags;
        }
      }

      const token = this.peek();
      if (this.isOp(token, ":") && PREC_COLON >= minPrec) {
        this.next();
        const right = this.parseExpr(PREC_COLON + 1);
        left = this.foldRef(":", left, right, token.start);
        continue;
      }

      if (this.isOp(token, "%") && PREC_PERCENT >= minPrec) {
        this.next();
        left = { kind: "percent", operand: this.finalize(left) };
        continue;
      }

      if (token.type === "op") {
        const prec = BINARY_PREC[token.text];
        if (prec !== undefined && prec >= minPrec) {
          this.next();
          const right = this.parseExpr(prec + 1);
          left = {
            kind: "binary",
            op: token.text as BinaryOp,
            left: this.finalize(left),
            right: this.finalize(right),
          };
          continue;
        }
      }
      return left;
    }
  }

  private parseUnary(): PNode {
    const token = this.peek();
    if (this.isOp(token, "-") || this.isOp(token, "+")) {
      this.next();
      const operand = this.parseExpr(PREC_UNARY);
      return { kind: "unary", op: token.text as "-" | "+", operand: this.finalize(operand) };
    }
    if (this.isOp(token, "@")) {
      this.next();
      const operand = this.finalize(this.parseExpr(PREC_UNARY));
      if (operand.kind === "implicitIntersection") {
        this.diag("repeated '@' implicit-intersection operator", token.start);
      }
      return { kind: "implicitIntersection", operand };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): PNode {
    let node = this.parsePrimary();
    for (;;) {
      // Spill operator binds directly to the ref (no whitespace).
      if (this.raw().type === "op" && this.raw().text === "#" && isPRefValued(node)) {
        this.pos++;
        node = { kind: "spill", operand: this.finalize(node) };
        continue;
      }
      // Calling an expression: LAMBDA(..)(args), (ref)(args).
      if (
        this.rawIsImmediate("(") &&
        (node.kind === "func" || node.kind === "callExpr" || node.kind === "group")
      ) {
        node = { kind: "callExpr", callee: this.finalize(node), args: this.parseArgList() };
        continue;
      }
      return node;
    }
  }

  /** Parse "( arg, arg, ... )" — caller has verified '(' is next. */
  private parseArgList(): Node[] {
    this.pos++; // consume '('
    const args: Node[] = [];
    if (this.isOp(this.peek(), ")")) {
      this.next();
      return args;
    }
    for (;;) {
      if (this.isOp(this.peek(), ",")) {
        args.push({ kind: "missing" });
        this.next();
        continue;
      }
      if (this.isOp(this.peek(), ")")) {
        args.push({ kind: "missing" });
        this.next();
        break;
      }
      const arg = this.parseExpr(0);
      args.push(this.finalize(arg));
      const token = this.peek();
      if (this.isOp(token, ",")) {
        this.next();
        if (this.isOp(this.peek(), ")")) {
          args.push({ kind: "missing" });
          this.next();
          break;
        }
        continue;
      }
      if (this.isOp(token, ")")) {
        this.next();
        break;
      }
      this.diag(`expected ',' or ')' in argument list`, token.start);
      break;
    }
    return args;
  }

  // ------------------------------------------------------------ primaries

  private parsePrimary(): PNode {
    const token = this.peek();

    switch (token.type) {
      case "number": {
        this.next();
        return { kind: "number", value: Number(token.text) };
      }
      case "string": {
        this.next();
        return { kind: "string", value: token.value ?? "" };
      }
      case "error": {
        this.next();
        return { kind: "error", value: token.value ?? token.text };
      }
      case "cell":
        return this.parseCellish(NO_SHEET);
      case "colRef": {
        this.next();
        return {
          kind: "colPiece",
          sheet: NO_SHEET,
          col: colLettersToIndex(token.text.replace("$", "")),
          abs: token.text.startsWith("$"),
        };
      }
      case "rowRef": {
        this.next();
        return {
          kind: "rowPiece",
          sheet: NO_SHEET,
          row: Number(token.text.replace("$", "")) - 1,
          abs: true,
        };
      }
      case "quoted":
        return this.parseQuotedSheetRef();
      case "bracket":
        return this.parseBracketPrimary();
      case "ident":
        return this.parseIdentish(NO_SHEET);
      case "op":
        if (token.text === "(") return this.parseGroup();
        if (token.text === "{") return this.parseArray();
        break;
      default:
        break;
    }

    this.diag(`unexpected token '${token.text || token.type}'`, token.start);
    if (token.type !== "eof") this.next();
    return { kind: "bad", raw: token.text, message: "unexpected token" };
  }

  /** cell token, possibly a sheet name (A1!B2) or a LOG10(-style function. */
  private parseCellish(sheet: SheetSpec): PNode {
    const token = this.next();
    // Sheet named like a cell: A1!B2
    if (sheet === NO_SHEET && this.rawIsImmediate("!")) {
      this.pos++; // consume !
      return this.parseAfterSheet({ start: token.text, end: null, external: null });
    }
    // Table named like a cell: Q1[Amount]
    if (sheet === NO_SHEET && this.raw().type === "bracket") {
      const bracket = this.raw();
      this.pos++;
      return this.parseStructured(token.text, bracket);
    }
    if (this.rawIsImmediate("(")) {
      // Only a handful of built-ins collide with the cell pattern; anything
      // else followed by '(' is a malformed call on a reference.
      if (CELL_LIKE_FUNCTIONS.has(token.text.toUpperCase())) {
        return this.parseFuncCall(token.text, sheet);
      }
      this.diag(`'${token.text}' is a cell reference and cannot be called`, token.start);
      this.parseArgList();
      return { kind: "bad", raw: token.text, message: "call on cell reference" };
    }
    const addr = parseCellText(token.text);
    if (!addr) {
      this.diag(`invalid cell reference '${token.text}'`, token.start);
      return { kind: "bad", raw: token.text, message: "invalid cell" };
    }
    // Beyond the grid (XFE1, A1048577): Excel treats these as defined names.
    if (addr.col > MAX_COL || addr.row > MAX_ROW) {
      return { kind: "name", sheet, name: token.text.replace(/\$/g, "") };
    }
    return { kind: "cell", sheet, addr };
  }

  private rawIsImmediate(text: string): boolean {
    const token = this.raw();
    return token.type === "op" && token.text === text;
  }

  private parseIdentish(sheet: SheetSpec): PNode {
    const token = this.next();
    const upper = token.text.toUpperCase();

    if (this.rawIsImmediate("(")) {
      return this.parseFuncCall(token.text, sheet);
    }
    // Structured ref: Table[ ... ] — bracket must be immediately adjacent.
    if (sheet === NO_SHEET && this.raw().type === "bracket") {
      const bracket = this.raw();
      this.pos++;
      return this.parseStructured(token.text, bracket);
    }
    // Sheet prefix: Name! or 3D span Name1:Name2!
    if (sheet === NO_SHEET && this.rawIsImmediate("!")) {
      this.pos++;
      return this.parseAfterSheet({ start: token.text, end: null, external: null });
    }
    if (
      sheet === NO_SHEET &&
      this.rawIsImmediate(":") &&
      this.raw(1).type === "ident" &&
      this.tokens[this.pos + 2] &&
      this.isOp(this.tokens[this.pos + 2]!, "!")
    ) {
      const endSheet = this.raw(1).text;
      this.pos += 3;
      return this.parseAfterSheet({ start: token.text, end: endSheet, external: null });
    }

    if (sheet === NO_SHEET && upper === "TRUE") return { kind: "bool", value: true } as BoolLit;
    if (sheet === NO_SHEET && upper === "FALSE") return { kind: "bool", value: false } as BoolLit;

    return { kind: "name", sheet, name: token.text };
  }

  /** After consuming `<sheet>!` — parse the referenced thing. */
  private parseAfterSheet(sheet: SheetSpec): PNode {
    const token = this.peek();
    switch (token.type) {
      case "cell": {
        const node = this.parseCellish(sheet);
        return node;
      }
      case "colRef": {
        this.next();
        return {
          kind: "colPiece",
          sheet,
          col: colLettersToIndex(token.text.replace("$", "")),
          abs: token.text.startsWith("$"),
        };
      }
      case "rowRef": {
        this.next();
        return {
          kind: "rowPiece",
          sheet,
          row: Number(token.text.replace("$", "")) - 1,
          abs: true,
        };
      }
      case "ident":
        return this.parseIdentish(sheet);
      case "number": {
        // Sheet1!1:3 (row range) — leave folding to ':' handler.
        this.next();
        return { kind: "rowPiece", sheet, row: Number(token.text) - 1, abs: false };
      }
      case "error": {
        this.next();
        this.diag(`broken reference ${sheet.start ?? ""}!${token.text}`, token.start);
        return { kind: "error", value: token.value ?? token.text };
      }
      default:
        this.diag(`expected reference after sheet name`, token.start);
        return { kind: "bad", raw: token.text, message: "expected ref after sheet" };
    }
  }

  private parseQuotedSheetRef(): PNode {
    const token = this.next();
    let name = token.value ?? "";
    let external: string | null = null;
    const extMatch = /^\[([^\]]+)\]/.exec(name);
    if (extMatch) {
      external = extMatch[1] ?? null;
      name = name.slice(extMatch[0].length);
    }
    let end: string | null = null;
    const spanIndex = name.indexOf(":");
    if (spanIndex > 0) {
      end = name.slice(spanIndex + 1);
      name = name.slice(0, spanIndex);
    }
    if (!this.rawIsImmediate("!")) {
      this.diag(`quoted name '${token.text}' must be followed by '!'`, token.end);
      return { kind: "bad", raw: token.text, message: "quoted sheet without !" };
    }
    this.pos++;
    return this.parseAfterSheet({ start: name, end, external });
  }

  /** Bare bracket: external prefix ([Book1.xlsx]Sheet1!A1) or structured ([@Col]). */
  private parseBracketPrimary(): PNode {
    const token = this.next();
    const after = this.raw();
    const isExternalPrefix =
      (after.type === "ident" || after.type === "cell" || after.type === "quoted") &&
      this.hasBangAhead();
    if (isExternalPrefix) {
      const sheetToken = this.next();
      let sheetName = sheetToken.type === "quoted" ? sheetToken.value ?? "" : sheetToken.text;
      let end: string | null = null;
      if (this.rawIsImmediate(":")) {
        this.pos++;
        const endToken = this.next();
        end = endToken.type === "quoted" ? endToken.value ?? "" : endToken.text;
      }
      if (this.rawIsImmediate("!")) {
        this.pos++;
        return this.parseAfterSheet({ start: sheetName, end, external: token.value ?? "" });
      }
      this.diag("expected '!' after external sheet reference", this.raw().start);
      return { kind: "bad", raw: token.text, message: "malformed external ref" };
    }
    return this.parseStructured("", token);
  }

  private hasBangAhead(): boolean {
    // lookahead over at most: ident (:ident) !
    let index = this.pos;
    const t = (k: number) => this.tokens[Math.min(index + k, this.tokens.length - 1)]!;
    if (!(t(0).type === "ident" || t(0).type === "cell" || t(0).type === "quoted")) return false;
    if (this.isOp(t(1), "!")) return true;
    if (this.isOp(t(1), ":") && (t(2).type === "ident" || t(2).type === "quoted"))
      return this.isOp(t(3), "!");
    return false;
  }

  /**
   * Structured reference: Table[...]. Parsed from the RAW bracket text so
   * that ' escapes stay visible — Table['#Col] selects a column literally
   * named "#Col", which is a different thing from the [#Totals] item
   * specifier. Unescaping only happens once a segment's role is decided.
   */
  private parseStructured(table: string, bracket: Token): StructuredRefNode {
    const node: StructuredRefNode = {
      kind: "structured",
      table,
      items: [],
      columns: [],
      thisRow: false,
      raw: bracket.text.length > 0 ? `${table}${bracket.text}` : `${table}[]`,
    };

    // Char stream with escape flags; outer [ ] stripped.
    const chars = escapeAwareChars(bracket.text.replace(/^\[/, "").replace(/\]$/, ""));

    let index = 0;
    while (index < chars.length && chars[index]!.ch === " " && !chars[index]!.escaped) index++;
    if (index < chars.length && chars[index]!.ch === "@" && !chars[index]!.escaped) {
      node.thisRow = true;
      index++;
    }

    const parts: EscapedChar[][] = [[]];
    let depth = 0;
    for (; index < chars.length; index++) {
      const entry = chars[index]!;
      if (!entry.escaped) {
        if (entry.ch === "[") depth++;
        else if (entry.ch === "]") depth--;
        else if (entry.ch === "," && depth === 0) {
          parts.push([]);
          continue;
        }
      }
      parts[parts.length - 1]!.push(entry);
    }

    for (const partChars of parts) {
      const part = trimChars(partChars);
      if (part.length === 0) continue;

      // Column span: [a]:[b]
      const span = splitSpan(part);
      if (span) {
        node.columns.push(unescapeChars(span[0]), unescapeChars(span[1]));
        continue;
      }

      const inner = unwrapBrackets(part);
      if (inner === null) {
        this.diag(`malformed structured reference '${bracket.text}'`, bracket.start);
        continue;
      }
      const trimmed = trimChars(inner);
      const text = unescapeChars(trimmed);

      // An UNESCAPED leading '#' marks an item specifier.
      if (trimmed.length > 0 && trimmed[0]!.ch === "#" && !trimmed[0]!.escaped) {
        const item = STRUCTURED_ITEMS[text.toUpperCase()];
        if (item) {
          node.items.push(item);
          if (item === "#This Row") node.thisRow = true;
        } else {
          this.diag(`unknown structured-reference item '${text}'`, bracket.start);
        }
        continue;
      }
      if (text === "") {
        this.diag("empty column specifier in structured reference", bracket.start);
        continue;
      }
      if (trimmed.some((c) => !c.escaped && (c.ch === "[" || c.ch === "]"))) {
        this.diag(`malformed structured reference '${bracket.text}'`, bracket.start);
        continue;
      }
      node.columns.push(text);
    }

    if (node.items.length === 0 && node.columns.length === 0 && !node.thisRow) {
      this.diag(`structured reference selects nothing: '${bracket.text}'`, bracket.start);
    }
    return node;
  }

  private parseFuncCall(rawName: string, sheet: SheetSpec): PNode {
    const args = this.parseArgList();
    let canonical = rawName.toUpperCase();
    if (canonical.startsWith("_XLFN.")) canonical = canonical.slice("_XLFN.".length);
    if (sheet !== NO_SHEET) {
      this.diag("sheet-qualified function call", 0);
    }
    // Arity rules for the two functions whose shape we depend on for scoping.
    // LET(name, value, ..., body) needs an odd count >= 3; LAMBDA needs a body.
    if (canonical === "LET" && (args.length < 3 || args.length % 2 === 0)) {
      this.diag(`LET expects name/value pairs then a body, got ${args.length} arguments`, 0);
    }
    if (canonical === "LAMBDA" && args.length < 1) {
      this.diag("LAMBDA requires a body expression", 0);
    }
    return { kind: "func", name: canonical, rawName, args };
  }

  private parseGroup(): PNode {
    this.next(); // consume '('
    let expr = this.parseExpr(0);
    // Union: (A1,B2,...) — commas at group top level over ref operands.
    while (this.isOp(this.peek(), ",")) {
      const comma = this.next();
      const right = this.parseExpr(0);
      if (!isPRefValued(expr) || !isPRefValued(right)) {
        this.diag("',' inside parentheses is only valid between references", comma.start);
      }
      expr = {
        kind: "binary",
        op: ",",
        left: this.finalize(expr),
        right: this.finalize(right),
      };
    }
    if (this.isOp(this.peek(), ")")) {
      this.next();
    } else {
      this.diag("missing closing ')'", this.peek().start);
    }
    return { kind: "group", expr: this.finalize(expr) };
  }

  private parseArray(): PNode {
    const open = this.next(); // consume '{'
    const rows: Node[][] = [[]];
    for (;;) {
      const token = this.peek();
      if (token.type === "eof") {
        this.diag("unterminated array literal", token.start);
        break;
      }
      if (this.isOp(token, "}")) {
        this.next();
        break;
      }
      if (this.isOp(token, ",")) {
        this.next();
        continue;
      }
      if (this.isOp(token, ";")) {
        this.next();
        rows.push([]);
        continue;
      }
      const element = this.parseArrayElement();
      rows[rows.length - 1]!.push(element);
    }
    // Excel array literals are rectangular and non-empty.
    const width = rows[0]?.length ?? 0;
    if (width === 0 || rows.some((row) => row.length !== width)) {
      this.diag("array literal rows must be non-empty and equal length", open.start);
    }
    return { kind: "array", rows };
  }

  private parseArrayElement(): Node {
    const token = this.peek();
    if (this.isOp(token, "-") || this.isOp(token, "+")) {
      this.next();
      const operand = this.parseArrayElement();
      return { kind: "unary", op: token.text as "-" | "+", operand };
    }
    if (token.type === "number") {
      this.next();
      return { kind: "number", value: Number(token.text) };
    }
    if (token.type === "string") {
      this.next();
      return { kind: "string", value: token.value ?? "" };
    }
    if (token.type === "error") {
      this.next();
      return { kind: "error", value: token.value ?? token.text };
    }
    if (token.type === "ident" && ["TRUE", "FALSE"].includes(token.text.toUpperCase())) {
      this.next();
      return { kind: "bool", value: token.text.toUpperCase() === "TRUE" };
    }
    this.diag(`invalid array element '${token.text}'`, token.start);
    this.next();
    return { kind: "bad", raw: token.text, message: "invalid array element" };
  }

  // ------------------------------------------------------- ref folding (:)

  private foldRef(op: ":" | " ", left: PNode, right: PNode, position: number): PNode {
    if (op === ":") {
      // cell:cell → rectangular range
      if (left.kind === "cell" && right.kind === "cell") {
        if (right.sheet.start === null || sameSheet(left.sheet, right.sheet)) {
          return {
            kind: "range",
            sheet: left.sheet,
            start: left.addr,
            end: right.addr,
          };
        }
        this.diag("range endpoints on different sheets", position);
      }
      // column ranges: colPiece:colPiece, or short names (A:B, TAX:TAX)
      const leftCol = asColPiece(left);
      const rightCol = asColPiece(right);
      if (leftCol && rightCol) {
        return {
          kind: "colRange",
          sheet: leftCol.sheet,
          startCol: leftCol.col,
          endCol: rightCol.col,
          startAbs: leftCol.abs,
          endAbs: rightCol.abs,
        };
      }
      // row ranges: rowPiece/int:rowPiece/int
      const leftRow = asRowPiece(left);
      const rightRow = asRowPiece(right);
      if (leftRow && rightRow) {
        return {
          kind: "rowRange",
          sheet: leftRow.sheet,
          startRow: leftRow.row,
          endRow: rightRow.row,
          startAbs: leftRow.abs,
          endAbs: rightRow.abs,
        };
      }
    }
    // General reference operator (INDEX(..):B5, A1:A10 B2:B10, unions…)
    return {
      kind: "binary",
      op,
      left: this.finalize(left),
      right: this.finalize(right),
    };
  }
}

// -------------------------------------------- structured-ref char helpers

interface EscapedChar {
  ch: string;
  escaped: boolean;
}

/** Split text into chars, resolving ' escapes but remembering they occurred. */
function escapeAwareChars(text: string): EscapedChar[] {
  const out: EscapedChar[] = [];
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "'" && i + 1 < text.length) {
      out.push({ ch: text[i + 1]!, escaped: true });
      i++;
      continue;
    }
    out.push({ ch: text[i]!, escaped: false });
  }
  return out;
}

function trimChars(chars: EscapedChar[]): EscapedChar[] {
  let start = 0;
  let end = chars.length;
  while (start < end && chars[start]!.ch === " " && !chars[start]!.escaped) start++;
  while (end > start && chars[end - 1]!.ch === " " && !chars[end - 1]!.escaped) end--;
  return chars.slice(start, end);
}

function unescapeChars(chars: EscapedChar[]): string {
  return chars.map((c) => c.ch).join("");
}

/** Strip one level of unescaped [ ] if present; null if unbalanced. */
function unwrapBrackets(chars: EscapedChar[]): EscapedChar[] | null {
  if (chars.length === 0) return chars;
  const first = chars[0]!;
  const last = chars[chars.length - 1]!;
  const opensBracket = first.ch === "[" && !first.escaped;
  const closesBracket = last.ch === "]" && !last.escaped;
  if (opensBracket && closesBracket && chars.length >= 2) return chars.slice(1, -1);
  if (opensBracket !== closesBracket) return null;
  return chars;
}

/** Match `[a]:[b]` over unescaped delimiters; returns the two inner spans. */
function splitSpan(chars: EscapedChar[]): [EscapedChar[], EscapedChar[]] | null {
  if (chars.length === 0 || chars[0]!.ch !== "[" || chars[0]!.escaped) return null;
  let depth = 0;
  for (let i = 0; i < chars.length; i++) {
    const entry = chars[i]!;
    if (entry.escaped) continue;
    if (entry.ch === "[") depth++;
    else if (entry.ch === "]") {
      depth--;
      if (depth === 0) {
        const rest = trimChars(chars.slice(i + 1));
        if (rest.length === 0 || rest[0]!.ch !== ":" || rest[0]!.escaped) return null;
        const right = unwrapBrackets(trimChars(rest.slice(1)));
        if (right === null) return null;
        return [chars.slice(1, i), right];
      }
    }
  }
  return null;
}

// ------------------------------------------------------------------ helpers

function containsBad(node: Node): boolean {
  let bad = false;
  const visit = (n: Node): void => {
    if (bad) return;
    if (n.kind === "bad") {
      bad = true;
      return;
    }
    switch (n.kind) {
      case "unary":
      case "percent":
      case "implicitIntersection":
      case "spill":
        visit(n.operand);
        break;
      case "binary":
        visit(n.left);
        visit(n.right);
        break;
      case "func":
        n.args.forEach(visit);
        break;
      case "group":
        visit(n.expr);
        break;
      case "array":
        n.rows.forEach((row) => row.forEach(visit));
        break;
      default:
        break;
    }
  };
  visit(node);
  return bad;
}

function isPRefValued(node: PNode): boolean {
  if (node.kind === "colPiece" || node.kind === "rowPiece") return true;
  return isRefValued(node);
}

function isRefStart(token: Token): boolean {
  return (
    token.type === "cell" ||
    token.type === "colRef" ||
    token.type === "rowRef" ||
    token.type === "ident" ||
    token.type === "quoted" ||
    token.type === "bracket" ||
    token.type === "number" || // row-range piece: A1:C10 2:2 (backtracks if not)
    (token.type === "op" && (token.text === "(" || token.text === "@"))
  );
}

function sameSheet(a: SheetSpec, b: SheetSpec): boolean {
  return (
    (a.start ?? "").toUpperCase() === (b.start ?? "").toUpperCase() &&
    (a.end ?? "") === (b.end ?? "") &&
    (a.external ?? "") === (b.external ?? "")
  );
}

function asColPiece(node: PNode): ColPiece | null {
  if (node.kind === "colPiece") return node;
  if (node.kind === "name" && /^[A-Za-z]{1,3}$/.test(node.name)) {
    const col = colLettersToIndex(node.name);
    if (col <= MAX_COL) {
      return { kind: "colPiece", sheet: node.sheet, col, abs: false };
    }
  }
  return null;
}

function asRowPiece(node: PNode): RowPiece | null {
  if (node.kind === "rowPiece") return node;
  if (node.kind === "number" && Number.isInteger(node.value) && node.value >= 1) {
    return { kind: "rowPiece", sheet: NO_SHEET, row: node.value - 1, abs: false };
  }
  return null;
}

export function parseCellText(text: string): CellAddr | null {
  const match = CELL_TOKEN_RE.exec(text);
  if (!match) return null;
  return {
    col: colLettersToIndex(match[2]!),
    row: Number(match[4]!) - 1,
    colAbs: match[1] === "$",
    rowAbs: match[3] === "$",
  };
}

/**
 * Parse a formula. Accepts with or without the leading '='; array-entered
 * formulas' braces ({=...}) are stripped.
 */
export function parseFormula(formula: string): ParseResult {
  let body = formula;
  let prefixLength = 0;
  if (body.startsWith("{=") && body.endsWith("}")) {
    body = body.slice(2, -1);
    prefixLength = 2;
  } else if (body.startsWith("=")) {
    body = body.slice(1);
    prefixLength = 1;
  }
  if (body.trim() === "") {
    return {
      source: formula,
      ast: { kind: "bad", raw: formula, message: "empty formula" },
      diagnostics: [{ message: "empty formula", position: prefixLength }],
      ok: false,
    };
  }
  const result = new Parser(body).parse();
  return { ...result, source: formula };
}

/** Quick check used by extraction: is this cell text a formula? */
export function isFormulaText(value: unknown): value is string {
  return typeof value === "string" && (value.startsWith("=") || value.startsWith("{="));
}

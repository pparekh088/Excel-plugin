import { describe, expect, it } from "vitest";
import {
  colIndexToLetters,
  colLettersToIndex,
  tokenize,
} from "../../src/parser/tokenizer";

function types(input: string): string[] {
  return tokenize(input)
    .filter((t) => t.type !== "eof")
    .map((t) => `${t.type}:${t.text}`);
}

describe("tokenizer", () => {
  it.each([
    ["1+2", ["number:1", "op:+", "number:2"]],
    ["1.5e-3", ["number:1.5e-3"]],
    [".5", ["number:.5"]],
    ['"a""b"', ['string:"a""b"']],
    ["A1", ["cell:A1"]],
    ["$A$1", ["cell:$A$1"]],
    ["A$1", ["cell:A$1"]],
    ["$A1", ["cell:$A1"]],
    ["XFD1048576", ["cell:XFD1048576"]],
    ["A1B", ["ident:A1B"]],
    ["ABCD1", ["ident:ABCD1"]],
    ["$A", ["colRef:$A"]],
    ["$1", ["rowRef:$1"]],
    ["SUM(", ["ident:SUM", "op:("]],
    ["Sheet1!A1", ["ident:Sheet1", "op:!", "cell:A1"]],
    ["'My Sheet'!A1", ["quoted:'My Sheet'", "op:!", "cell:A1"]],
    ["'It''s'!B2", ["quoted:'It''s'", "op:!", "cell:B2"]],
    ["#REF!", ["error:#REF!"]],
    ["#DIV/0!", ["error:#DIV/0!"]],
    ["#N/A", ["error:#N/A"]],
    ["#NAME?", ["error:#NAME?"]],
    ["#GETTING_DATA", ["error:#GETTING_DATA"]],
    ["A1#", ["cell:A1", "op:#"]],
    ["@A1", ["op:@", "cell:A1"]],
    ["<>", ["op:<>"]],
    ["<=", ["op:<="]],
    [">=", ["op:>="]],
    ["a<b", ["ident:a", "op:<", "ident:b"]],
    ["{1,2;3}", ["op:{", "number:1", "op:,", "number:2", "op:;", "number:3", "op:}"]],
    ["Table1[Amount]", ["ident:Table1", "bracket:[Amount]"]],
    ["Table1[[#Headers],[Amt]]", ["ident:Table1", "bracket:[[#Headers],[Amt]]"]],
    ["[@Col]", ["bracket:[@Col]"]],
    // S1 lexes as a cell; the parser resolves it as a sheet name via the '!'.
    ["[Book.xlsx]S1!A1", ["bracket:[Book.xlsx]", "cell:S1", "op:!", "cell:A1"]],
    ["[Book.xlsx]Sheet1!A1", ["bracket:[Book.xlsx]", "ident:Sheet1", "op:!", "cell:A1"]],
    ["T['#x]", ["ident:T", "bracket:['#x]"]],
    ["a b", ["ident:a", "ws: ", "ident:b"]],
    ["1 \t\n 2", ["number:1", "ws: \t\n ", "number:2"]],
    ["Umsätze", ["ident:Umsätze"]],
    ["收入", ["ident:收入"]],
    ["A1:B2", ["cell:A1", "op::", "cell:B2"]],
    ["5%", ["number:5", "op:%"]],
  ] as const)("tokenizes %s", (input, expected) => {
    expect(types(input)).toEqual(expected as unknown as string[]);
  });

  it("unterminated string/quote/bracket yields unknown token, never throws", () => {
    expect(tokenize('"abc')[0]!.type).toBe("unknown");
    expect(tokenize("'abc")[0]!.type).toBe("unknown");
    expect(tokenize("[abc")[0]!.type).toBe("unknown");
  });

  it("column letters round-trip across the full Excel width", () => {
    expect(colLettersToIndex("A")).toBe(0);
    expect(colLettersToIndex("Z")).toBe(25);
    expect(colLettersToIndex("AA")).toBe(26);
    expect(colLettersToIndex("XFD")).toBe(16383);
    for (let col = 0; col <= 16383; col++) {
      expect(colLettersToIndex(colIndexToLetters(col))).toBe(col);
    }
  });
});

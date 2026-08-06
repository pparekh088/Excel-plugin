import { describe, expect, it } from "vitest";
import { parseFormula } from "../../src/parser/parser";
import { normalizeR1C1 } from "../../src/parser/r1c1";

function norm(formula: string, row: number, col: number): string {
  const result = parseFormula(formula);
  expect(result.ok, `parse ${formula}`).toBe(true);
  return normalizeR1C1(result.ast, row, col);
}

describe("R1C1 normalization", () => {
  it.each([
    // [formula, hostRow, hostCol, expected]
    ["=B2", 2, 2, "R[-1]C[-1]"],
    ["=B2", 1, 1, "RC"],
    ["=$B$2", 10, 10, "R2C2"],
    ["=B$2", 5, 3, "R2C[-2]"],
    ["=$B2", 5, 3, "R[-4]C2"],
    ["=A5", 4, 1, "RC[-1]"],
    ["=B4", 4, 1, "R[-1]C"],
    ["=D10*2", 9, 2, "RC[1]*2"],
    ["=SUM(A1:A10)", 10, 2, "SUM(R[-10]C[-2]:R[-1]C[-2])"],
    ["=SUM($A$1:$A$10)", 10, 2, "SUM(R1C1:R10C1)"],
    ["=Sheet2!A1", 0, 1, "Sheet2!RC[-1]"],
    ["='My Sheet'!A1", 0, 1, "'My Sheet'!RC[-1]"],
    ["=SUM(A:A)", 0, 2, "SUM(C[-2]:C[-2])"],
    ["=SUM($A:$B)", 0, 5, "SUM(C1:C2)"],
    ["=SUM(1:1)", 3, 0, "SUM(R[-3]:R[-3])"],
    ["=SUM($1:$2)", 3, 0, "SUM(R1:R2)"],
    ["=revenue*Tax_rate", 0, 0, "REVENUE*TAX_RATE"],
    ["=IF(A1>0,B1,C1)", 0, 3, "IF(RC[-3]>0,RC[-2],RC[-1])"],
    ['=TEXT(A1,"0.00")', 0, 1, 'TEXT(RC[-1],"0.00")'],
    ["=Table1[@Amount]", 5, 5, "TABLE1[@AMOUNT]"],
    ["=A1#", 1, 1, "R[-1]C[-1]#"],
    ["=@A1:A10", 0, 1, "@RC[-1]:R[9]C[-1]"],
    ["=-2^2", 0, 0, "-2^2"],
    ["=5%", 0, 0, "5%"],
    ["={1,2;3,4}", 0, 0, "{1,2;3,4}"],
    ["=SUM(Sheet1:Sheet3!B2)", 1, 1, "SUM(Sheet1:Sheet3!RC)"],
    ["=LET(x,A1,x*2)", 0, 1, "LET(X,RC[-1],X*2)"],
  ] as const)("normalizes %s at (%i,%i)", (formula, row, col, expected) => {
    expect(norm(formula, row, col)).toBe(expected);
  });

  describe("fill-pattern equivalence (run-detection property)", () => {
    const copies: Array<[string, number, number, string, number, number]> = [
      // same formula filled down: =A1+B1 in C1, =A2+B2 in C2
      ["=A1+B1", 0, 2, "=A2+B2", 1, 2],
      ["=SUM(A1:A10)", 0, 3, "=SUM(B1:B10)", 0, 4],
      ["=$A$1*B2", 1, 2, "=$A$1*B3", 2, 2],
      ["=Sheet2!A1", 0, 1, "=Sheet2!A5", 4, 1],
      ["=A$1*$B2", 1, 3, "=B$1*$B3", 2, 4],
      ["=VLOOKUP(A2,Data!$A:$D,3,FALSE)", 1, 5, "=VLOOKUP(A9,Data!$A:$D,3,FALSE)", 8, 5],
      ["=B2/B$10", 1, 1, "=B7/B$10", 6, 1],
      ["=IF(C3>0,C3*Rate,0)", 2, 4, "=IF(C8>0,C8*Rate,0)", 7, 4],
    ];
    it.each(copies)(
      "%s@(%i,%i) matches %s@(%i,%i)",
      (f1, r1, c1, f2, r2, c2) => {
        expect(norm(f1, r1, c1)).toBe(norm(f2, r2, c2));
      }
    );

    const breaks: Array<[string, number, number, string, number, number]> = [
      // hardcoded override breaks the pattern
      ["=A1+B1", 0, 2, "=A2+100", 1, 2],
      // absolute vs relative differ
      ["=$A$1*B2", 1, 2, "=A1*B3", 2, 2],
      // shifted ref that is NOT the fill pattern
      ["=B2/B$10", 1, 1, "=B7/B$11", 6, 1],
      // different function
      ["=SUM(A1:A10)", 0, 3, "=AVERAGE(B1:B10)", 0, 4],
    ];
    it.each(breaks)(
      "%s@(%i,%i) differs from %s@(%i,%i)",
      (f1, r1, c1, f2, r2, c2) => {
        expect(norm(f1, r1, c1)).not.toBe(norm(f2, r2, c2));
      }
    );
  });
});

import { describe, expect, it } from "vitest";
import { ERRORS, EvalContext, evaluateFormula } from "../../src/sim/evaluator";
import { Simulator } from "../../src/sim/simulator";
import { Workbook } from "../../src/model/workbook";
import { workbookOf } from "../helpers/build";
import { threeStatementModel, dcfModel, budgetVsActual } from "../../src/corpus/models";

function contextFor(workbook: Workbook, sheet = "S", row = 0, col = 0): EvalContext {
  const simulator = Simulator.of(workbook);
  return {
    workbook,
    sheet,
    row,
    col,
    read: (s, r, c) => simulator.read(s, r, c),
    locals: new Map(),
  };
}

function evalIn(cells: Record<string, string | number | boolean | null>, formula: string) {
  const workbook = workbookOf({ S: cells });
  return evaluateFormula(formula, contextFor(workbook, "S", 20, 20));
}

describe("evaluator arithmetic and coercion", () => {
  it.each([
    ["=1+2", 3],
    ["=10-4", 6],
    ["=3*4", 12],
    ["=10/4", 2.5],
    ["=2^10", 1024],
    ["=-2^2", 4], // unary binds tighter
    ["=2^3^2", 64], // left associative
    ["=1+2*3", 7],
    ["=(1+2)*3", 9],
    ["=50%", 0.5],
    ["=10%*200", 20],
    ['="a"&"b"', "ab"],
    ['="n="&5', "n=5"],
    ["=1=1", true],
    ["=1<>2", true],
    ["=2<1", false],
    ["=1<=1", true],
    ['="a"="A"', true], // Excel comparison is case-insensitive
    ["=1/0", ERRORS.div0],
    ['="x"+1', ERRORS.value],
    ["=TRUE", true],
    ["=NOT(TRUE)", false],
    ["=AND(TRUE,FALSE)", false],
    ["=OR(FALSE,TRUE)", true],
  ] as const)("evaluates %s", (formula, expected) => {
    expect(evalIn({}, formula)).toBe(expected);
  });

  it("treats blank cells as zero in arithmetic", () => {
    expect(evalIn({ A1: 5 }, "=A1+B1")).toBe(5);
  });

  it("propagates errors through arithmetic", () => {
    expect(evalIn({ A1: "#REF!" }, "=A1+1")).toBe("#REF!");
  });
});

describe("evaluator functions", () => {
  const data = { A1: 10, A2: 20, A3: 30, B1: "x", B2: "y", B3: "x", C1: 1, C2: 2, C3: 3 };

  it.each([
    ["=SUM(A1:A3)", 60],
    ["=AVERAGE(A1:A3)", 20],
    ["=MIN(A1:A3)", 10],
    ["=MAX(A1:A3)", 30],
    ["=COUNT(A1:B3)", 3],
    ["=COUNTA(A1:B3)", 6],
    ["=MEDIAN(A1:A3)", 20],
    ["=ROUND(2.567,2)", 2.57],
    ["=ROUND(-2.567,2)", -2.57],
    ["=ROUNDUP(2.001,2)", 2.01],
    ["=ROUNDDOWN(2.999,2)", 2.99],
    ["=ABS(-7)", 7],
    ["=SQRT(16)", 4],
    ["=POWER(2,8)", 256],
    ["=MOD(7,3)", 1],
    ["=MOD(-7,3)", 2], // Excel MOD follows the divisor's sign
    ["=INT(2.9)", 2],
    ["=SIGN(-3)", -1],
    ['=IF(A1>5,"big","small")', "big"],
    ["=IFERROR(1/0,-1)", -1],
    ["=IFERROR(5,-1)", 5],
    ['=CONCATENATE("a",1,TRUE)', "a1TRUE"],
    ['=LEFT("hello",2)', "he"],
    ['=RIGHT("hello",3)', "llo"],
    ['=MID("hello",2,3)', "ell"],
    ['=LEN("hello")', 5],
    ['=UPPER("ab")', "AB"],
    ['=TRIM("  a  b  ")', "a b"],
    ['=VALUE("42")', 42],
    ["=COUNTIF(B1:B3,\"x\")", 2],
    ['=SUMIF(A1:A3,">15")', 50],
    ['=SUMIF(B1:B3,"x",A1:A3)', 40],
    ['=SUMIFS(A1:A3,B1:B3,"x")', 40],
    ['=COUNTIFS(B1:B3,"x")', 2],
    ["=SUMPRODUCT(A1:A3,C1:C3)", 140],
    ["=INDEX(A1:A3,2)", 20],
    ["=MATCH(20,A1:A3,0)", 2],
    ['=MATCH("y",B1:B3,0)', 2],
    ["=SMALL(A1:A3,1)", 10],
    ["=LARGE(A1:A3,1)", 30],
    ["=SUBTOTAL(9,A1:A3)", 60],
    ["=ISNUMBER(A1)", true],
    ["=ISERROR(1/0)", true],
    ["=ISBLANK(Z9)", true],
    ['=TEXT(0.1234,"0.0%")', "12.3%"],
    ['=TEXT(1234.5,"#,##0.00")', "1,234.50"],
  ] as const)("evaluates %s", (formula, expected) => {
    expect(evalIn(data, formula)).toBe(expected);
  });

  it("VLOOKUP finds exact matches", () => {
    const workbook = workbookOf({
      S: { A1: "a", B1: 1, A2: "b", B2: 2, A3: "c", B3: 3 },
    });
    expect(evaluateFormula('=VLOOKUP("b",A1:B3,2,FALSE)', contextFor(workbook, "S", 9, 9))).toBe(2);
    expect(evaluateFormula('=VLOOKUP("z",A1:B3,2,FALSE)', contextFor(workbook, "S", 9, 9))).toBe(
      ERRORS.na
    );
  });

  it("XLOOKUP returns the fallback when nothing matches", () => {
    const workbook = workbookOf({ S: { A1: "a", B1: 1 } });
    expect(
      evaluateFormula('=XLOOKUP("z",A1:A1,B1:B1,"none")', contextFor(workbook, "S", 9, 9))
    ).toBe("none");
  });

  it("NPV discounts each period", () => {
    const value = evalIn({}, "=NPV(0.1,100,100)");
    expect(typeof value === "number" ? Number(value.toFixed(4)) : value).toBe(173.5537);
  });

  it("supports LET bindings", () => {
    expect(evalIn({ A1: 5 }, "=LET(x,A1,y,x*2,y+1)")).toBe(11);
  });

  it("supports immediately-applied LAMBDA", () => {
    expect(evalIn({ A1: 5 }, "=LAMBDA(v,v*3)(A1)")).toBe(15);
  });

  it("returns #NAME? for unsupported functions instead of guessing", () => {
    expect(evalIn({}, "=BESSELJ(1,2)")).toBe(ERRORS.name);
  });

  it("evaluates array literals", () => {
    expect(evalIn({}, "=SUM({1,2;3,4})")).toBe(10);
  });
});

describe("simulator recalculation", () => {
  it("recalculates in dependency order", () => {
    const workbook = workbookOf({
      S: { A1: 10, B1: "=A1*2", C1: "=B1+5", D1: "=C1*10" },
    });
    const simulator = Simulator.of(workbook);
    const result = simulator.recalculate();
    expect(result.circular).toEqual([]);
    expect(simulator.read("S", 0, 1)).toBe(20);
    expect(simulator.read("S", 0, 2)).toBe(25);
    expect(simulator.read("S", 0, 3)).toBe(250);
  });

  it("propagates an input change through the whole chain", () => {
    const workbook = workbookOf({
      S: { A1: 10, B1: "=A1*2", C1: "=B1+5" },
    });
    const simulator = Simulator.of(workbook);
    simulator.recalculate();
    simulator.write("S", 0, 0, 100);
    const result = simulator.recalculate();
    expect(simulator.read("S", 0, 2)).toBe(205);
    expect(result.changed.size).toBe(2);
  });

  it("marks genuinely circular cells rather than looping", () => {
    const workbook = workbookOf({ S: { A1: "=B1+1", B1: "=A1+1" } });
    const simulator = Simulator.of(workbook);
    const result = simulator.recalculate();
    expect(result.circular).toHaveLength(2);
    expect(simulator.read("S", 0, 0)).toBe(ERRORS.circular);
  });

  it("handles a cascading fill without reporting it circular", () => {
    const workbook = workbookOf({
      S: { A1: 1, B1: "=A1+1", C1: "=B1+1", D1: "=C1+1" },
    });
    const simulator = Simulator.of(workbook);
    const result = simulator.recalculate();
    expect(result.circular).toEqual([]);
    expect(simulator.read("S", 0, 3)).toBe(4);
  });

  it("reports functions it cannot evaluate", () => {
    const workbook = workbookOf({ S: { A1: "=BESSELJ(1,2)" } });
    const result = Simulator.of(workbook).recalculate();
    expect(result.unsupportedFunctions).toContain("BESSELJ");
  });

  describe("corpus models compute end to end", () => {
    it("3-statement model balances (assets == liabilities + equity)", () => {
      const { workbook } = threeStatementModel();
      const simulator = Simulator.of(workbook);
      const result = simulator.recalculate();
      expect(result.unsupportedFunctions).toEqual([]);
      expect(result.circular).toEqual([]);
      // Row 8 (0-based) is the tie-out check; every year must be ~0.
      for (let col = 1; col <= 5; col++) {
        const check = simulator.read("BS", 8, col);
        expect(typeof check).toBe("number");
        expect(Math.abs(check as number)).toBeLessThan(1e-6);
      }
    });

    it("DCF model produces a positive per-share value", () => {
      const { workbook } = dcfModel();
      const simulator = Simulator.of(workbook);
      const result = simulator.recalculate();
      expect(result.unsupportedFunctions).toEqual([]);
      const perShare = simulator.read("DCF", 16, 1);
      expect(typeof perShare).toBe("number");
      expect(perShare as number).toBeGreaterThan(0);
    });

    it("budget vs actual totals tie across sheets", () => {
      const { workbook } = budgetVsActual();
      const simulator = Simulator.of(workbook);
      simulator.recalculate();
      const budgetTotal = simulator.read("Budget", 6, 13) as number;
      const actualTotal = simulator.read("Actual", 6, 13) as number;
      const varianceTotal = simulator.read("Variance", 6, 13) as number;
      expect(Math.abs(actualTotal - budgetTotal - varianceTotal)).toBeLessThan(1e-6);
    });
  });
});

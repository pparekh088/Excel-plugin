import { describe, expect, it } from "vitest";
import { parseFormula } from "../../src/parser/parser";
import { extractRefs } from "../../src/parser/refs";
import { serialize } from "../../src/parser/serialize";

function refs(formula: string) {
  const parsed = parseFormula(formula);
  return extractRefs(parsed.ast);
}

/** Compact "A1"/"A1:B2" rendering of extracted areas for assertions. */
function areaTexts(formula: string): string[] {
  return refs(formula).areas.map((area) => {
    const sheet = area.sheet.start ? `${area.sheet.start}!` : "";
    const rows = `${area.startRow}-${area.endRow}`;
    const cols = `${area.startCol}-${area.endCol}`;
    return `${sheet}r${rows}c${cols}`;
  });
}

describe("reference extraction", () => {
  it("extracts single cells", () => {
    expect(areaTexts("=A1")).toEqual(["r0-0c0-0"]);
  });

  it("extracts ranges with normalized (sorted) bounds", () => {
    expect(areaTexts("=B2:A1")).toEqual(["r0-1c0-1"]);
  });

  it("extracts column and row ranges with null on the open axis", () => {
    const col = refs("=SUM(A:C)").areas[0]!;
    expect(col).toMatchObject({ kind: "colRange", startRow: null, startCol: 0, endCol: 2 });
    const row = refs("=SUM(2:4)").areas[0]!;
    expect(row).toMatchObject({ kind: "rowRange", startCol: null, startRow: 1, endRow: 3 });
  });

  it("collects every operand of an arithmetic chain", () => {
    expect(areaTexts("=A1+B1*C1-D1")).toEqual([
      "r0-0c0-0",
      "r0-0c1-1",
      "r0-0c2-2",
      "r0-0c3-3",
    ]);
  });

  it("records sheet-qualified refs with their sheet", () => {
    expect(areaTexts("=Sheet2!A1+Sheet3!B2")).toEqual([
      "Sheet2!r0-0c0-0",
      "Sheet3!r1-1c1-1",
    ]);
  });

  it("flags 3D spans", () => {
    expect(refs("=SUM(Sheet1:Sheet3!A1)").threeD).toBe(true);
    expect(refs("=SUM(Sheet1!A1)").threeD).toBe(false);
  });

  it("flags external workbook refs", () => {
    expect(refs("=[Book1.xlsx]Sheet1!A1").external).toBe(true);
    expect(refs("=Sheet1!A1").external).toBe(false);
  });

  it("collects defined names", () => {
    const result = refs("=Revenue*(1+Growth)");
    expect(result.names.map((n) => n.name)).toEqual(["Revenue", "Growth"]);
  });

  it("collects sheet-scoped names", () => {
    const result = refs("=Sheet1!LocalName");
    expect(result.names[0]).toMatchObject({ name: "LocalName" });
    expect(result.names[0]!.sheet.start).toBe("Sheet1");
  });

  it("does not treat TRUE/FALSE as names", () => {
    expect(refs("=IF(TRUE,1,FALSE)").names).toEqual([]);
  });

  it("collects structured refs", () => {
    const result = refs("=SUM(Sales[Amount])/SUM(Sales[Qty])");
    expect(result.structured.map((s) => `${s.table}.${s.columns.join("+")}`)).toEqual([
      "Sales.Amount",
      "Sales.Qty",
    ]);
  });

  it("collects function names, canonicalized and deduplicated", () => {
    expect(refs("=SUM(A1)+sum(B1)+IF(C1,1,2)").functions).toEqual(["IF", "SUM"]);
  });

  describe("opacity (INV-4 honesty)", () => {
    it("flags INDIRECT", () => {
      expect(refs('=INDIRECT("A"&B1)').opaque).toBe(true);
    });

    it("flags OFFSET", () => {
      expect(refs("=SUM(OFFSET(A1,1,0,10,1))").opaque).toBe(true);
    });

    it("flags computed range endpoints", () => {
      expect(refs("=SUM(A1:INDEX(A:A,10))").opaque).toBe(true);
    });

    it("does not flag INDEX used as a plain lookup", () => {
      expect(refs("=INDEX(A1:C10,2,3)").opaque).toBe(false);
    });

    it("still extracts the literal refs inside an opaque formula", () => {
      const result = refs("=SUM(OFFSET(A1,1,0,10,1))");
      expect(result.areas).toHaveLength(1);
      expect(result.opaque).toBe(true);
    });
  });

  describe("volatility", () => {
    it.each(["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"])(
      "flags %s as volatile",
      (fn) => {
        expect(refs(`=${fn}()`).volatile).toBe(true);
      }
    );

    it("does not flag ordinary functions", () => {
      expect(refs("=SUM(A1:A10)").volatile).toBe(false);
    });
  });

  it("counts #REF! literals as broken references", () => {
    expect(refs("=#REF!+#REF!").brokenRefs).toBe(2);
    expect(refs("=A1").brokenRefs).toBe(0);
  });

  it("counts spill references", () => {
    expect(refs("=A1#+B1#").spills).toBe(2);
  });

  describe("LET/LAMBDA scoping", () => {
    it("LET locals do not leak as workbook names", () => {
      const result = refs("=LET(x,A1,y,B1,x*y)");
      expect(result.names).toEqual([]);
      expect(result.areas).toHaveLength(2);
    });

    it("LET body may reference workbook names not bound locally", () => {
      const result = refs("=LET(x,A1,x*TaxRate)");
      expect(result.names.map((n) => n.name)).toEqual(["TaxRate"]);
    });

    it("LET value expressions may use earlier locals", () => {
      const result = refs("=LET(x,A1,y,x+1,y*2)");
      expect(result.names).toEqual([]);
    });

    it("a LET local shadows a same-named workbook name", () => {
      const result = refs("=LET(Revenue,A1,Revenue*2)");
      expect(result.names).toEqual([]);
    });

    it("LAMBDA parameters do not leak", () => {
      const result = refs("=BYROW(A1:C10,LAMBDA(r,SUM(r)))");
      expect(result.names).toEqual([]);
      expect(result.areas).toHaveLength(1);
    });

    it("LAMBDA body still sees outer refs", () => {
      const result = refs("=MAP(A1:A10,LAMBDA(v,v*Factor))");
      expect(result.names.map((n) => n.name)).toEqual(["Factor"]);
    });
  });

  it("array literals contribute no refs", () => {
    const result = refs("=MMULT({1,2;3,4},{5;6})");
    expect(result.areas).toEqual([]);
    expect(result.names).toEqual([]);
  });

  it("union and intersection contribute all operand areas", () => {
    expect(areaTexts("=SUM((A1:A3,C1:C3))")).toEqual(["r0-2c0-0", "r0-2c2-2"]);
    expect(areaTexts("=B5:B15 C5:C15")).toEqual(["r4-14c1-1", "r4-14c2-2"]);
  });

  it("never throws on unparseable input and still reports what it found", () => {
    const parsed = parseFormula("=SUM(A1:A10");
    const result = extractRefs(parsed.ast);
    expect(parsed.ok).toBe(false);
    expect(result.areas.length).toBeGreaterThan(0);
    expect(serialize(parsed.ast)).toContain("A1");
  });
});

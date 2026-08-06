import { describe, expect, it } from "vitest";
import { Node } from "../../src/parser/ast";
import { parseFormula } from "../../src/parser/parser";

function ast(input: string): Node {
  const result = parseFormula(input);
  expect(result.ok, `expected clean parse for ${input}`).toBe(true);
  return result.ast;
}

describe("AST shapes", () => {
  it("negation binds tighter than ^ (Excel: -2^2 = 4)", () => {
    const node = ast("=-2^2");
    expect(node.kind).toBe("binary");
    if (node.kind !== "binary") return;
    expect(node.op).toBe("^");
    expect(node.left.kind).toBe("unary");
    expect(node.right.kind).toBe("number");
  });

  it("^ is left-associative (Excel: 2^3^2 = 64)", () => {
    const node = ast("=2^3^2");
    if (node.kind !== "binary") throw new Error("expected binary");
    expect(node.left.kind).toBe("binary");
    expect(node.right.kind).toBe("number");
  });

  it("percent applies after negation", () => {
    const node = ast("=-A1%");
    expect(node.kind).toBe("percent");
    if (node.kind !== "percent") return;
    expect(node.operand.kind).toBe("unary");
  });

  it("concatenation binds looser than +", () => {
    const node = ast('="x"&1+2');
    if (node.kind !== "binary") throw new Error("expected binary");
    expect(node.op).toBe("&");
    expect(node.right.kind).toBe("binary");
  });

  it("comparison is loosest", () => {
    const node = ast('=A1&B1="ab"');
    if (node.kind !== "binary") throw new Error("expected binary");
    expect(node.op).toBe("=");
    expect(node.left.kind).toBe("binary");
  });

  it("missing args occupy their slots", () => {
    const node = ast("=IF(A1,,B1)");
    if (node.kind !== "func") throw new Error("expected func");
    expect(node.args).toHaveLength(3);
    expect(node.args[1]!.kind).toBe("missing");
  });

  it("trailing comma produces a trailing missing arg", () => {
    const node = ast("=IF(A1,B1,)");
    if (node.kind !== "func") throw new Error("expected func");
    expect(node.args).toHaveLength(3);
    expect(node.args[2]!.kind).toBe("missing");
  });

  it("empty call has zero args", () => {
    const node = ast("=PI()");
    if (node.kind !== "func") throw new Error("expected func");
    expect(node.args).toHaveLength(0);
  });

  it("TRUE alone is a boolean literal; TRUE() is a function", () => {
    expect(ast("=TRUE").kind).toBe("bool");
    expect(ast("=TRUE()").kind).toBe("func");
  });

  it("LOG10 parses as a function despite matching the cell pattern", () => {
    const node = ast("=LOG10(100)");
    if (node.kind !== "func") throw new Error("expected func");
    expect(node.name).toBe("LOG10");
  });

  it("cell:cell folds to a rectangular range", () => {
    const node = ast("=Sheet1!A1:B2");
    expect(node.kind).toBe("range");
    if (node.kind !== "range") return;
    expect(node.sheet.start).toBe("Sheet1");
    expect(node.start).toMatchObject({ row: 0, col: 0 });
    expect(node.end).toMatchObject({ row: 1, col: 1 });
  });

  it("A:D folds to a column range; TAX:REV stays column-valid", () => {
    const node = ast("=A:D");
    expect(node.kind).toBe("colRange");
    const tax = ast("=SUM(TAX:TAX)");
    if (tax.kind !== "func") throw new Error("expected func");
    expect(tax.args[0]!.kind).toBe("colRange");
  });

  it("1:3 folds to a row range", () => {
    const node = ast("=1:3");
    expect(node.kind).toBe("rowRange");
    if (node.kind !== "rowRange") return;
    expect(node.startRow).toBe(0);
    expect(node.endRow).toBe(2);
  });

  it("3D span sheets carry start and end", () => {
    const node = ast("=SUM(Sheet1:Sheet3!A1)");
    if (node.kind !== "func") throw new Error("expected func");
    const arg = node.args[0]!;
    if (arg.kind !== "cell") throw new Error("expected cell");
    expect(arg.sheet).toMatchObject({ start: "Sheet1", end: "Sheet3" });
  });

  it("external workbook refs carry the workbook name", () => {
    const node = ast("=[Book1.xlsx]Sheet1!A1");
    if (node.kind !== "cell") throw new Error("expected cell");
    expect(node.sheet.external).toBe("Book1.xlsx");
  });

  it("union inside parens over refs", () => {
    const node = ast("=(A1,B2)");
    if (node.kind !== "group") throw new Error("expected group");
    if (node.expr.kind !== "binary") throw new Error("expected binary");
    expect(node.expr.op).toBe(",");
  });

  it("intersection via whitespace over refs", () => {
    const node = ast("=B5:B15 C5:C15");
    if (node.kind !== "binary") throw new Error("expected binary");
    expect(node.op).toBe(" ");
    expect(node.left.kind).toBe("range");
    expect(node.right.kind).toBe("range");
  });

  it("colon binds tighter than intersection", () => {
    const node = ast("=A1:A10 B2:B10");
    if (node.kind !== "binary") throw new Error("expected binary");
    expect(node.op).toBe(" ");
    expect(node.left.kind).toBe("range");
  });

  it("computed range endpoint stays a binary ':'", () => {
    const node = ast("=INDEX(A1:C10,2,3):E5");
    if (node.kind !== "binary") throw new Error("expected binary");
    expect(node.op).toBe(":");
    expect(node.left.kind).toBe("func");
    expect(node.right.kind).toBe("cell");
  });

  it("lambda immediate call becomes callExpr", () => {
    const node = ast("=LAMBDA(x,x*2)(A1)");
    if (node.kind !== "callExpr") throw new Error("expected callExpr");
    expect(node.callee.kind).toBe("func");
    expect(node.args).toHaveLength(1);
  });

  it("curried lambda call chains callExpr", () => {
    const node = ast("=LAMBDA(x,LAMBDA(y,x+y))(1)(2)");
    if (node.kind !== "callExpr") throw new Error("expected callExpr");
    expect(node.callee.kind).toBe("callExpr");
  });

  it("spill operator wraps the ref", () => {
    const node = ast("=A1#");
    expect(node.kind).toBe("spill");
  });

  it("implicit intersection captures the full range", () => {
    const node = ast("=@A1:A10");
    if (node.kind !== "implicitIntersection") throw new Error("expected @");
    expect(node.operand.kind).toBe("range");
  });

  it("structured ref parses table, items, columns", () => {
    const node = ast("=Table1[[#Headers],[Amount]]");
    if (node.kind !== "structured") throw new Error("expected structured");
    expect(node.table).toBe("Table1");
    expect(node.items).toEqual(["#Headers"]);
    expect(node.columns).toEqual(["Amount"]);
  });

  it("bare [@Col] is a this-row structured ref with empty table", () => {
    const node = ast("=[@Col]");
    if (node.kind !== "structured") throw new Error("expected structured");
    expect(node.table).toBe("");
    expect(node.thisRow).toBe(true);
    expect(node.columns).toEqual(["Col"]);
  });

  it("structured column span", () => {
    const node = ast("=Table1[[Q1]:[Q4]]");
    if (node.kind !== "structured") throw new Error("expected structured");
    expect(node.columns).toEqual(["Q1", "Q4"]);
  });

  it("_xlfn. prefix is stripped to the canonical name", () => {
    const node = ast("=_xlfn.XLOOKUP(A1,B:B,C:C)");
    if (node.kind !== "func") throw new Error("expected func");
    expect(node.name).toBe("XLOOKUP");
    expect(node.rawName).toBe("_xlfn.XLOOKUP");
  });

  it("quoted 3D span splits into start/end sheets", () => {
    const node = ast("='Jan Data:Mar Data'!B2");
    if (node.kind !== "cell") throw new Error("expected cell");
    expect(node.sheet).toMatchObject({ start: "Jan Data", end: "Mar Data" });
  });

  it("array literal rows and columns", () => {
    const node = ast("={1,2;3,4}");
    if (node.kind !== "array") throw new Error("expected array");
    expect(node.rows).toHaveLength(2);
    expect(node.rows[0]).toHaveLength(2);
  });

  it("legacy CSE braces are stripped", () => {
    const node = ast("{=SUM(A1:A3)}");
    expect(node.kind).toBe("func");
  });
});

describe("parse diagnostics on odd-but-recoverable inputs", () => {
  it("Sheet1!#REF! yields an error literal plus a diagnostic", () => {
    const result = parseFormula("=Sheet1!#REF!+1");
    expect(result.ok).toBe(false);
    expect(result.diagnostics.length).toBeGreaterThan(0);
    expect(result.ast.kind).toBe("binary");
  });

  it("cross-sheet range endpoints produce a diagnostic", () => {
    const result = parseFormula("=Sheet1!A1:Sheet2!B2");
    expect(result.ok).toBe(false);
    expect(
      result.diagnostics.some((d) => d.message.includes("different sheets"))
    ).toBe(true);
  });
});

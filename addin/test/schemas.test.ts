import { describe, expect, it } from "vitest";
import { rangeReadParams, rangeReadResult, toolDefinitions } from "../src/tools/schemas";

describe("rangeReadParams", () => {
  it("accepts a well-formed call and applies the include default", () => {
    const parsed = rangeReadParams.parse({ sheet: "Model", a1: "B2:D500" });
    expect(parsed.include).toEqual(["values"]);
  });

  it.each(["A1", "A1:D10", "$A$1:$D$10", "AAA1:AAB2", "A:D", "3:7"])(
    "accepts A1 form %s",
    (a1) => {
      expect(rangeReadParams.safeParse({ sheet: "S", a1 }).success).toBe(true);
    }
  );

  it.each([
    "Sheet1!A1", // sheet travels separately
    "A0", // rows are 1-based
    "1A",
    "A1:D",
    "",
    "A1 :D10",
    "R1C1",
  ])("rejects malformed A1 form %s", (a1) => {
    expect(rangeReadParams.safeParse({ sheet: "S", a1 }).success).toBe(false);
  });

  it("rejects unknown properties (strict)", () => {
    const parsed = rangeReadParams.safeParse({ sheet: "S", a1: "A1", extra: true });
    expect(parsed.success).toBe(false);
  });

  it("rejects an empty include list and out-of-range maxCells", () => {
    expect(rangeReadParams.safeParse({ sheet: "S", a1: "A1", include: [] }).success).toBe(false);
    expect(
      rangeReadParams.safeParse({ sheet: "S", a1: "A1", maxCells: 300_000 }).success
    ).toBe(false);
  });
});

describe("rangeReadResult", () => {
  const base = {
    sheet: "Model",
    a1: "Model!A1:B2",
    rowCount: 2,
    columnCount: 2,
    cellCount: 4,
    chunkCount: 1,
  };

  it("accepts grids with the Office.js value union (formulas may hold raw values)", () => {
    const parsed = rangeReadResult.safeParse({
      ...base,
      values: [["x", 1], [true, ""]],
      formulas: [["=B1*2", 42], [true, ""]],
      numberFormats: [["General", "0.00%"], ["General", "General"]],
    });
    expect(parsed.success).toBe(true);
  });

  it("rejects non-string number formats and non-integer counts", () => {
    expect(
      rangeReadResult.safeParse({ ...base, numberFormats: [[1, "General"]] }).success
    ).toBe(false);
    expect(rangeReadResult.safeParse({ ...base, rowCount: 1.5 }).success).toBe(false);
  });
});

describe("toolDefinitions", () => {
  it("registers range.read as a no-risk read tool", () => {
    const def = toolDefinitions["range.read"];
    expect(def).toBeDefined();
    expect(def?.access).toBe("read");
    expect(def?.risk).toBe("none");
  });
});

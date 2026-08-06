import { describe, expect, it } from "vitest";
import { parseFormula } from "../../src/parser/parser";
import { serialize } from "../../src/parser/serialize";
import { CORPUS } from "./corpus";

describe("round-trip corpus", () => {
  it.each(CORPUS.map((entry) => [entry.input, entry.canonical] as const))(
    "parses and canonicalizes %s",
    (input, canonical) => {
      const expected =
        canonical ?? (input.startsWith("=") ? input.slice(1) : input.replace(/^\{=|\}$/g, ""));
      const result = parseFormula(input);
      expect(result.diagnostics, `diagnostics for ${input}`).toEqual([]);
      expect(result.ok, `ok for ${input}`).toBe(true);
      expect(serialize(result.ast)).toBe(expected);
    }
  );

  it.each(CORPUS.map((entry) => [entry.input, entry.canonical] as const))(
    "canonical form is a fixpoint for %s",
    (input, canonical) => {
      const first = serialize(parseFormula(input).ast);
      const second = parseFormula("=" + first);
      expect(second.ok, `reparse ok for ${first}`).toBe(true);
      expect(serialize(second.ast)).toBe(first);
      void canonical;
    }
  );
});

import { describe, expect, it } from "vitest";
import { computeHealth, runAudit, safeAutoFixes } from "../../src/audit/engine";
import { Simulator } from "../../src/sim/simulator";
import { Workbook } from "../../src/model/workbook";
import { workbookOf } from "../helpers/build";
import { threeStatementModel, dcfModel } from "../../src/corpus/models";
import { Finding } from "../../src/audit/types";

function audit(workbook: Workbook, rules?: string[]) {
  Simulator.of(workbook).recalculate();
  return runAudit(workbook, rules ? { rules } : {});
}

function rulesFired(findings: Finding[], rule: string): string[] {
  return findings.filter((f) => f.rule === rule).map((f) => f.address);
}

describe("AUD-001 formula inconsistency", () => {
  it("flags a cell that breaks a filled row", () => {
    const report = audit(
      workbookOf({
        S: {
          A1: 1, B1: 2, C1: 3, D1: 4, E1: 5,
          A2: "=A1*2", B2: "=B1*2", C2: "=C1*3", D2: "=D1*2", E2: "=E1*2",
        },
      }),
      ["AUD-001"]
    );
    expect(rulesFired(report.findings, "AUD-001")).toEqual(["S!C2"]);
  });

  it("flags a cell that breaks a filled column", () => {
    const report = audit(
      workbookOf({
        S: {
          A1: 1, A2: 2, A3: 3, A4: 4, A5: 5,
          B1: "=A1*2", B2: "=A2*2", B3: "=A3*3", B4: "=A4*2", B5: "=A5*2",
        },
      }),
      ["AUD-001"]
    );
    expect(rulesFired(report.findings, "AUD-001")).toEqual(["S!B3"]);
  });

  it("does not flag the edge of a block", () => {
    // The run simply ends; there is no pattern on the far side.
    const report = audit(
      workbookOf({
        S: { A1: 1, B1: 2, C1: 3, A2: "=A1*2", B2: "=B1*2", C2: "=SUM(A2:B2)" },
      }),
      ["AUD-001"]
    );
    expect(rulesFired(report.findings, "AUD-001")).toEqual([]);
  });

  it("does not flag rows that coincidentally share a signature", () => {
    // Gross profit and EBITDA are both "sum of the two rows above" — vertically
    // adjacent line items, not a fill pattern.
    const report = audit(
      workbookOf({
        S: {
          B1: 100, C1: 110, D1: 120,
          B2: "=B1*0.4", C2: "=C1*0.4", D2: "=D1*0.4",
          B3: "=B1+B2", C3: "=C1+C2", D3: "=D1+D2",
          B4: "=B1*0.1", C4: "=C1*0.1", D4: "=D1*0.1",
          B5: "=B3+B4", C5: "=C3+C4", D5: "=D3+D4",
        },
      }),
      ["AUD-001"]
    );
    expect(rulesFired(report.findings, "AUD-001")).toEqual([]);
  });

  it("still finds a hole when other defects fragment the block", () => {
    // A 2D block with two unrelated holes; the middle cell is still anomalous.
    const report = audit(
      workbookOf({
        S: {
          A1: 1, B1: 2, C1: 3, D1: 4, E1: 5,
          A2: 1, B2: 2, C2: 3, D2: 4, E2: 5,
          A3: "=A1*2", B3: "=B1*2", C3: "=C1*9", D3: "=D1*2", E3: "=E1*2",
          A4: "=A2*2", B4: 999, C4: "=C2*2", D4: "=D2*2", E4: "=E2*2",
        },
      }),
      ["AUD-001"]
    );
    expect(rulesFired(report.findings, "AUD-001")).toContain("S!C3");
  });

  it("reports zero findings on clean corpus models", () => {
    for (const { workbook } of [threeStatementModel(), dcfModel()]) {
      expect(rulesFired(audit(workbook, ["AUD-001"]).findings, "AUD-001")).toEqual([]);
    }
  });
});

describe("AUD-002 hardcoded value inside a formula", () => {
  it("flags a magic multiplier", () => {
    const report = audit(workbookOf({ S: { A1: 100, B1: "=A1*1.085" } }), ["AUD-002"]);
    expect(rulesFired(report.findings, "AUD-002")).toEqual(["S!B1"]);
  });

  it("flags a literal even when the formula has no references", () => {
    const report = audit(workbookOf({ S: { A1: "=1/(1+0.095)^3" } }), ["AUD-002"]);
    expect(rulesFired(report.findings, "AUD-002")).toEqual(["S!A1"]);
  });

  it("does not flag structural arguments", () => {
    const workbook = workbookOf({
      S: {
        A1: 1, A2: 2, A3: 3, B1: "x", B2: "y", B3: "z",
        D1: "=INDEX(A1:A3,2)",
        D2: "=VLOOKUP(B1,A1:B3,2,FALSE)",
        D3: "=ROUND(A1,2)",
        D4: "=LEFT(B1,3)",
        D5: "=A1/365",
        D6: "=A1*100",
      },
    });
    expect(rulesFired(audit(workbook, ["AUD-002"]).findings, "AUD-002")).toEqual([]);
  });

  it("does not flag a discount-factor period exponent", () => {
    const workbook = workbookOf({ S: { A1: 0.09, B1: "=1/(1+A1)^5" } });
    expect(rulesFired(audit(workbook, ["AUD-002"]).findings, "AUD-002")).toEqual([]);
  });

  it("does not flag a percentage literal", () => {
    const workbook = workbookOf({ S: { A1: 100, B1: "=A1*5%" } });
    expect(rulesFired(audit(workbook, ["AUD-002"]).findings, "AUD-002")).toEqual([]);
  });

  it("does not flag a bare literal (that is AUD-003 territory)", () => {
    const workbook = workbookOf({ S: { A1: "=41250", B1: "=A1*2" } });
    expect(rulesFired(audit(workbook, ["AUD-002"]).findings, "AUD-002")).toEqual([]);
  });
});

describe("AUD-003 hardcoded value in a calculation chain", () => {
  it("flags a plug surrounded by formulas", () => {
    const workbook = workbookOf({
      S: {
        A1: 10, A2: 20, A3: 30, A4: 40,
        B1: "=A1*2", B2: 999, B3: "=A3*2", B4: "=A4*2",
        C2: "=B2+1",
      },
    });
    expect(rulesFired(audit(workbook, ["AUD-003"]).findings, "AUD-003")).toEqual(["S!B2"]);
  });

  it("does not flag constants in a labelled input block", () => {
    const { workbook } = threeStatementModel();
    expect(rulesFired(audit(workbook, ["AUD-003"]).findings, "AUD-003")).toEqual([]);
  });

  it("does not flag a constant nothing reads", () => {
    const workbook = workbookOf({ S: { A1: "=1+1", B1: 42, C1: "=A1*2" } });
    expect(rulesFired(audit(workbook, ["AUD-003"]).findings, "AUD-003")).toEqual([]);
  });
});

describe("AUD-004 error cells", () => {
  it("reports the error and its blast radius", () => {
    const workbook = workbookOf({ S: { A1: "#REF!", B1: "=A1*2", C1: "=B1+1" } });
    const report = audit(workbook, ["AUD-004"]);
    const finding = report.findings.find((f) => f.address === "S!A1");
    expect(finding).toBeDefined();
    expect(finding!.severity).toBe("critical");
    expect(finding!.explanation).toContain("downstream");
  });

  it("says so when nothing depends on the error", () => {
    const workbook = workbookOf({ S: { A1: "#DIV/0!" } });
    const finding = audit(workbook, ["AUD-004"]).findings[0]!;
    expect(finding.explanation).toContain("Nothing depends on it");
  });
});

describe("AUD-005 circular references", () => {
  it("reports a real cycle once, listing its members", () => {
    const report = audit(workbookOf({ S: { A1: "=B1+1", B1: "=A1*2" } }), ["AUD-005"]);
    expect(report.findings).toHaveLength(1);
    expect(report.findings[0]!.explanation).toContain("->");
  });

  it("does not report cascading fills", () => {
    const workbook = workbookOf({ S: { A1: 1, B1: "=A1+1", C1: "=B1+1", D1: "=C1+1" } });
    expect(audit(workbook, ["AUD-005"]).findings).toEqual([]);
  });
});

describe("AUD-006 references to empty cells", () => {
  it("flags a single-cell reference to a blank", () => {
    const report = audit(workbookOf({ S: { A1: "=Z9*2" } }), ["AUD-006"]);
    expect(rulesFired(report.findings, "AUD-006")).toEqual(["S!A1"]);
  });

  it("does not flag blanks inside an aggregate range", () => {
    const workbook = workbookOf({ S: { A1: 1, A5: 2, B1: "=SUM(A1:A10)" } });
    expect(rulesFired(audit(workbook, ["AUD-006"]).findings, "AUD-006")).toEqual([]);
  });
});

describe("AUD-008 balance assertions", () => {
  it("auto-detects a labelled check row and passes when it ties", () => {
    const { workbook } = threeStatementModel();
    const report = audit(workbook, ["AUD-008"]);
    expect(report.rulesRun).toContain("AUD-008");
    expect(rulesFired(report.findings, "AUD-008")).toEqual([]);
  });

  it("fails loudly when the check row does not tie", () => {
    const { workbook } = threeStatementModel();
    // Break equity so assets no longer equal liabilities + equity.
    workbook.sheet("BS")!.set({ row: 6, col: 2, value: 0, formula: "=B7+IS!C10+5000" });
    const report = audit(workbook, ["AUD-008"]);
    const findings = rulesFired(report.findings, "AUD-008");
    expect(findings.length).toBeGreaterThan(0);
    expect(report.findings[0]!.severity).toBe("critical");
    expect(report.findings[0]!.explanation).toContain("off by");
  });

  it("accepts explicit assertions", () => {
    const workbook = workbookOf({ S: { A1: 5 } });
    const report = runAudit(workbook, {
      rules: ["AUD-008"],
      assertions: [{ name: "A1 is zero", sheet: "S", a1: "A1", expected: 0 }],
      skipAssertionDetection: true,
    });
    expect(report.findings).toHaveLength(1);
  });
});

describe("AUD-009 volatile functions", () => {
  it("flags volatile usage and names the functions", () => {
    const report = audit(workbookOf({ S: { A1: "=NOW()+RAND()" } }), ["AUD-009"]);
    expect(report.findings[0]!.explanation).toContain("NOW");
    expect(report.findings[0]!.explanation).toContain("RAND");
  });

  it("escalates severity when volatiles are widespread", () => {
    const cells: Record<string, string> = {};
    for (let row = 1; row <= 60; row++) cells[`A${row}`] = `=NOW()+${row}`;
    const report = audit(workbookOf({ S: cells }), ["AUD-009"]);
    expect(report.findings[0]!.severity).toBe("medium");
  });
});

describe("AUD-011 external and opaque references", () => {
  it("flags an external workbook link", () => {
    const report = audit(workbookOf({ S: { A1: "=[Other.xlsx]Sheet1!A1*2" } }), ["AUD-011"]);
    expect(report.findings[0]!.explanation).toContain("another workbook");
  });

  it("flags INDIRECT as hiding its references", () => {
    const report = audit(workbookOf({ S: { A1: '=INDIRECT("B"&1)' } }), ["AUD-011"]);
    expect(report.findings[0]!.explanation).toContain("INDIRECT");
  });
});

describe("audit report", () => {
  it("uses zero LLM calls (INV: demoable inside a compliance boundary)", () => {
    const { workbook } = threeStatementModel();
    expect(audit(workbook).stats.llmCallsUsed).toBe(0);
  });

  it("states coverage honestly", () => {
    const clean = audit(threeStatementModel().workbook);
    expect(clean.coverage.complete).toBe(true);
    expect(clean.coverage.caveat).toContain("complete");

    const opaque = audit(workbookOf({ S: { A1: '=INDIRECT("A2")' } }));
    expect(opaque.coverage.complete).toBe(false);
    expect(opaque.coverage.caveat).toContain("INCOMPLETE");
  });

  it("ranks findings by severity then blast radius", () => {
    const workbook = workbookOf({
      S: {
        A1: "#REF!",
        B1: "=A1*2",
        C1: "=B1*1.085",
        D1: "=NOW()",
      },
    });
    const report = audit(workbook);
    const severities = report.findings.map((f) => f.severity);
    const rank = (s: string) => ["critical", "high", "medium", "low", "info"].indexOf(s);
    for (let i = 1; i < severities.length; i++) {
      expect(rank(severities[i]!)).toBeGreaterThanOrEqual(rank(severities[i - 1]!));
    }
  });

  it("scores a clean workbook at 100 and a broken one lower", () => {
    const clean = audit(threeStatementModel().workbook);
    expect(clean.health.score).toBe(100);
    expect(clean.health.band).toBe("healthy");

    const broken = audit(workbookOf({ S: { A1: "#REF!", B1: "=A1*2", C1: "=B1+1" } }));
    expect(broken.health.score).toBeLessThan(100);
  });

  it("restricts findings to the requested sheets", () => {
    const workbook = workbookOf({
      Good: { A1: 1 },
      Bad: { A1: "#REF!" },
    });
    const report = runAudit(workbook, { sheets: ["Good"] });
    expect(report.findings).toEqual([]);
  });

  it("exposes only low-risk auto-fixes as safe", () => {
    const workbook = workbookOf({
      S: {
        A1: 1, B1: 2, C1: 3, D1: 4, E1: 5,
        A2: "=A1*2", B2: "=B1*2", C2: "=C1*3", D2: "=D1*2", E2: "=E1*2",
      },
    });
    const report = audit(workbook);
    // The AUD-001 fix overwrites a formula, so it must never be "safe".
    for (const finding of safeAutoFixes(report)) {
      expect(finding.autoFix!.risk).toBe("low");
    }
    expect(report.findings.some((f) => f.autoFix?.risk === "high")).toBe(true);
  });

  it("every finding carries a plain-language explanation and a trace", () => {
    const workbook = workbookOf({
      S: { A1: "#REF!", B1: "=A1*1.085", C1: "=NOW()", D1: "=Z9*2" },
    });
    for (const finding of audit(workbook).findings) {
      expect(finding.explanation.length).toBeGreaterThan(40);
      expect(finding.trace.length).toBeGreaterThan(0);
      expect(finding.title.length).toBeGreaterThan(0);
    }
  });
});

describe("health scoring", () => {
  it("is 100 with no findings", () => {
    expect(computeHealth([]).score).toBe(100);
  });

  it("penalizes critical findings more than low ones", () => {
    const base = {
      title: "t",
      confidence: "certain" as const,
      address: "S!A1",
      sheet: "S",
      row: 0,
      col: 0,
      explanation: "e",
      trace: [],
      blastRadius: 0,
    };
    const critical = computeHealth([{ ...base, rule: "AUD-004", severity: "critical" }]);
    const low = computeHealth([{ ...base, rule: "AUD-009", severity: "low" }]);
    expect(critical.score).toBeLessThan(low.score);
  });

  it("weights findings with a larger blast radius more heavily", () => {
    const base = {
      rule: "AUD-004",
      title: "t",
      severity: "critical" as const,
      confidence: "certain" as const,
      address: "S!A1",
      sheet: "S",
      row: 0,
      col: 0,
      explanation: "e",
      trace: [],
    };
    const wide = computeHealth([{ ...base, blastRadius: 5000 }]);
    const narrow = computeHealth([{ ...base, blastRadius: 0 }]);
    expect(wide.score).toBeLessThan(narrow.score);
  });

  it("stays within 0..100 under a flood of findings", () => {
    const findings = Array.from({ length: 500 }, (_, i) => ({
      rule: "AUD-004",
      title: "t",
      severity: "critical" as const,
      confidence: "certain" as const,
      address: `S!A${i}`,
      sheet: "S",
      row: i,
      col: 0,
      explanation: "e",
      trace: [],
      blastRadius: 100,
    }));
    const health = computeHealth(findings);
    expect(health.score).toBeGreaterThanOrEqual(0);
    expect(health.score).toBeLessThanOrEqual(100);
    expect(health.band).toBe("critical");
  });
});

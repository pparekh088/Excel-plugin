/**
 * The eval corpus (handoff §9): clean models plus intentionally broken
 * variants, each carrying its ground-truth defect list. Versioned by
 * CORPUS_VERSION so eval scores remain comparable over time.
 */

import { CorpusWorkbook } from "./models";
import {
  brokenVariant,
  injectBalanceBreak,
  injectCircular,
  injectEmptyRef,
  injectExternalLink,
  injectFormulaInconsistency,
  injectHardcodedInFormula,
  injectPluggedConstant,
  injectRefError,
  injectVolatiles,
} from "./defects";
import {
  budgetVsActual,
  cohortAnalysis,
  dcfModel,
  threeStatementModel,
} from "./models";

export const CORPUS_VERSION = 1;

export function buildCorpus(): CorpusWorkbook[] {
  const corpus: CorpusWorkbook[] = [];

  // --- clean baselines: any finding here is a false positive -------------
  corpus.push(threeStatementModel("clean-3statement"));
  corpus.push(dcfModel("clean-dcf"));
  corpus.push(budgetVsActual("clean-bva"));
  corpus.push(cohortAnalysis("clean-cohort"));

  // --- single-defect variants: isolate each rule ------------------------
  corpus.push(
    brokenVariant(
      threeStatementModel("x"),
      "broken-3s-inconsistency",
      "3-statement with a broken revenue fill",
      // Revenue row: D2 should follow the growth pattern; it points at the wrong driver row.
      [injectFormulaInconsistency("IS", "D2", "=C2*(1+Assumptions!D3)")]
    )
  );
  corpus.push(
    brokenVariant(
      threeStatementModel("x"),
      "broken-3s-hardcode-in-formula",
      "3-statement with a magic growth rate",
      [injectHardcodedInFormula("IS", "E2", "=D2*1.085")]
    )
  );
  corpus.push(
    brokenVariant(
      threeStatementModel("x"),
      "broken-3s-plugged-constant",
      "3-statement with a plugged EBITDA",
      [injectPluggedConstant("IS", "D6", 41_250)]
    )
  );
  corpus.push(
    brokenVariant(
      threeStatementModel("x"),
      "broken-3s-ref-error",
      "3-statement with a deleted precedent",
      [injectRefError("CF", "D2")]
    )
  );
  corpus.push(
    brokenVariant(
      dcfModel("x"),
      "broken-dcf-circular",
      "DCF with a circular EV/equity loop",
      [injectCircular("DCF", "B15", "B16")]
    )
  );
  corpus.push(
    brokenVariant(
      dcfModel("x"),
      "broken-dcf-empty-ref",
      "DCF referencing an empty assumption cell",
      [injectEmptyRef("DCF", "B18", "Inputs!B9")]
    )
  );
  corpus.push(
    brokenVariant(
      threeStatementModel("x"),
      "broken-3s-balance",
      "3-statement with a balance sheet that no longer ties",
      [injectBalanceBreak("BS", "D6", 1_500)]
    )
  );
  corpus.push(
    brokenVariant(
      dcfModel("x"),
      "broken-dcf-volatile",
      "DCF littered with volatile functions",
      [injectVolatiles("DCF", ["D19", "D20", "D21"])]
    )
  );
  corpus.push(
    brokenVariant(
      budgetVsActual("x"),
      "broken-bva-external",
      "Budget vs actual with an external link",
      [injectExternalLink("Variance", "B10")]
    )
  );

  // --- multi-defect variants: the realistic case ------------------------
  corpus.push(
    brokenVariant(
      threeStatementModel("x"),
      "broken-3s-multi",
      "3-statement with several unrelated defects",
      [
        injectFormulaInconsistency("IS", "D5", "=-D2*Assumptions!C4"),
        injectHardcodedInFormula("IS", "F3", "=-F2*(1-0.635)"),
        injectPluggedConstant("BS", "E6", 92_000),
        injectVolatiles("CF", ["G2"]),
      ]
    )
  );
  corpus.push(
    brokenVariant(
      dcfModel("x"),
      "broken-dcf-multi",
      "DCF with hardcodes and a broken discount factor",
      [
        injectHardcodedInFormula("DCF", "D9", "=1/(1+0.095)^3"),
        injectPluggedConstant("DCF", "E8", 61_000),
        injectExternalLink("DCF", "B20"),
      ]
    )
  );
  corpus.push(
    brokenVariant(
      cohortAnalysis("x"),
      "broken-cohort-multi",
      "Cohort model with a broken retention chain",
      [
        injectFormulaInconsistency("Cohorts", "E4", "=D4*0.75"),
        injectPluggedConstant("Cohorts", "F5", 640),
        injectEmptyRef("Cohorts", "J2", "Rates!B20"),
      ]
    )
  );
  corpus.push(
    brokenVariant(
      budgetVsActual("x"),
      "broken-bva-multi",
      "Budget vs actual with copy-paste damage",
      [
        injectFormulaInconsistency("Variance", "G3", "=Actual!G3-Budget!F3"),
        injectPluggedConstant("Variance", "H4", -1_250),
        injectRefError("Variance", "I5"),
      ]
    )
  );

  return corpus;
}

export * from "./models";
export * from "./defects";

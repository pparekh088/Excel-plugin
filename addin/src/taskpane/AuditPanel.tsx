import { useCallback, useMemo, useState } from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Body1,
  Button,
  Caption1,
  Divider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  ProgressBar,
  Spinner,
  Text,
  Title3,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  DependencyGraph,
  runAudit,
  type AuditReport,
  type Finding,
  type Severity,
  type Workbook,
} from "ledger-engine";
import { extractWorkbook, type ExtractProgress } from "../excel/extract";
import { clearHighlight, highlightTrace, navigateTo } from "../excel/highlight";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", rowGap: "12px" },
  row: { display: "flex", alignItems: "center", columnGap: "8px", flexWrap: "wrap" },
  scoreRow: { display: "flex", alignItems: "baseline", columnGap: "12px" },
  score: { fontSize: "40px", fontWeight: 700, lineHeight: "44px" },
  finding: {
    display: "flex",
    flexDirection: "column",
    rowGap: "4px",
    paddingBottom: "8px",
  },
  evidence: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: "11px",
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "4px 6px",
    borderRadius: tokens.borderRadiusSmall,
    overflowX: "auto",
    whiteSpace: "pre",
  },
  muted: { color: tokens.colorNeutralForeground3 },
});

const SEVERITY_COLOR: Record<Severity, "danger" | "warning" | "important" | "informative"> = {
  critical: "danger",
  high: "danger",
  medium: "warning",
  low: "important",
  info: "informative",
};

const BAND_LABEL: Record<string, string> = {
  healthy: "Healthy",
  "needs-attention": "Needs attention",
  "at-risk": "At risk",
  critical: "Critical",
};

type State =
  | { kind: "idle" }
  | { kind: "extracting"; progress: ExtractProgress }
  | { kind: "auditing" }
  | { kind: "done"; report: AuditReport; workbook: Workbook; extractMs: number; auditMs: number }
  | { kind: "error"; message: string };

interface Props {
  hostReady: boolean;
}

export function AuditPanel({ hostReady }: Props) {
  const styles = useStyles();
  const [state, setState] = useState<State>({ kind: "idle" });
  const [tracing, setTracing] = useState<string | null>(null);

  const run = useCallback(async () => {
    try {
      setState({
        kind: "extracting",
        progress: { phase: "sheets", sheetIndex: 0, sheetCount: 0, cellsRead: 0 },
      });
      const extractStarted = performance.now();
      const workbook = await extractWorkbook({
        includeNumberFormats: true,
        onProgress: (progress) => setState({ kind: "extracting", progress }),
      });
      const extractMs = Math.round(performance.now() - extractStarted);

      setState({ kind: "auditing" });
      const auditStarted = performance.now();
      // Runs entirely in-process: no backend, no LLM, no network.
      const graph = DependencyGraph.build(workbook);
      const report = runAudit(workbook, { graph });
      const auditMs = Math.round(performance.now() - auditStarted);

      setState({ kind: "done", report, workbook, extractMs, auditMs });
    } catch (error) {
      setState({
        kind: "error",
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }, []);

  const trace = useCallback(async (finding: Finding) => {
    setTracing(finding.address);
    try {
      const primary = { sheet: finding.sheet, row: finding.row, col: finding.col };
      await highlightTrace(primary, finding.trace);
      await navigateTo(primary);
    } finally {
      setTracing(null);
    }
  }, []);

  const grouped = useMemo(() => {
    if (state.kind !== "done") return [];
    const order: Severity[] = ["critical", "high", "medium", "low", "info"];
    return order
      .map((severity) => ({
        severity,
        findings: state.report.findings.filter((finding) => finding.severity === severity),
      }))
      .filter((group) => group.findings.length > 0);
  }, [state]);

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Button appearance="primary" disabled={!hostReady || state.kind === "extracting" || state.kind === "auditing"} onClick={() => void run()}>
          {state.kind === "done" ? "Re-run audit" : "Run audit"}
        </Button>
        {state.kind === "done" && (
          <Button onClick={() => void clearHighlight()}>Clear highlights</Button>
        )}
        <Caption1 className={styles.muted}>Runs locally — no data leaves Excel.</Caption1>
      </div>

      {!hostReady && (
        <MessageBar intent="warning">
          <MessageBarBody>Open this add-in inside Excel to audit a workbook.</MessageBarBody>
        </MessageBar>
      )}

      {state.kind === "extracting" && (
        <>
          <ProgressBar />
          <Caption1>
            Reading {state.progress.sheetName ?? "workbook"} — sheet{" "}
            {state.progress.sheetIndex}/{state.progress.sheetCount},{" "}
            {state.progress.cellsRead.toLocaleString()} cells
          </Caption1>
        </>
      )}

      {state.kind === "auditing" && (
        <div className={styles.row}>
          <Spinner size="tiny" />
          <Caption1>Analysing dependencies…</Caption1>
        </div>
      )}

      {state.kind === "error" && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Audit failed</MessageBarTitle>
            {state.message}
          </MessageBarBody>
        </MessageBar>
      )}

      {state.kind === "done" && (
        <>
          <div className={styles.scoreRow}>
            <Text className={styles.score}>{state.report.health.score}</Text>
            <div>
              <Badge
                appearance="filled"
                color={state.report.health.score >= 85 ? "success" : state.report.health.score >= 65 ? "warning" : "danger"}
              >
                {BAND_LABEL[state.report.health.band]}
              </Badge>
              <Caption1 className={styles.muted}>
                {" "}
                {state.report.findings.length} issue
                {state.report.findings.length === 1 ? "" : "s"} ·{" "}
                {state.report.coverage.formulaCells.toLocaleString()} formulas
              </Caption1>
            </div>
          </div>

          <Caption1 className={styles.muted}>
            Extracted in {state.extractMs}ms, audited in {state.auditMs}ms,{" "}
            {state.report.stats.llmCallsUsed} LLM calls.
          </Caption1>

          {!state.report.coverage.complete && (
            <MessageBar intent="warning">
              <MessageBarBody>
                <MessageBarTitle>Coverage is incomplete</MessageBarTitle>
                {state.report.coverage.caveat}
              </MessageBarBody>
            </MessageBar>
          )}

          {state.report.findings.length === 0 && (
            <MessageBar intent="success">
              <MessageBarBody>
                No issues found by {state.report.rulesRun.length} rules. Coverage:{" "}
                {state.report.coverage.caveat}
              </MessageBarBody>
            </MessageBar>
          )}

          <Divider />

          <Accordion multiple collapsible defaultOpenItems={grouped[0]?.severity}>
            {grouped.map((group) => (
              <AccordionItem key={group.severity} value={group.severity}>
                <AccordionHeader>
                  <Badge appearance="filled" color={SEVERITY_COLOR[group.severity]}>
                    {group.severity}
                  </Badge>
                  <Text>&nbsp;{group.findings.length}</Text>
                </AccordionHeader>
                <AccordionPanel>
                  {group.findings.map((finding) => (
                    <div key={finding.address + finding.rule} className={styles.finding}>
                      <div className={styles.row}>
                        <Text weight="semibold">{finding.address}</Text>
                        <Caption1 className={styles.muted}>
                          {finding.rule} · {finding.title}
                        </Caption1>
                      </div>
                      <Body1>{finding.explanation}</Body1>
                      {finding.alsoFlaggedBy && finding.alsoFlaggedBy.length > 0 && (
                        <Caption1 className={styles.muted}>
                          Also flagged by{" "}
                          {finding.alsoFlaggedBy.map((item) => `${item.rule} (${item.title})`).join(", ")}
                        </Caption1>
                      )}
                      {finding.evidence && <div className={styles.evidence}>{finding.evidence}</div>}
                      <div className={styles.row}>
                        <Button size="small" onClick={() => void trace(finding)}>
                          {tracing === finding.address ? "Tracing…" : "Trace"}
                        </Button>
                        {finding.blastRadius > 0 && (
                          <Caption1 className={styles.muted}>
                            {finding.blastRadius} cell{finding.blastRadius === 1 ? "" : "s"} downstream
                          </Caption1>
                        )}
                        {finding.autoFix && (
                          <Caption1 className={styles.muted}>
                            Fix available ({finding.autoFix.risk} risk)
                          </Caption1>
                        )}
                      </div>
                    </div>
                  ))}
                </AccordionPanel>
              </AccordionItem>
            ))}
          </Accordion>
        </>
      )}

      {state.kind === "idle" && hostReady && (
        <Body1 className={styles.muted}>
          Reads every formula in this workbook, builds a dependency graph, and checks{" "}
          11 model-integrity rules. Entirely local and deterministic.
        </Body1>
      )}
    </div>
  );
}

export function auditTitle(): JSX.Element {
  return <Title3>Audit</Title3>;
}

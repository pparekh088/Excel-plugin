import { useEffect, useState } from "react";
import {
  Badge,
  Caption1,
  Divider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", rowGap: "12px" },
  total: { fontSize: "28px", fontWeight: 700 },
  muted: { color: tokens.colorNeutralForeground3 },
});

interface TierCost {
  inputTokens: number;
  outputTokens: number;
  cachedInputTokens: number;
  costUsd: number;
  model: string;
}

interface CostSnapshot {
  total_cost_usd: number;
  by_tier: Record<string, TierCost>;
}

interface BudgetEvent {
  used: number;
  budget: number;
}

interface Props {
  backendUrl: string;
  sessionId: string | null;
}

/** Formats sub-cent amounts usefully rather than showing "$0.00". */
function formatUsd(amount: number): string {
  if (amount === 0) return "$0";
  if (amount < 0.01) return `$${amount.toFixed(4)}`;
  return `$${amount.toFixed(2)}`;
}

export function CostPanel({ backendUrl, sessionId }: Props) {
  const styles = useStyles();
  const [snapshot, setSnapshot] = useState<CostSnapshot | null>(null);
  const [budgetEvent, setBudgetEvent] = useState<BudgetEvent | null>(null);
  const [error, setError] = useState<string | null>(null);

  // The custom-functions runtime raises this when the breaker fires; without
  // it a user sees #AI_BUDGET! in the grid with nothing explaining it.
  useEffect(() => {
    const handler = (event: Event): void => {
      const detail = (event as CustomEvent<BudgetEvent>).detail;
      if (detail) setBudgetEvent(detail);
    };
    window.addEventListener("ledger:ai-budget-exceeded", handler);
    return () => window.removeEventListener("ledger:ai-budget-exceeded", handler);
  }, []);

  useEffect(() => {
    if (!sessionId) return;
    let cancelled = false;

    const poll = async (): Promise<void> => {
      try {
        const response = await fetch(
          `${backendUrl}/api/v1/sessions/${sessionId}/agent/cost`
        );
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const body = (await response.json()) as CostSnapshot;
        if (!cancelled) {
          setSnapshot(body);
          setError(null);
        }
      } catch (exc) {
        if (!cancelled) setError(exc instanceof Error ? exc.message : String(exc));
      }
    };

    void poll();
    const timer = setInterval(() => void poll(), 5000);
    return () => {
      cancelled = true;
      clearInterval(timer);
    };
  }, [backendUrl, sessionId]);

  const tiers = Object.entries(snapshot?.by_tier ?? {});
  const totalCached = tiers.reduce((sum, [, tier]) => sum + tier.cachedInputTokens, 0);
  const totalInput = tiers.reduce((sum, [, tier]) => sum + tier.inputTokens, 0);

  return (
    <div className={styles.root}>
      {budgetEvent && (
        <MessageBar intent="warning">
          <MessageBarBody>
            <MessageBarTitle>AI budget reached</MessageBarTitle>
            {budgetEvent.used} of {budgetEvent.budget} allowed calls were used in this
            recalculation, so the remaining AI.* cells returned #AI_BUDGET! rather than
            continuing to spend. Recalculate to start a new cycle, or reduce how many distinct
            values the formula has to classify.
          </MessageBarBody>
        </MessageBar>
      )}

      {!sessionId && (
        <Caption1 className={styles.muted}>
          Connect to the backend to see session cost.
        </Caption1>
      )}

      {error && (
        <Caption1 className={styles.muted}>Cost unavailable: {error}</Caption1>
      )}

      {snapshot && (
        <>
          <div>
            <Text className={styles.total}>{formatUsd(snapshot.total_cost_usd)}</Text>{" "}
            <Caption1 className={styles.muted}>this session</Caption1>
          </div>

          {totalInput > 0 && (
            <Caption1 className={styles.muted}>
              {totalCached.toLocaleString()} of {totalInput.toLocaleString()} input tokens came
              from cache
              {totalCached > 0
                ? ` (${((totalCached / totalInput) * 100).toFixed(0)}% saved)`
                : ""}
              .
            </Caption1>
          )}

          <Divider>By model tier</Divider>
          {tiers.length === 0 ? (
            <Caption1 className={styles.muted}>No model calls yet.</Caption1>
          ) : (
            <Table size="extra-small">
              <TableHeader>
                <TableRow>
                  <TableHeaderCell>Tier</TableHeaderCell>
                  <TableHeaderCell>Model</TableHeaderCell>
                  <TableHeaderCell>In</TableHeaderCell>
                  <TableHeaderCell>Out</TableHeaderCell>
                  <TableHeaderCell>Cost</TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {tiers.map(([tier, usage]) => (
                  <TableRow key={tier}>
                    <TableCell>
                      <Badge appearance="outline">{tier}</Badge>
                    </TableCell>
                    <TableCell>{usage.model}</TableCell>
                    <TableCell>{usage.inputTokens.toLocaleString()}</TableCell>
                    <TableCell>{usage.outputTokens.toLocaleString()}</TableCell>
                    <TableCell>{formatUsd(usage.costUsd)}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          )}
        </>
      )}
    </div>
  );
}

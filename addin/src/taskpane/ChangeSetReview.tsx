import {
  Badge,
  Body1,
  Button,
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
import type { ChangeSet, DiffEntry, RiskTier } from "ledger-engine";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", rowGap: "12px" },
  row: { display: "flex", alignItems: "center", columnGap: "8px", flexWrap: "wrap" },
  formula: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: "11px",
    whiteSpace: "nowrap",
  },
  before: { color: tokens.colorPaletteRedForeground1 },
  after: { color: tokens.colorPaletteGreenForeground1 },
  muted: { color: tokens.colorNeutralForeground3 },
  scroll: { maxHeight: "300px", overflowY: "auto" },
});

const RISK_COLOR: Record<RiskTier, "danger" | "warning" | "success"> = {
  high: "danger",
  medium: "warning",
  low: "success",
};

const RISK_EXPLANATION: Record<RiskTier, string> = {
  high: "Overwrites existing formulas or data, or deletes something. Review every row below.",
  medium: "Writes into empty cells or adds structure. Nothing existing is overwritten.",
  low: "Formatting only. Nothing that affects a calculation.",
};

function describeBefore(entry: DiffEntry): string {
  if (entry.before.absent) return "(empty)";
  if (entry.before.formula) return entry.before.formula;
  return String(entry.before.value ?? "(empty)");
}

function describeAfter(entry: DiffEntry): string {
  if (entry.after.cleared) return "(cleared)";
  if (entry.after.formula !== undefined) return entry.after.formula;
  if (entry.after.value !== undefined) return String(entry.after.value);
  if (entry.after.numberFormat !== undefined) return `format: ${entry.after.numberFormat}`;
  return "(unchanged)";
}

interface Props {
  changeSet: ChangeSet;
  explanation: string;
  busy?: boolean;
  onApprove: () => void;
  onReject: () => void;
}

export function ChangeSetReview({ changeSet, explanation, busy, onApprove, onReject }: Props) {
  const styles = useStyles();
  const overwrites = changeSet.diff.filter((entry) => entry.overwritesFormula);

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Text weight="semibold">{changeSet.summary}</Text>
        <Badge appearance="filled" color={RISK_COLOR[changeSet.risk]}>
          {changeSet.risk} risk
        </Badge>
      </div>
      <Caption1 className={styles.muted}>{RISK_EXPLANATION[changeSet.risk]}</Caption1>

      {overwrites.length > 0 && (
        <MessageBar intent="warning">
          <MessageBarBody>
            <MessageBarTitle>
              {overwrites.length} existing formula
              {overwrites.length === 1 ? "" : "s"} will be replaced
            </MessageBarTitle>
            {overwrites
              .slice(0, 4)
              .map((entry) => entry.address)
              .join(", ")}
            {overwrites.length > 4 ? `, and ${overwrites.length - 4} more` : ""}
          </MessageBarBody>
        </MessageBar>
      )}

      <Divider>Downstream impact</Divider>
      <Body1>
        {changeSet.impact.affectedCells} cell
        {changeSet.impact.affectedCells === 1 ? "" : "s"} will recalculate across{" "}
        {changeSet.impact.affectedNodes} block
        {changeSet.impact.affectedNodes === 1 ? "" : "s"}.
      </Body1>
      {changeSet.impact.affectedOutputs.length > 0 && (
        <Caption1 className={styles.muted}>
          Outputs affected: {changeSet.impact.affectedOutputs.slice(0, 6).join(", ")}
        </Caption1>
      )}
      {changeSet.impact.affectedPivots.length > 0 && (
        <Caption1 className={styles.muted}>
          Pivots on touched sheets ({changeSet.impact.affectedPivots.join(", ")}) may need a
          manual refresh — the API cannot refresh all pivot types.
        </Caption1>
      )}
      {changeSet.impact.opaqueDownstream > 0 && (
        <MessageBar intent="warning">
          <MessageBarBody>
            {changeSet.impact.opaqueDownstream} downstream block(s) use INDIRECT, OFFSET or
            external links, so the impact figures above are a lower bound — there may be more.
          </MessageBarBody>
        </MessageBar>
      )}

      <Divider>Changes ({changeSet.diff.length})</Divider>
      <div className={styles.scroll}>
        <Table size="extra-small">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Cell</TableHeaderCell>
              <TableHeaderCell>Before</TableHeaderCell>
              <TableHeaderCell>After</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {changeSet.diff.slice(0, 200).map((entry) => (
              <TableRow key={entry.address}>
                <TableCell>
                  <Text className={styles.formula}>{entry.address}</Text>
                </TableCell>
                <TableCell>
                  <Text className={`${styles.formula} ${styles.before}`}>
                    {describeBefore(entry)}
                  </Text>
                </TableCell>
                <TableCell>
                  <Text className={`${styles.formula} ${styles.after}`}>
                    {describeAfter(entry)}
                  </Text>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
        {changeSet.diff.length > 200 && (
          <Caption1 className={styles.muted}>
            Showing the first 200 of {changeSet.diff.length} changes.
          </Caption1>
        )}
      </div>

      <Divider>Explanation</Divider>
      <Caption1 className={styles.muted} style={{ whiteSpace: "pre-wrap" }}>
        {explanation}
      </Caption1>

      <div className={styles.row}>
        <Button appearance="primary" disabled={busy} onClick={onApprove}>
          {busy ? "Applying…" : "Apply change set"}
        </Button>
        <Button disabled={busy} onClick={onReject}>
          Discard
        </Button>
        <Caption1 className={styles.muted}>
          Nothing has been written yet. Applying re-checks for other people's edits first.
        </Caption1>
      </div>
    </div>
  );
}

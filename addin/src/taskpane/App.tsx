import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Badge,
  Button,
  Checkbox,
  Divider,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Spinner,
  Tab,
  TabList,
  Text,
  Title3,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { AuditPanel } from "./AuditPanel";
import { CostPanel } from "./CostPanel";
import { APP_NAME, PHASE, getBackendUrl } from "../config";
import { createAuthProvider } from "../auth/provider";
import { ApiError, LedgerClient, type Health } from "../api/client";
import { setActiveSessionId } from "../api/backend";
import { runTool, type RunToolOutcome } from "../api/runTool";
import type { RangeReadInclude } from "./types";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: "12px",
    padding: "16px",
    maxWidth: "480px",
  },
  header: { display: "flex", alignItems: "center", columnGap: "8px" },
  row: { display: "flex", alignItems: "center", columnGap: "8px", flexWrap: "wrap" },
  output: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    padding: "8px",
    maxHeight: "320px",
    overflow: "auto",
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: "11px",
    whiteSpace: "pre",
  },
  muted: { color: tokens.colorNeutralForeground3 },
});

type BackendState =
  | { kind: "checking" }
  | { kind: "online"; health: Health; sessionId: string }
  | { kind: "offline"; message: string };

interface Props {
  hostReady: boolean;
}

const ALL_INCLUDES: RangeReadInclude[] = ["values", "formulas", "numberFormats"];

function formatOutcome(outcome: RunToolOutcome): string {
  return JSON.stringify(
    {
      tool_call_id: outcome.toolCallId,
      ack: outcome.ack,
      timings_ms: outcome.timings,
      ...(outcome.error ? { error: outcome.error } : {}),
      ...(outcome.result ? { result: outcome.result } : {}),
    },
    null,
    2
  );
}

function describeError(exc: unknown): string {
  if (exc instanceof ApiError) {
    return `HTTP ${exc.status}: ${JSON.stringify(exc.body, null, 2)}`;
  }
  return exc instanceof Error ? exc.message : String(exc);
}

export function App({ hostReady }: Props) {
  const styles = useStyles();
  const backendUrl = getBackendUrl();
  const client = useMemo(() => new LedgerClient(backendUrl, createAuthProvider()), [backendUrl]);

  const [backend, setBackend] = useState<BackendState>({ kind: "checking" });
  const [sheet, setSheet] = useState("Sheet1");
  const [a1, setA1] = useState("A1:D10");
  const [include, setInclude] = useState<RangeReadInclude[]>(["values", "formulas"]);
  const [running, setRunning] = useState(false);
  const [output, setOutput] = useState<string>("");
  const [tab, setTab] = useState<"audit" | "tools" | "cost">("audit");

  const connect = useCallback(async () => {
    setBackend({ kind: "checking" });
    try {
      const health = await client.health();
      const session = await client.createSession();
      // Publish the session so AI.* batches (shared runtime) meter their
      // spend against this workbook's session rather than the catch-all.
      setActiveSessionId(session.session_id);
      setBackend({ kind: "online", health, sessionId: session.session_id });
    } catch (exc) {
      setActiveSessionId(null);
      setBackend({ kind: "offline", message: describeError(exc) });
    }
  }, [client]);

  useEffect(() => {
    void connect();
  }, [connect]);

  // Prefill the sheet input with the active worksheet's name.
  useEffect(() => {
    if (!hostReady) return;
    void Excel.run(async (context) => {
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await context.sync();
      setSheet(active.name);
    }).catch(() => {
      /* leave the default */
    });
  }, [hostReady]);

  const toggleInclude = (value: RangeReadInclude, checked: boolean) => {
    setInclude((current) => {
      const next = checked ? [...current, value] : current.filter((item) => item !== value);
      return next.length > 0 ? next : current;
    });
  };

  const run = async () => {
    if (backend.kind !== "online") return;
    setRunning(true);
    setOutput("");
    try {
      const outcome = await runTool(client, backend.sessionId, "range.read", {
        sheet,
        a1,
        include,
      });
      setOutput(formatOutcome(outcome));
    } catch (exc) {
      setOutput(describeError(exc));
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title3>{APP_NAME}</Title3>
        <Badge appearance="outline">{PHASE}</Badge>
      </div>

      <TabList selectedValue={tab} onTabSelect={(_, data) => setTab(data.value as "audit" | "tools" | "cost")}>
        <Tab value="audit">Audit</Tab>
        <Tab value="tools">Tools</Tab>
        <Tab value="cost">Cost</Tab>
      </TabList>

      {tab === "audit" && <AuditPanel hostReady={hostReady} />}

      {tab === "cost" && (
        <CostPanel
          backendUrl={backendUrl}
          sessionId={backend.kind === "online" ? backend.sessionId : null}
        />
      )}

      {tab === "tools" && (
        <>
      {!hostReady && (
        <MessageBar intent="warning">
          <MessageBarBody>
            Not running inside Excel — sideload the manifest to run the tool round trip. Backend
            status still works below.
          </MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.row}>
        <Text weight="semibold">Backend</Text>
        {backend.kind === "checking" && <Spinner size="tiny" />}
        {backend.kind === "online" && (
          <>
            <Badge appearance="filled" color="success">
              online
            </Badge>
            <Text className={styles.muted}>
              v{backend.health.version} · {backend.health.session_store} · auth=
              {backend.health.auth_mode}
            </Text>
          </>
        )}
        {backend.kind === "offline" && (
          <Badge appearance="filled" color="danger">
            offline
          </Badge>
        )}
        <Button size="small" onClick={() => void connect()}>
          Reconnect
        </Button>
      </div>

      {backend.kind === "offline" && (
        <MessageBar intent="error">
          <MessageBarBody>
            Backend unreachable at {backendUrl} — agent features are disabled. Local-only features
            keep working. ({backend.message})
          </MessageBarBody>
        </MessageBar>
      )}

      {backend.kind === "online" && (
        <Text className={styles.muted}>
          session {backend.sessionId} · tools: {backend.health.tools.join(", ")}
        </Text>
      )}

      <Divider>range.read round trip</Divider>

      <Field label="Sheet">
        <Input value={sheet} onChange={(_, data) => setSheet(data.value)} />
      </Field>
      <Field label="Range (A1)">
        <Input value={a1} onChange={(_, data) => setA1(data.value)} />
      </Field>
      <div className={styles.row}>
        {ALL_INCLUDES.map((value) => (
          <Checkbox
            key={value}
            label={value}
            checked={include.includes(value)}
            onChange={(_, data) => toggleInclude(value, data.checked === true)}
          />
        ))}
      </div>
      <div className={styles.row}>
        <Button
          appearance="primary"
          disabled={!hostReady || backend.kind !== "online" || running}
          onClick={() => void run()}
        >
          {running ? "Running…" : "Run range.read"}
        </Button>
        {running && <Spinner size="tiny" />}
      </div>

      {output && <div className={styles.output}>{output}</div>}
        </>
      )}
    </div>
  );
}

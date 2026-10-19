import { Badge, makeStyles, Spinner, Tab, TabList, Text, tokens } from "@fluentui/react-components";
import * as React from "react";
import { AppState, defaultState } from "../excel/model";
import { loadState, saveState } from "../excel/storage";
import { describeError, excelApi, MIN_EXCEL_API } from "../excel/support";
import { AnalysisTab } from "./components/AnalysisTab";
import { BuildTab } from "./components/BuildTab";
import { Status, StatusMessage } from "./components/common";
import { DistributionsTab } from "./components/DistributionsTab";

const useAppStyles = makeStyles({
  root: { padding: "8px 12px 24px", display: "flex", flexDirection: "column", gap: "8px" },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px" },
  busy: { display: "flex", alignItems: "center", gap: "8px", color: tokens.colorNeutralForeground2 },
});

type TabId = "build" | "list" | "analysis";

export interface Ops {
  busy: boolean;
  /** Run a workbook operation with a busy indicator and uniform error reporting. */
  run: <T>(label: string, op: () => Promise<T>, done?: (result: T) => StatusMessage | null) => Promise<void>;
  setState: (next: AppState) => void;
  updateSettings: (patch: Partial<Pick<AppState, "chartSettings" | "rangeSettings" | "outputSettings" | "analysisSettings" | "tolerance">>) => void;
  setMessage: (m: StatusMessage | null) => void;
}

export function App() {
  const styles = useAppStyles();
  const [state, setState] = React.useState<AppState | null>(null);
  const [busyLabel, setBusyLabel] = React.useState<string | null>(null);
  const [message, setMessage] = React.useState<StatusMessage | null>(null);
  const [tab, setTab] = React.useState<TabId>("build");
  const saveTimer = React.useRef<number | undefined>(undefined);
  const supported = excelApi(MIN_EXCEL_API);

  React.useEffect(() => {
    loadState()
      .then(setState)
      .catch((error) => {
        setState(defaultState());
        setMessage({ intent: "warning", text: `Saved settings could not be loaded: ${describeError(error)}` });
      });
  }, []);

  const updateSettings: Ops["updateSettings"] = (patch) => {
    setState((prev) => {
      if (!prev) return prev;
      const next = { ...prev, ...patch };
      window.clearTimeout(saveTimer.current);
      saveTimer.current = window.setTimeout(() => {
        saveState(next).catch(() => undefined);
      }, 600);
      return next;
    });
  };

  const run: Ops["run"] = async (label, op, done) => {
    setBusyLabel(label);
    setMessage(null);
    try {
      const result = await op();
      const msg = done ? done(result) : null;
      if (msg) setMessage(msg);
    } catch (error) {
      setMessage({ intent: "error", title: `${label} failed`, text: describeError(error) });
    } finally {
      setBusyLabel(null);
    }
  };

  if (!state) {
    return (
      <div className={styles.root}>
        <Spinner label="Loading…" size="small" />
      </div>
    );
  }

  const ops: Ops = { busy: busyLabel !== null, run, setState, updateSettings, setMessage };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text weight="semibold" size={400}>
          Probability Step Chart
        </Text>
        {busyLabel ? (
          <span className={styles.busy}>
            <Spinner size="extra-tiny" />
            <Text size={200}>{busyLabel}…</Text>
          </span>
        ) : null}
      </div>
      {!supported ? (
        <Status
          message={{
            intent: "error",
            title: "Unsupported Excel version",
            text: `This add-in needs Excel JavaScript API ${MIN_EXCEL_API} or later (Excel 2019+, Microsoft 365, Excel for Mac or Excel on the web).`,
          }}
        />
      ) : null}
      <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as TabId)} size="small">
        <Tab value="build">Build</Tab>
        <Tab value="list">
          Distributions{" "}
          {state.distributions.length ? <Badge size="small" appearance="tint">{state.distributions.length}</Badge> : null}
        </Tab>
        <Tab value="analysis">Analysis</Tab>
      </TabList>
      <Status message={message} />
      {tab === "build" ? <BuildTab state={state} ops={ops} /> : null}
      {tab === "list" ? <DistributionsTab state={state} ops={ops} /> : null}
      {tab === "analysis" ? <AnalysisTab state={state} ops={ops} /> : null}
    </div>
  );
}

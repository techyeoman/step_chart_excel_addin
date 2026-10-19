import { Button, Input, Link, Switch, Text } from "@fluentui/react-components";
import { ArrowClockwiseRegular, DeleteRegular } from "@fluentui/react-icons";
import * as React from "react";
import { qualify, rangeRef } from "../../core/a1";
import { computeStats } from "../../core/stats";
import { AppState, DistributionEntry } from "../../excel/model";
import { selectRange } from "../../excel/rangeIO";
import { removeDistribution, renderAll, updateEntry } from "../../excel/service";
import type { Ops } from "../App";
import { fmtNum, SectionTitle, useStyles } from "./common";

interface Props {
  state: AppState;
  ops: Ops;
}

export function DistributionsTab({ state, ops }: Props) {
  const styles = useStyles();

  function regenerate() {
    void ops.run(
      "Regenerating",
      () => renderAll(state, { reread: true }),
      (result) => {
        ops.setState(result.state);
        return {
          intent: result.notes.length ? "warning" : "success",
          text: `Re-read ${result.state.distributions.length} source range(s), rewrote the tables${
            result.state.chartSettings.createChart ? " and refreshed the chart" : ""
          }.`,
          notes: result.notes,
        };
      },
    );
  }

  if (state.distributions.length === 0) {
    return (
      <div className={styles.section}>
        <Text className={styles.muted}>
          No distributions yet. On the Build tab, select an outcome/probability range and generate its CDF. Each
          distribution you add becomes a series on the same chart.
        </Text>
      </div>
    );
  }

  return (
    <div>
      <div className={styles.section}>
        <SectionTitle>Distributions on the chart</SectionTitle>
        <Text className={styles.muted}>
          After editing source data, choose Regenerate to re-read every source range, rewrite the coordinate tables
          and refresh the chart.
        </Text>
        <Button icon={<ArrowClockwiseRegular />} onClick={regenerate} disabled={ops.busy} style={{ alignSelf: "flex-start" }}>
          Regenerate all from source data
        </Button>
      </div>
      {state.distributions.map((entry) => (
        <EntryCard key={entry.id} entry={entry} state={state} ops={ops} />
      ))}
    </div>
  );
}

function EntryCard({ entry, state, ops }: { entry: DistributionEntry; state: AppState; ops: Ops }) {
  const styles = useStyles();
  const [name, setName] = React.useState(entry.name);
  React.useEffect(() => setName(entry.name), [entry.name]);
  const stats = React.useMemo(() => {
    try {
      return computeStats({ outcomes: entry.outcomes });
    } catch {
      return null;
    }
  }, [entry.outcomes]);

  const apply = (patch: Parameters<typeof updateEntry>[2], label: string) =>
    ops.run(label, () => updateEntry(state, entry.id, patch), (result) => {
      ops.setState(result.state);
      return result.notes.length ? { intent: "warning", text: `${label} done.`, notes: result.notes } : null;
    });

  function commitName() {
    const trimmed = name.trim();
    if (!trimmed || trimmed === entry.name) {
      setName(entry.name);
      return;
    }
    void apply({ name: trimmed }, "Renaming");
  }

  function remove() {
    void ops.run(
      "Removing",
      () => removeDistribution(state, entry.id),
      (result) => {
        ops.setState(result.state);
        return {
          intent: "success",
          text: `Removed "${entry.name}" and its coordinate table. The source data was not changed.`,
          notes: result.notes,
        };
      },
    );
  }

  const out = entry.output;
  const outAddress = out ? qualify(out.sheetName, rangeRef(out.row, out.col, out.rows, out.cols)) : null;

  return (
    <div className={styles.section}>
      <div className={styles.card}>
        <Input
          value={name}
          onChange={(_, d) => setName(d.value)}
          onBlur={commitName}
          onKeyDown={(e) => e.key === "Enter" && commitName()}
          disabled={ops.busy}
          aria-label="Distribution name"
        />
        <Text className={styles.muted}>
          Source:{" "}
          <Link
            onClick={() => void ops.run("Selecting", () => selectRange(entry.source.sheetId, entry.source.address))}
          >
            {qualify(entry.source.sheetName, entry.source.address)}
          </Link>
        </Text>
        {out && outAddress ? (
          <Text className={styles.muted}>
            Output:{" "}
            <Link
              onClick={() =>
                void ops.run("Selecting", () => selectRange(out.sheetId, rangeRef(out.row, out.col, out.rows, out.cols)))
              }
            >
              {outAddress}
            </Link>{" "}
            ({out.tableName ? `table ${out.tableName}` : `name ${out.definedName}`}, {out.dataCount} points)
          </Text>
        ) : null}
        {stats ? (
          <Text className={styles.muted}>
            {stats.outcomes} outcomes · EV {fmtNum(stats.expectedValue)} · SD {fmtNum(stats.standardDeviation)}
          </Text>
        ) : null}
        <div className={styles.row} style={{ justifyContent: "space-between", alignItems: "center" }}>
          <Switch
            checked={entry.includeInChart}
            label="Show on chart"
            disabled={ops.busy}
            onChange={(_, d) => void apply({ includeInChart: d.checked }, d.checked ? "Adding to chart" : "Hiding from chart")}
          />
          <Button appearance="subtle" icon={<DeleteRegular />} onClick={remove} disabled={ops.busy}>
            Remove
          </Button>
        </div>
      </div>
    </div>
  );
}

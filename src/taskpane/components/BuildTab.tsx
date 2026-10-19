import {
  Button,
  Checkbox,
  Field,
  Input,
  Radio,
  RadioGroup,
  Select,
  Text,
} from "@fluentui/react-components";
import { ArrowSyncRegular, CursorClickRegular, TableSimpleRegular } from "@fluentui/react-icons";
import * as React from "react";
import { sanitizeExcelName } from "../../core/a1";
import { buildStepCoordinates, cumulativeProbabilities, EndMode, resolveXRange, StartMode } from "../../core/cdf";
import { HeaderMode, parseDistribution, ProbabilityFormat } from "../../core/parse";
import { computeStats } from "../../core/stats";
import { AppState, Placement } from "../../excel/model";
import { CellSnapshot, getActiveSheet, RangeSnapshot, readSelectedCell, readSelection } from "../../excel/rangeIO";
import { addOrUpdate, toXRangeOptions } from "../../excel/service";
import { UserError } from "../../excel/support";
import type { Ops } from "../App";
import { fmtNum, fmtPct, IssueList, SectionTitle, useStyles } from "./common";

interface Props {
  state: AppState;
  ops: Ops;
}

export function BuildTab({ state, ops }: Props) {
  const styles = useStyles();
  const [snapshot, setSnapshot] = React.useState<RangeSnapshot | null>(null);
  const [headerMode, setHeaderMode] = React.useState<HeaderMode>("auto");
  const [probabilityFormat, setProbabilityFormat] = React.useState<ProbabilityFormat>("auto");
  const [name, setName] = React.useState("");
  const [outputName, setOutputName] = React.useState("");
  const [outputNameTouched, setOutputNameTouched] = React.useState(false);
  const [cellTarget, setCellTarget] = React.useState<CellSnapshot | null>(null);
  const [showSteps, setShowSteps] = React.useState(false);

  const { chartSettings: chart, rangeSettings: range, outputSettings: output } = state;

  const parse = React.useMemo(
    () =>
      snapshot
        ? parseDistribution(snapshot.values, {
            headerMode,
            probabilityFormat,
            tolerance: state.tolerance,
            origin: { rowIndex: snapshot.rowIndex, columnIndex: snapshot.columnIndex },
          })
        : null,
    [snapshot, headerMode, probabilityFormat, state.tolerance],
  );

  // Step coordinates preview using the shared X range of every charted distribution.
  const preview = React.useMemo(() => {
    if (!parse?.distribution || !snapshot) return null;
    const others = state.distributions
      .filter((e) => !(e.source.sheetId === snapshot.source.sheetId && e.source.address === snapshot.source.address))
      .map((e) => ({ outcomes: e.outcomes }));
    try {
      const xr = resolveXRange([...others, parse.distribution], toXRangeOptions(range));
      return { points: buildStepCoordinates(parse.distribution, { ...xr, tolerance: state.tolerance }), error: null };
    } catch (error) {
      return { points: [], error: (error as Error).message };
    }
  }, [parse, snapshot, state.distributions, range, state.tolerance]);

  const existing = snapshot
    ? state.distributions.find(
        (e) => e.source.sheetId === snapshot.source.sheetId && e.source.address === snapshot.source.address,
      )
    : undefined;

  function loadSelection() {
    void ops.run("Reading selection", readSelection, (snap) => {
      setSnapshot(snap);
      const match = state.distributions.find(
        (e) => e.source.sheetId === snap.source.sheetId && e.source.address === snap.source.address,
      );
      const defaultName = match?.name ?? snap.labelAbove ?? `Distribution ${state.distributions.length + 1}`;
      setName(defaultName);
      setHeaderMode(match?.headerMode ?? "auto");
      setProbabilityFormat(match?.probabilityFormat ?? "auto");
      setOutputName(match?.outputName ?? `CDF_${sanitizeExcelName(defaultName)}`);
      setOutputNameTouched(!!match);
      return snap.expanded
        ? { intent: "info", text: `Expanded the selected cell to the surrounding data ${snap.displayAddress}.` }
        : null;
    });
  }

  function onNameChange(value: string) {
    setName(value);
    if (!outputNameTouched) setOutputName(`CDF_${sanitizeExcelName(value)}`);
  }

  function pickCell() {
    void ops.run("Reading output cell", readSelectedCell, (cell) => {
      setCellTarget(cell);
      return null;
    });
  }

  async function placement(): Promise<Placement> {
    switch (output.location) {
      case "newSheet":
        return { kind: "newSheet", sheetName: output.newSheetName };
      case "cell":
        if (!cellTarget) throw new UserError('Choose "Use selected cell" to pick the output cell first.');
        return { kind: "cell", sheetId: cellTarget.sheetId, row: cellTarget.row, col: cellTarget.col };
      default: {
        const sheet = await getActiveSheet();
        return { kind: "active", sheetId: sheet.id };
      }
    }
  }

  function generate() {
    if (!snapshot || !parse) return;
    void ops.run(
      existing ? "Updating CDF" : "Generating CDF",
      async () =>
        addOrUpdate(state, {
          snapshot,
          parse,
          name,
          outputName,
          headerMode,
          probabilityFormat,
          includeHeaders: output.includeHeaders,
          placement: await placement(),
        }),
      (result) => {
        ops.setState(result.state);
        const entry = result.state.distributions.find(
          (e) => e.source.sheetId === snapshot.source.sheetId && e.source.address === snapshot.source.address,
        );
        const where = entry?.output
          ? ` in ${entry.output.tableName ?? entry.output.definedName} on "${entry.output.sheetName}"`
          : "";
        const charted = state.chartSettings.createChart ? " and updated the chart" : "";
        return {
          intent: result.notes.length ? "warning" : "success",
          text: `Wrote ${entry?.output?.dataCount ?? 0} step coordinates for "${entry?.name}"${where}${charted}.`,
          notes: result.notes,
        };
      },
    );
  }

  const stats = parse?.distribution ? computeStats(parse.distribution) : null;
  const cdf = parse?.distribution ? cumulativeProbabilities(parse.distribution, state.tolerance) : [];

  return (
    <div>
      {/* ---------------- Input ---------------- */}
      <div className={styles.section}>
        <SectionTitle>Input</SectionTitle>
        <Text className={styles.muted}>
          Select two columns, outcome (X) and probability, with or without a header row. Then choose Use selection.
        </Text>
        <div className={styles.row}>
          <Button appearance="primary" icon={<CursorClickRegular />} onClick={loadSelection} disabled={ops.busy}>
            Use selection
          </Button>
          {snapshot ? <Text className={styles.mono}>{snapshot.displayAddress}</Text> : null}
        </div>
        {snapshot && parse ? (
          <>
            <Field label="Header row">
                <Select value={headerMode} onChange={(_, d) => setHeaderMode(d.value as HeaderMode)}>
                  <option value="auto">Auto ({parse.headerDetected ? "header found" : "no header"})</option>
                  <option value="yes">First row is a header</option>
                  <option value="no">No header row</option>
                </Select>
            </Field>
            <Field label="Probabilities">
                <Select value={probabilityFormat} onChange={(_, d) => setProbabilityFormat(d.value as ProbabilityFormat)}>
                  <option value="auto">Auto</option>
                  <option value="decimal">Decimals (0–1)</option>
                  <option value="percent">Percentage points (0–100)</option>
                </Select>
            </Field>
            <Field label="Distribution name (series name)">
              <Input value={name} onChange={(_, d) => onNameChange(d.value)} />
            </Field>
            {existing ? (
              <Text className={styles.muted}>
                This range is already in the chart as “{existing.name}”. Generating again updates it in place.
              </Text>
            ) : null}
            <IssueList issues={parse.issues} />
            {parse.distribution ? (
              <>
                <div className={styles.scroll}>
                  <table className={styles.table}>
                    <thead>
                      <tr>
                        <th>{parse.headers?.[0] || "Outcome"}</th>
                        <th>P</th>
                        <th>Cumulative</th>
                      </tr>
                    </thead>
                    <tbody>
                      {parse.distribution.outcomes.map((o, i) => (
                        <tr key={i}>
                          <td>{fmtNum(o.x)}</td>
                          <td>{fmtPct(o.p, 4)}</td>
                          <td>{fmtPct(cdf[i]?.F ?? 0, 4)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {stats ? (
                  <Text className={styles.muted}>
                    Σp = {fmtNum(parse.probabilitySum, 10)} · EV = {fmtNum(stats.expectedValue)} · SD ={" "}
                    {fmtNum(stats.standardDeviation)}
                  </Text>
                ) : null}
              </>
            ) : null}
          </>
        ) : null}
      </div>

      {/* ---------------- CDF range ---------------- */}
      <div className={styles.section}>
        <SectionTitle>Step coordinates</SectionTitle>
        <div className={styles.row}>
          <Field
            label="Start X (probability 0)"
            className={styles.grow}
            hint={range.startMode === "auto" ? "0, or a round number below a negative minimum." : undefined}
          >
            <Select
              value={range.startMode}
              onChange={(_, d) => ops.updateSettings({ rangeSettings: { ...range, startMode: d.value as StartMode } })}
            >
              <option value="auto">Automatic</option>
              <option value="zero">Zero</option>
              <option value="min">Chart minimum X</option>
              <option value="custom">Custom value</option>
            </Select>
          </Field>
          {range.startMode === "custom" ? (
            <Field label="Start X" style={{ width: 90 }}>
              <Input
                type="number"
                value={range.customStart}
                onChange={(_, d) => ops.updateSettings({ rangeSettings: { ...range, customStart: d.value } })}
              />
            </Field>
          ) : null}
        </div>
        <div className={styles.row}>
          <Field label="Extend final segment (probability 1) to" className={styles.grow}>
            <Select
              value={range.endMode}
              onChange={(_, d) => ops.updateSettings({ rangeSettings: { ...range, endMode: d.value as EndMode } })}
            >
              <option value="max">Chart maximum X</option>
              <option value="none">Do not extend</option>
              <option value="custom">Custom value</option>
            </Select>
          </Field>
          {range.endMode === "custom" ? (
            <Field label="End X" style={{ width: 90 }}>
              <Input
                type="number"
                value={range.customEnd}
                onChange={(_, d) => ops.updateSettings({ rangeSettings: { ...range, customEnd: d.value } })}
              />
            </Field>
          ) : null}
        </div>
        <Field label="Sum-to-1 tolerance">
          <Select
            value={String(state.tolerance)}
            onChange={(_, d) => ops.updateSettings({ tolerance: Number(d.value) })}
          >
            {[1e-9, 1e-6, 1e-4, 1e-3, 1e-2].map((t) => (
              <option key={t} value={String(t)}>
                ± {t.toExponential(0).replace("e-", "e−")}
              </option>
            ))}
          </Select>
        </Field>
        {preview ? (
          preview.error ? (
            <IssueList issues={[{ severity: "error", message: preview.error }]} />
          ) : (
            <>
              <Button
                appearance="subtle"
                size="small"
                icon={<TableSimpleRegular />}
                onClick={() => setShowSteps((s) => !s)}
                style={{ alignSelf: "flex-start" }}
              >
                {showSteps ? "Hide" : "Preview"} {preview.points.length} step coordinates
              </Button>
              {showSteps ? (
                <div className={styles.scroll}>
                  <table className={styles.table}>
                    <thead>
                      <tr>
                        <th>X (Outcome)</th>
                        <th>Y (Cumulative)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {preview.points.map((p, i) => (
                        <tr key={i}>
                          <td>{fmtNum(p.x)}</td>
                          <td>{fmtNum(p.y, 10)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : null}
            </>
          )
        ) : null}
      </div>

      {/* ---------------- Output ---------------- */}
      <div className={styles.section}>
        <SectionTitle>Output</SectionTitle>
        <RadioGroup
          value={output.location}
          onChange={(_, d) =>
            ops.updateSettings({ outputSettings: { ...output, location: d.value as typeof output.location } })
          }
        >
          <Radio value="active" label="Active sheet (next to existing data)" />
          <Radio value="newSheet" label="New worksheet" />
          <Radio value="cell" label="Selected cell" />
        </RadioGroup>
        {output.location === "newSheet" ? (
          <Field label="Worksheet name" hint="Created if missing; later distributions are added beside earlier ones.">
            <Input
              value={output.newSheetName}
              onChange={(_, d) => ops.updateSettings({ outputSettings: { ...output, newSheetName: d.value } })}
            />
          </Field>
        ) : null}
        {output.location === "cell" ? (
          <div className={styles.row}>
            <Button icon={<CursorClickRegular />} onClick={pickCell} disabled={ops.busy}>
              Use selected cell
            </Button>
            <Text className={styles.mono}>{cellTarget?.displayAddress ?? "No cell chosen"}</Text>
          </div>
        ) : null}
        <Field
          label={output.includeHeaders ? "Output table name" : "Output range name"}
          hint="Letters, digits and underscores; made unique automatically."
        >
          <Input
            value={outputName}
            onChange={(_, d) => {
              setOutputName(d.value);
              setOutputNameTouched(true);
            }}
          />
        </Field>
        <Checkbox
          checked={output.includeHeaders}
          onChange={(_, d) => ops.updateSettings({ outputSettings: { ...output, includeHeaders: !!d.checked } })}
          label="Include headers (as an Excel table)"
        />
      </div>

      {/* ---------------- Chart ---------------- */}
      <div className={styles.section}>
        <SectionTitle>Chart</SectionTitle>
        <Checkbox
          checked={chart.createChart}
          onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, createChart: !!d.checked } })}
          label="Create / update chart (XY scatter with straight lines)"
        />
        {chart.createChart ? (
          <>
            <Field label="Chart title">
              <Input
                value={chart.title}
                onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, title: d.value } })}
              />
            </Field>
            <div className={styles.row}>
              <Field label="X-axis label" className={styles.grow}>
                <Input
                  value={chart.xAxisTitle}
                  onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, xAxisTitle: d.value } })}
                />
              </Field>
              <Field label="Y-axis label" className={styles.grow}>
                <Input
                  value={chart.yAxisTitle}
                  onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, yAxisTitle: d.value } })}
                />
              </Field>
            </div>
            <Checkbox
              checked={chart.yAsPercent}
              onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, yAsPercent: !!d.checked } })}
              label="Y axis as percentages (bounds fixed at 0–100%)"
            />
            <Checkbox
              checked={chart.gridlines}
              onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, gridlines: !!d.checked } })}
              label="Major gridlines"
            />
            <Checkbox
              checked={chart.markers}
              onChange={(_, d) => ops.updateSettings({ chartSettings: { ...chart, markers: !!d.checked } })}
              label="Point markers"
            />
          </>
        ) : null}
      </div>

      <div className={styles.section} style={{ borderBottom: "none" }}>
        <Button
          appearance="primary"
          icon={<ArrowSyncRegular />}
          onClick={generate}
          disabled={ops.busy || !parse?.ok || !!preview?.error}
        >
          {existing ? "Update CDF coordinates" : "Generate CDF coordinates"}
        </Button>
        {!snapshot ? <Text className={styles.muted}>Start by choosing Use selection.</Text> : null}
        {state.distributions.length > 0 && !existing && parse?.ok ? (
          <Text className={styles.muted}>
            The new distribution is added to the existing chart ({state.distributions.length} already charted).
          </Text>
        ) : null}
      </div>
    </div>
  );
}

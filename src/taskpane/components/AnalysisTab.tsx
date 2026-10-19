import { Badge, Button, Field, Radio, RadioGroup, Select, Text } from "@fluentui/react-components";
import { DocumentTableRegular } from "@fluentui/react-icons";
import * as React from "react";
import { fmt } from "../../core/numeric";
import { compareDistributions, explainDominance, Preference, Verdict } from "../../core/dominance";
import { computeStats, DistributionStats } from "../../core/stats";
import { writeAnalysisSheet } from "../../excel/analysisSheet";
import { AppState } from "../../excel/model";
import type { Ops } from "../App";
import { fmtNum, fmtPct, SectionTitle, useStyles } from "./common";

interface Props {
  state: AppState;
  ops: Ops;
}

function verdictBadge(v: Verdict, nameA: string, nameB: string) {
  if (v === "A" || v === "B") {
    return (
      <Badge appearance="filled" color="success">
        {v === "A" ? nameA : nameB} dominates
      </Badge>
    );
  }
  return (
    <Badge appearance="tint" color={v === "identical" ? "informative" : "subtle"}>
      {v === "identical" ? "Identical" : "No dominance"}
    </Badge>
  );
}

export function AnalysisTab({ state, ops }: Props) {
  const styles = useStyles();
  const dists = state.distributions;
  const [idA, setIdA] = React.useState(dists[0]?.id ?? "");
  const [idB, setIdB] = React.useState(dists[1]?.id ?? "");
  const [showTable, setShowTable] = React.useState(false);
  const preference = state.analysisSettings.preference;

  React.useEffect(() => {
    if (!dists.some((d) => d.id === idA)) setIdA(dists[0]?.id ?? "");
    if (!dists.some((d) => d.id === idB)) setIdB(dists.find((d) => d.id !== idA)?.id ?? "");
  }, [dists, idA, idB]);

  const a = dists.find((d) => d.id === idA);
  const b = dists.find((d) => d.id === idB);

  const analysis = React.useMemo(() => {
    if (!a || !b || a.id === b.id) return null;
    const statsA = computeStats({ outcomes: a.outcomes });
    const statsB = computeStats({ outcomes: b.outcomes });
    const result = compareDistributions({ outcomes: a.outcomes }, { outcomes: b.outcomes }, {
      preference,
      tolerance: Math.min(state.tolerance, 1e-9),
    });
    const explanation = explainDominance(result, a.name, b.name, { a: statsA.expectedValue, b: statsB.expectedValue });
    return { statsA, statsB, result, explanation };
  }, [a, b, preference, state.tolerance]);

  if (dists.length < 2) {
    return (
      <div className={styles.section}>
        <SectionTitle>Stochastic dominance</SectionTitle>
        <Text className={styles.muted}>
          Add at least two distributions on the Build tab to compare them. The analysis reports expected value,
          variance and standard deviation, and tests first-order (FSD) and second-order (SSD) stochastic dominance
          exactly from the discrete CDFs.
        </Text>
      </div>
    );
  }

  function exportSheet() {
    if (!analysis || !a || !b) return;
    void ops.run(
      "Writing analysis",
      () => writeAnalysisSheet(a.name, b.name, analysis.statsA, analysis.statsB, analysis.result, analysis.explanation),
      (sheet) => ({ intent: "success", text: `Wrote the comparison to the worksheet "${sheet}".` }),
    );
  }

  const statRows: [string, (s: DistributionStats) => number][] = [
    ["Expected value", (s) => s.expectedValue],
    ["Variance", (s) => s.variance],
    ["Std. deviation", (s) => s.standardDeviation],
    ["Minimum", (s) => s.min],
    ["Median", (s) => s.median],
    ["Maximum", (s) => s.max],
  ];

  return (
    <div>
      <div className={styles.section}>
        <SectionTitle>Compare two distributions</SectionTitle>
        <div className={styles.row}>
          <Field label="Distribution A" className={styles.grow}>
            <Select value={idA} onChange={(_, d) => setIdA(d.value)}>
              {dists.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.name}
                </option>
              ))}
            </Select>
          </Field>
          <Field label="Distribution B" className={styles.grow}>
            <Select value={idB} onChange={(_, d) => setIdB(d.value)}>
              {dists.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.name}
                </option>
              ))}
            </Select>
          </Field>
        </div>
        <Field label="Outcomes are">
          <RadioGroup
            value={preference}
            layout="horizontal"
            onChange={(_, d) =>
              ops.updateSettings({ analysisSettings: { ...state.analysisSettings, preference: d.value as Preference } })
            }
          >
            <Radio value="higher" label="Higher is better (profit)" />
            <Radio value="lower" label="Lower is better (cost)" />
          </RadioGroup>
        </Field>
        {a && b && a.id === b.id ? <Text className={styles.muted}>Choose two different distributions.</Text> : null}
      </div>

      {analysis && a && b ? (
        <>
          <div className={styles.section}>
            <SectionTitle>Summary statistics</SectionTitle>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th />
                  <th>{a.name}</th>
                  <th>{b.name}</th>
                </tr>
              </thead>
              <tbody>
                {statRows.map(([label, get]) => (
                  <tr key={label}>
                    <td>{label}</td>
                    <td>{fmtNum(get(analysis.statsA))}</td>
                    <td>{fmtNum(get(analysis.statsB))}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Text className={styles.muted}>Variance and SD are for the discrete distribution: Σ p·(x − EV)².</Text>
          </div>

          <div className={styles.section}>
            <div className={styles.card}>
              <div className={styles.row} style={{ justifyContent: "space-between", alignItems: "center" }}>
                <Text weight="semibold">First-order (FSD)</Text>
                {verdictBadge(analysis.result.fsd, a.name, b.name)}
              </div>
              <Text>{analysis.explanation.fsdHeadline}</Text>
              <Text className={styles.muted}>{analysis.explanation.fsdDetail}</Text>
            </div>
            <div className={styles.card}>
              <div className={styles.row} style={{ justifyContent: "space-between", alignItems: "center" }}>
                <Text weight="semibold">Second-order (SSD)</Text>
                {verdictBadge(analysis.result.ssd, a.name, b.name)}
              </div>
              <Text>{analysis.explanation.ssdHeadline}</Text>
              <Text className={styles.muted}>{analysis.explanation.ssdDetail}</Text>
            </div>
            {analysis.explanation.notes.map((n, i) => (
              <Text key={i} className={styles.muted}>
                {n}
              </Text>
            ))}
          </div>

          <div className={styles.section}>
            <div className={styles.row}>
              <Button size="small" appearance="subtle" onClick={() => setShowTable((s) => !s)}>
                {showTable ? "Hide" : "Show"} CDF comparison ({analysis.result.rows.length} points)
              </Button>
              <Button size="small" icon={<DocumentTableRegular />} onClick={exportSheet} disabled={ops.busy}>
                Write to new sheet
              </Button>
            </div>
            {showTable ? (
              <div className={styles.scroll}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>x</th>
                      <th>F A</th>
                      <th>F B</th>
                      <th title={preference === "higher" ? "∫ F(t) dt from −∞ to x" : "∫ (1 − F(t)) dt from x to ∞"}>
                        ∫ A
                      </th>
                      <th>∫ B</th>
                    </tr>
                  </thead>
                  <tbody>
                    {analysis.result.rows.map((r) => (
                      <tr key={r.x}>
                        <td>{fmt(r.x)}</td>
                        <td>{fmtPct(r.FA)}</td>
                        <td>{fmtPct(r.FB)}</td>
                        <td>{fmtNum(r.IA, 4)}</td>
                        <td>{fmtNum(r.IB, 4)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <Text className={styles.muted}>
                  ∫ = {preference === "higher" ? "∫₋∞ˣ F(t) dt" : "∫ₓ^∞ (1 − F(t)) dt"}, the quantity compared for SSD.
                </Text>
              </div>
            ) : null}
          </div>
        </>
      ) : null}
    </div>
  );
}

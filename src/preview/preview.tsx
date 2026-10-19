/**
 * Browser preview harness (development only; not part of the add-in build).
 *
 * Runs the real task pane against the in-memory Excel fake used by the
 * tests, next to a live view of the fake workbook, so the UI can be tried in
 * any browser without Excel: `npm run preview`, then open http://localhost:8123.
 */
import { Button, FluentProvider, Input, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import { createRoot } from "react-dom/client";
import { cellRef } from "../core/a1";
import { App } from "../taskpane/App";
import { FakeSheet, installFakeExcel } from "../../tests/fakes/fakeExcel";

const env = installFakeExcel();
const wb = env.workbook;
const data = wb.addSheet("Data");
data.set("A1:B5", [
  ["Make Component", ""],
  ["Outcome", "Probability"],
  [11, 0.4],
  [32, 0.3],
  [53, 0.3],
]);
data.set("A8:B12", [
  ["Buy Component", ""],
  ["Outcome", "Probability"],
  [15, 0.4],
  [30, 0.3],
  [45, 0.3],
]);
data.set("A15:B19", [
  ["Unsorted with a duplicate (percent text)", ""],
  ["Outcome", "Probability"],
  [40, "25%"],
  [-10, "50%"],
  [40, "25%"],
]);
wb.select(data, "A2:B5");

function SheetView({ sheet }: { sheet: FakeSheet }) {
  let maxR = 0;
  let maxC = 0;
  for (const key of sheet.cells.keys()) {
    const [r, c] = key.split(",").map(Number);
    maxR = Math.max(maxR, r);
    maxC = Math.max(maxC, c);
  }
  const sel = wb.selection?.sheet === sheet ? wb.selection.rects[0] : null;
  const inSel = (r: number, c: number) =>
    !!sel && r >= sel.row && r < sel.row + sel.rows && c >= sel.col && c < sel.col + sel.cols;
  return (
    <div>
      <table style={{ borderCollapse: "collapse", fontSize: 12, fontFamily: "monospace" }}>
        <tbody>
          {Array.from({ length: maxR + 2 }, (_, r) => (
            <tr key={r}>
              <td style={{ color: "#999", paddingRight: 6 }}>{r + 1}</td>
              {Array.from({ length: maxC + 2 }, (_, c) => {
                const cell = sheet.cells.get(`${r},${c}`);
                const v = cell?.value ?? "";
                return (
                  <td
                    key={c}
                    title={cellRef(r, c)}
                    onClick={(e) => {
                      if (e.shiftKey && sel) {
                        const r0 = Math.min(sel.row, r);
                        const c0 = Math.min(sel.col, c);
                        const r1 = Math.max(sel.row + sel.rows - 1, r);
                        const c1 = Math.max(sel.col + sel.cols - 1, c);
                        wb.select(sheet, `${cellRef(r0, c0)}:${cellRef(r1, c1)}`);
                      } else {
                        wb.select(sheet, cellRef(r, c));
                      }
                    }}
                    style={{
                      border: "1px solid #ddd",
                      padding: "1px 6px",
                      minWidth: 48,
                      cursor: "cell",
                      fontWeight: cell?.bold ? 700 : 400,
                      background: inSel(r, c) ? "#cfe3f7" : undefined,
                      textAlign: typeof v === "number" ? "right" : "left",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {typeof v === "number" ? (cell?.format === "0%" ? `${Math.round(v * 100)}%` : String(v)) : String(v)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      {sheet.chartsList.map((chart) => (
        <div key={chart._name} style={{ marginTop: 12, fontSize: 12, fontFamily: "monospace" }}>
          <b>Chart “{chart._name}”</b> ({chart.seriesList[0]?.chartType}) title=“{chart.title.text}” Y=[
          {String(chart.axes.valueAxis.minimum)}, {String(chart.axes.valueAxis.maximum)}] fmt=
          {chart.axes.valueAxis.numberFormat} X.min={String(chart.axes.categoryAxis.minimum)} legend=
          {String(chart.legend.visible)}
          <ul style={{ margin: 0 }}>
            {chart.seriesList.map((s, i) => (
              <li key={i}>
                {s.name}: X={s.x ? `${cellRef(s.x.r, s.x.c)}:${cellRef(s.x.r + s.x.nr - 1, s.x.c)}` : "—"} Y=
                {s.y ? `${cellRef(s.y.r, s.y.c)}:${cellRef(s.y.r + s.y.nr - 1, s.y.c)}` : "—"} markers={s.markerStyle}
              </li>
            ))}
          </ul>
        </div>
      ))}
    </div>
  );
}

function Workbook() {
  const [, setTick] = React.useState(0);
  const [address, setAddress] = React.useState("A9:B12");
  const [sheetIdx, setSheetIdx] = React.useState(0);
  React.useEffect(() => {
    const t = window.setInterval(() => setTick((x) => x + 1), 400);
    return () => window.clearInterval(t);
  }, []);
  const sheet = wb.sheetsList[Math.min(sheetIdx, wb.sheetsList.length - 1)];
  return (
    <div style={{ padding: 12, overflow: "auto", flex: 1 }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8, flexWrap: "wrap" }}>
        {wb.sheetsList.map((s, i) => (
          <Button key={s._id} size="small" appearance={s === sheet ? "primary" : "secondary"} onClick={() => setSheetIdx(i)}>
            {s._name}
          </Button>
        ))}
        <Input size="small" value={address} onChange={(_, d) => setAddress(d.value)} style={{ width: 110 }} />
        <Button size="small" onClick={() => wb.select(sheet, address)}>
          Select range
        </Button>
        <span style={{ fontSize: 12, color: "#666" }}>
          Selection: {wb.selection ? `${wb.selection.sheet._name}!${wb.selection.rects.map((r) => cellRef(r.row, r.col) + ":" + cellRef(r.row + r.rows - 1, r.col + r.cols - 1)).join(",")}` : "—"} (click a cell, shift-click to extend)
        </span>
      </div>
      <SheetView sheet={sheet} />
    </div>
  );
}

createRoot(document.getElementById("container")!).render(
  <FluentProvider theme={webLightTheme}>
    <div style={{ display: "flex", minHeight: "100vh" }}>
      <div style={{ width: 340, flexShrink: 0, borderRight: "1px solid #ddd" }}>
        <App />
      </div>
      <Workbook />
    </div>
  </FluentProvider>,
);

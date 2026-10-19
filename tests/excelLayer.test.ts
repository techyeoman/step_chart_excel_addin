/**
 * Integration tests of the Excel layer (src/excel/*) against an in-memory
 * fake of the Excel JavaScript API that enforces the load/sync contract.
 * They cover the real-workbook workflow: read selection -> validate ->
 * write tables -> chart -> add second distribution -> regenerate -> remove.
 */
import { beforeEach, describe, expect, it } from "vitest";
import { parseDistribution } from "../src/core/parse";
import { compareDistributions, explainDominance } from "../src/core/dominance";
import { computeStats } from "../src/core/stats";
import { writeAnalysisSheet } from "../src/excel/analysisSheet";
import { AppState, defaultState, Placement } from "../src/excel/model";
import { readSelectedCell, readSelection } from "../src/excel/rangeIO";
import { addOrUpdate, Draft, removeDistribution, renderAll, updateEntry } from "../src/excel/service";
import { loadState } from "../src/excel/storage";
import { FakeEnv, FakeSheet, installFakeExcel } from "./fakes/fakeExcel";

let env: FakeEnv;
let data: FakeSheet;

function seed() {
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
}

async function draftFor(address: string, placement?: Placement, overrides: Partial<Draft> = {}): Promise<Draft> {
  env.workbook.select(data, address);
  const snapshot = await readSelection();
  const parse = parseDistribution(snapshot.values, {
    origin: { rowIndex: snapshot.rowIndex, columnIndex: snapshot.columnIndex },
  });
  const name = snapshot.labelAbove ?? "Distribution";
  return {
    snapshot,
    parse,
    name,
    outputName: `CDF_${name.replace(/\s+/g, "_")}`,
    headerMode: "auto",
    probabilityFormat: "auto",
    includeHeaders: true,
    placement: placement ?? { kind: "active", sheetId: data._id },
    ...overrides,
  };
}

const chartOf = (sheet: FakeSheet) => sheet.chartsList[0];

beforeEach(() => {
  env = installFakeExcel();
  data = env.workbook.addSheet("Data");
  seed();
});

describe("reading the selection", () => {
  it("reads a two-column range with headers and the label above it", async () => {
    env.workbook.select(data, "A2:B5");
    const snap = await readSelection();
    expect(snap.displayAddress).toBe("Data!A2:B5");
    expect(snap.labelAbove).toBe("Make Component");
    expect(snap.values[0]).toEqual(["Outcome", "Probability"]);
    expect(snap.expanded).toBe(false);
  });

  it("expands a single cell to its region and drops a title row", async () => {
    env.workbook.select(data, "A3");
    const snap = await readSelection();
    expect(snap.expanded).toBe(true);
    expect(snap.source.address).toBe("A2:B5");
    expect(snap.labelAbove).toBe("Make Component");
  });

  it("trims whole-column selections to the used range", async () => {
    env.workbook.select(data, "A1:B1048576");
    const snap = await readSelection();
    expect(snap.source.address).toBe("A1:B12");
  });

  it("rejects multi-area selections", async () => {
    env.workbook.select(data, "A2:B5,A9:B12");
    await expect(readSelection()).rejects.toThrow(/multiple areas/);
  });

  it("reads the selected output cell", async () => {
    env.workbook.select(data, "K20");
    expect(await readSelectedCell()).toMatchObject({ row: 19, col: 10, displayAddress: "Data!K20" });
  });
});

describe("generating tables and the chart", () => {
  it("writes the Make Component table and creates a formatted XY scatter chart", async () => {
    const { state, notes } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    expect(notes).toEqual([]);
    const entry = state.distributions[0];
    expect(entry.output).toMatchObject({ sheetName: "Data", row: 0, col: 3, dataRow: 2, dataCount: 7, tableName: "CDF_Make_Component" });
    expect(data.readBlock("D1:E9")).toEqual([
      ["Make Component", ""],
      ["X (Outcome)", "Y (Cumulative Probability)"],
      [0, 0],
      [11, 0],
      [11, 0.4],
      [32, 0.4],
      [32, 0.7],
      [53, 0.7],
      [53, 1],
    ]);
    expect(env.workbook.tablesList.map((t) => t._name)).toEqual(["CDF_Make_Component"]);

    const chart = chartOf(data);
    expect(chart._name).toBe("ProbabilityStepChart");
    expect(chart.seriesList).toHaveLength(1);
    const s = chart.seriesList[0];
    expect(s.name).toBe("Make Component");
    expect([s.x!.r, s.x!.c, s.x!.nr, s.y!.c]).toEqual([2, 3, 7, 4]);
    expect(s.chartType).toBe("XYScatterLinesNoMarkers");
    expect(s.markerStyle).toBe("None");
    expect(s.smooth).toBe(false);
    const y = chart.axes.valueAxis;
    expect([y.minimum, y.maximum, y.numberFormat]).toEqual([0, 1, "0%"]);
    expect(y.title).toEqual({ text: "Cumulative Probability", visible: true });
    expect(chart.axes.categoryAxis.title.text).toBe("Outcome");
    expect(chart.axes.categoryAxis.minimum).toBe(0);
    expect(y.majorGridlines.visible).toBe(true);
    expect(chart.legend.visible).toBe(false);
    expect(chart.title).toEqual({ text: "Cumulative Probability Distribution", visible: true });
    expect(state.chart).toEqual({ sheetId: data._id, name: "ProbabilityStepChart" });
  });

  it("adds a second distribution to the same chart with a shared X range", async () => {
    let { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    ({ state } = await addOrUpdate(state, await draftFor("A9:B12")));
    expect(state.distributions).toHaveLength(2);
    const buy = state.distributions[1];
    expect(buy.output).toMatchObject({ col: 6, dataCount: 8, tableName: "CDF_Buy_Component" });
    // Extended to the chart maximum (53) at probability 1.
    expect(data.readBlock("G10:H10")).toEqual([[53, 1]]);
    const chart = chartOf(data);
    expect(data.chartsList).toHaveLength(1);
    expect(chart.seriesList.map((s) => s.name)).toEqual(["Make Component", "Buy Component"]);
    expect(chart.legend.visible).toBe(true);
    expect(chart.axes.categoryAxis.maximum).toBe(53);
    // The chart was moved off the new table.
    expect(chart["_left"]).toBeGreaterThanOrEqual(8 * 64);
  });

  it("updates in place when the same range is generated again", async () => {
    let { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    ({ state } = await addOrUpdate(state, await draftFor("A2:B5", undefined, { name: "Make (v2)" })));
    expect(state.distributions).toHaveLength(1);
    expect(data.read(0, 3)).toBe("Make (v2)");
    expect(env.workbook.tablesList).toHaveLength(1);
    expect(chartOf(data).seriesList.map((s) => s.name)).toEqual(["Make (v2)"]);
  });

  it("regenerates from edited source data", async () => {
    let { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    data.set("B3:B4", [[0.5], [0.2]]);
    ({ state } = await renderAll(state, { reread: true }));
    expect(state.distributions[0].outcomes.map((o) => o.p)).toEqual([0.5, 0.2, 0.3]);
    expect(data.readBlock("D5:E7")).toEqual([
      [11, 0.5],
      [32, 0.5],
      [32, 0.7],
    ]);
  });

  it("refuses to regenerate from invalid source data and names the cell", async () => {
    const { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    data.set("B3", [[0.9]]);
    await expect(renderAll(state, { reread: true })).rejects.toThrow(/Make Component.*sum to 1\.5/);
    expect(data.read(4, 4)).toBe(0.4); // table untouched
  });

  it("never overwrites existing data at a chosen output cell", async () => {
    const draft = await draftFor("A2:B5", { kind: "cell", sheetId: data._id, row: 6, col: 0 });
    await expect(addOrUpdate(defaultState(), draft)).rejects.toThrow(/A8.*not empty/);
    expect(env.workbook.tablesList).toHaveLength(0);
    expect(data.chartsList).toHaveLength(0);
  });

  it("writes to a new worksheet and places later distributions beside earlier ones", async () => {
    let { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5", { kind: "newSheet", sheetName: "CDF Output" }));
    ({ state } = await addOrUpdate(state, await draftFor("A9:B12", { kind: "newSheet", sheetName: "CDF Output" })));
    const out = env.workbook.sheetByName("CDF Output")!;
    expect(state.distributions.map((d) => [d.output!.sheetName, d.output!.col])).toEqual([
      ["CDF Output", 0],
      ["CDF Output", 3],
    ]);
    expect(out.read(0, 0)).toBe("Make Component");
    expect(chartOf(out).seriesList).toHaveLength(2);
  });

  it("writes plain data with a defined name when headers are off", async () => {
    const { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5", undefined, { includeHeaders: false }));
    expect(state.distributions[0].output).toMatchObject({ definedName: "CDF_Make_Component", tableName: null, dataRow: 0, rows: 7 });
    expect(data.readBlock("D1:E1")).toEqual([[0, 0]]);
    expect(env.workbook.namesMap.has("cdf_make_component")).toBe(true);
  });

  it("makes output table names unique", async () => {
    let { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5", undefined, { outputName: "Results" }));
    ({ state } = await addOrUpdate(state, await draftFor("A9:B12", undefined, { outputName: "Results" })));
    expect(state.distributions.map((d) => d.output!.tableName)).toEqual(["Results", "Results_2"]);
  });

  it("falls back to percentage-formatted cells when chart axis formatting (ExcelApi 1.8) is unavailable", async () => {
    env = installFakeExcel({ apiVersion: "1.7" });
    data = env.workbook.addSheet("Data");
    seed();
    const { notes } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    expect(notes.join(" ")).toMatch(/ExcelApi 1\.8/);
    expect(data.cells.get("2,4")?.format).toBe("0%");
    expect(chartOf(data).axes.valueAxis.numberFormat).toBe("General");
  });

  it("validates the shared X range before writing", async () => {
    const base: AppState = { ...defaultState(), rangeSettings: { startMode: "custom", customStart: "20", endMode: "none", customEnd: "" } };
    await expect(addOrUpdate(base, await draftFor("A2:B5"))).rejects.toThrow(/starting X value \(20\)/);
    expect(env.workbook.tablesList).toHaveLength(0);
  });
});

describe("managing distributions", () => {
  async function both() {
    let { state } = await addOrUpdate(defaultState(), await draftFor("A2:B5"));
    ({ state } = await addOrUpdate(state, await draftFor("A9:B12")));
    return state;
  }

  it("renames a distribution (series and label)", async () => {
    const state = await both();
    const { state: next } = await updateEntry(state, state.distributions[1].id, { name: "Outsource" });
    expect(next.distributions[1].name).toBe("Outsource");
    expect(chartOf(data).seriesList[1].name).toBe("Outsource");
    expect(data.read(0, 6)).toBe("Outsource");
  });

  it("hides a distribution from the chart without deleting its table", async () => {
    const state = await both();
    await updateEntry(state, state.distributions[0].id, { includeInChart: false });
    expect(chartOf(data).seriesList.map((s) => s.name)).toEqual(["Buy Component"]);
    expect(env.workbook.tablesList).toHaveLength(2);
  });

  it("removes distributions, their tables and finally the chart", async () => {
    let state = await both();
    ({ state } = await removeDistribution(state, state.distributions[1].id));
    expect(env.workbook.tablesList.map((t) => t._name)).toEqual(["CDF_Make_Component"]);
    expect(data.readBlock("G1:H10").flat().every((v) => v === "")).toBe(true);
    expect(chartOf(data).seriesList).toHaveLength(1);
    // Source data untouched.
    expect(data.readBlock("A9:B12")[1]).toEqual([15, 0.4]);
    ({ state } = await removeDistribution(state, state.distributions[0].id));
    expect(state.chart).toBeNull();
    expect(data.chartsList).toHaveLength(0);
    expect(env.workbook.tablesList).toHaveLength(0);
  });

  it("persists state in the workbook settings", async () => {
    await both();
    const loaded = await loadState();
    expect(loaded.distributions.map((d) => d.name)).toEqual(["Make Component", "Buy Component"]);
    expect(loaded.chart?.name).toBe("ProbabilityStepChart");
  });

  it("recreates the chart if the user deleted it", async () => {
    const state = await both();
    chartOf(data).delete();
    await renderAll(state, { reread: false });
    expect(chartOf(data).seriesList).toHaveLength(2);
  });
});

describe("analysis export", () => {
  it("writes statistics, verdicts and the comparison table to a new sheet", async () => {
    const a = { outcomes: [{ x: 11, p: 0.4 }, { x: 32, p: 0.3 }, { x: 53, p: 0.3 }] };
    const b = { outcomes: [{ x: 15, p: 0.4 }, { x: 30, p: 0.3 }, { x: 45, p: 0.3 }] };
    const result = compareDistributions(a, b);
    const name = await writeAnalysisSheet("Make", "Buy", computeStats(a), computeStats(b), result, explainDominance(result, "Make", "Buy"));
    expect(name).toBe("Dominance Analysis");
    const sheet = env.workbook.sheetByName(name)!;
    expect(sheet.readBlock("A5:C5")).toEqual([["Expected value", 29.9, 28.5]]);
    expect(String(sheet.read(11, 1))).toMatch(/Neither/);
    expect(env.workbook.tablesList.some((t) => t._name === "Dominance_Comparison")).toBe(true);
    expect(await writeAnalysisSheet("Make", "Buy", computeStats(a), computeStats(b), result, explainDominance(result, "Make", "Buy"))).toBe(
      "Dominance Analysis 2",
    );
  });
});

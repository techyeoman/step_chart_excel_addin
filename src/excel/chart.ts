/**
 * XY scatter (straight lines) step chart.
 *
 * API requirements (checked against @types/office-js):
 *   charts.add, title, legend, axis min/max/majorUnit/title/gridlines   ExcelApi 1.1
 *   series.add/delete, setXAxisValues, setValues, chartType,
 *   markerStyle, smooth, format.line.weight                              ExcelApi 1.7 (manifest minimum)
 *   axis.numberFormat, axis.position                                     ExcelApi 1.8 (feature-detected)
 *   Range.left/top/width/height (overlap check)                          ExcelApi 1.10 (feature-detected)
 *
 * Every formatting group is synced separately inside try/catch, so a
 * platform that rejects one property still gets a working chart; the
 * skipped feature is reported back as a note.
 */
import { cellRef, uniqueName } from "../core/a1";
import { CHART_NAME, ChartRef, ChartSettings, OutputBlock } from "./model";
import { describeError, excelApi } from "./support";

export interface SeriesSource {
  name: string;
  block: OutputBlock;
}

export interface ChartResult {
  chart: ChartRef | null;
  notes: string[];
}

async function tryStep(context: Excel.RequestContext, label: string, notes: string[], apply: () => void): Promise<void> {
  try {
    apply();
    await context.sync();
  } catch (error) {
    notes.push(`${label} could not be applied on this version of Excel (${describeError(error)}).`);
  }
}

async function findChart(context: Excel.RequestContext, ref: ChartRef | null): Promise<Excel.Chart | null> {
  if (!ref) return null;
  const sheet = context.workbook.worksheets.getItemOrNullObject(ref.sheetId);
  await context.sync();
  if (sheet.isNullObject) return null;
  const chart = sheet.charts.getItemOrNullObject(ref.name);
  await context.sync();
  return chart.isNullObject ? null : chart;
}

export async function deleteChart(context: Excel.RequestContext, ref: ChartRef | null): Promise<void> {
  const chart = await findChart(context, ref);
  if (chart) {
    chart.delete();
    await context.sync();
  }
}

function dataRanges(context: Excel.RequestContext, block: OutputBlock): { x: Excel.Range; y: Excel.Range } {
  const sheet = context.workbook.worksheets.getItem(block.sheetId);
  return {
    x: sheet.getRangeByIndexes(block.dataRow, block.col, block.dataCount, 1),
    y: sheet.getRangeByIndexes(block.dataRow, block.col + 1, block.dataCount, 1),
  };
}

export async function renderChart(
  context: Excel.RequestContext,
  ref: ChartRef | null,
  series: SeriesSource[],
  xRange: { xStart: number; xEnd: number | null },
  settings: ChartSettings,
): Promise<ChartResult> {
  const notes: string[] = [];
  if (series.length === 0) {
    await deleteChart(context, ref);
    return { chart: null, notes };
  }

  const chartType = settings.markers ? Excel.ChartType.xyscatterLines : Excel.ChartType.xyscatterLinesNoMarkers;
  let chart = await findChart(context, ref);
  let chartSheetId: string;
  const created = !chart;

  if (chart) {
    chart.worksheet.load("id");
    chart.load("name");
    await context.sync();
    chartSheetId = chart.worksheet.id;
  } else {
    chartSheetId = series[0].block.sheetId;
    const sheet = context.workbook.worksheets.getItem(chartSheetId);
    sheet.charts.load("items/name");
    await context.sync();
    const takenNames = sheet.charts.items.map((c) => c.name);
    const first = dataRanges(context, series[0].block);
    chart = sheet.charts.add(chartType, first.y, Excel.ChartSeriesBy.columns);
    chart.name = uniqueName(CHART_NAME, takenNames);
    // Place the chart to the right of the output blocks on this sheet.
    const onSheet = series.map((s) => s.block).filter((b) => b.sheetId === chartSheetId);
    const right = Math.max(...onSheet.map((b) => b.col + b.cols));
    const top = Math.min(...onSheet.map((b) => b.row));
    chart.setPosition(cellRef(top, right + 1), cellRef(top + 19, right + 9));
    chart.load("name");
    await context.sync();
  }
  const theChart = chart;

  // Series: reuse existing ones (keeps any colours the user picked), add or delete the rest.
  theChart.series.load("items/name");
  await context.sync();
  const existing = theChart.series.items.length;
  series.forEach((s, i) => {
    const target = i < existing ? theChart.series.getItemAt(i) : theChart.series.add(s.name);
    const { x, y } = dataRanges(context, s.block);
    target.name = s.name;
    target.setXAxisValues(x);
    target.setValues(y);
  });
  for (let i = existing - 1; i >= series.length; i--) theChart.series.getItemAt(i).delete();
  await context.sync();

  theChart.series.load("items");
  await context.sync();
  const allSeries = theChart.series.items;

  await tryStep(context, "Straight-line scatter style", notes, () => {
    for (const s of allSeries) {
      s.chartType = chartType;
      s.smooth = false;
    }
  });
  await tryStep(context, "Marker setting", notes, () => {
    for (const s of allSeries) s.markerStyle = settings.markers ? Excel.ChartMarkerStyle.automatic : Excel.ChartMarkerStyle.none;
  });
  await tryStep(context, "Line weight", notes, () => {
    for (const s of allSeries) s.format.line.weight = 2.25;
  });

  await tryStep(context, "Chart title", notes, () => {
    if (settings.title.trim()) {
      theChart.title.text = settings.title.trim();
      theChart.title.visible = true;
    } else {
      theChart.title.visible = false;
    }
  });
  await tryStep(context, "Legend", notes, () => {
    theChart.legend.visible = series.length > 1;
    theChart.legend.position = Excel.ChartLegendPosition.bottom;
  });

  const yAxis = theChart.axes.valueAxis;
  const xAxis = theChart.axes.categoryAxis; // the horizontal (X value) axis of a scatter chart
  await tryStep(context, "Y-axis bounds 0–1", notes, () => {
    yAxis.minimum = 0;
    yAxis.maximum = 1;
    yAxis.majorUnit = 0.1;
  });
  await tryStep(context, "Axis titles", notes, () => {
    setAxisTitle(yAxis, settings.yAxisTitle);
    setAxisTitle(xAxis, settings.xAxisTitle);
  });
  await tryStep(context, "Gridlines", notes, () => {
    yAxis.majorGridlines.visible = settings.gridlines;
    xAxis.majorGridlines.visible = settings.gridlines;
  });
  await tryStep(context, "X-axis bounds", notes, () => {
    xAxis.minimum = xRange.xStart;
    if (xRange.xEnd !== null && xRange.xEnd > xRange.xStart) xAxis.maximum = xRange.xEnd;
  });

  if (excelApi("1.8")) {
    await tryStep(context, "Y-axis number format", notes, () => {
      yAxis.numberFormat = settings.yAsPercent ? "0%" : "0.0";
    });
    // Keep the Y axis at the left edge even when X values are negative.
    await tryStep(context, "Y-axis position", notes, () => {
      xAxis.position = Excel.ChartAxisPosition.minimum;
    });
  } else if (settings.yAsPercent) {
    // Fallback: the axis number format is linked to the source cells by default,
    // so formatting the Y data cells as percentages formats the axis too.
    await tryStep(context, "Percentage format (via source cells)", notes, () => {
      for (const s of series) dataRanges(context, s.block).y.numberFormat = Array.from({ length: s.block.dataCount }, () => ["0%"]);
    });
    notes.push("This Excel version cannot format chart axes directly (ExcelApi 1.8); Y values were formatted as percentages in the table instead.");
  }

  if (!created) await moveOffData(context, theChart, chartSheetId, series, notes);

  return { chart: { sheetId: chartSheetId, name: theChart.name }, notes };
}

function setAxisTitle(axis: Excel.ChartAxis, text: string): void {
  if (text.trim()) {
    axis.title.text = text.trim();
    axis.title.visible = true;
  } else {
    axis.title.visible = false;
  }
}

/** If an existing chart now covers output data on its sheet, move it to the right of that data. */
async function moveOffData(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  sheetId: string,
  series: SeriesSource[],
  notes: string[],
): Promise<void> {
  if (!excelApi("1.10")) return;
  const blocks = series.map((s) => s.block).filter((b) => b.sheetId === sheetId);
  if (blocks.length === 0) return;
  try {
    const sheet = context.workbook.worksheets.getItem(sheetId);
    const ranges = blocks.map((b) => {
      const r = sheet.getRangeByIndexes(b.row, b.col, b.rows, b.cols);
      r.load("left,top,width,height");
      return r;
    });
    chart.load("left,top,width,height");
    await context.sync();
    const overlaps = ranges.some(
      (r) =>
        chart.left < r.left + r.width &&
        chart.left + chart.width > r.left &&
        chart.top < r.top + r.height &&
        chart.top + chart.height > r.top,
    );
    if (overlaps) {
      chart.left = Math.max(...ranges.map((r) => r.left + r.width)) + 24;
      await context.sync();
    }
  } catch (error) {
    notes.push(`Could not check the chart position (${describeError(error)}).`);
  }
}

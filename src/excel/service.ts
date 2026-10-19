/**
 * High-level operations used by the task pane. Each operation validates
 * everything it can before touching the workbook, writes all output tables,
 * updates the chart, and persists the resulting state.
 */
import { buildStepCoordinates, resolveXRange, XRangeOptions } from "../core/cdf";
import { parseDistribution, ParseResult } from "../core/parse";
import { Point } from "../core/types";
import { renderChart, deleteChart, SeriesSource } from "./chart";
import { AppState, DistributionEntry, newId, Placement, RangeSettings } from "./model";
import { clearBlock, writeBlock } from "./output";
import { readSource, RangeSnapshot } from "./rangeIO";
import { saveState } from "./storage";
import { UserError } from "./support";

export interface OperationResult {
  state: AppState;
  notes: string[];
}

export interface Draft {
  snapshot: RangeSnapshot;
  parse: ParseResult;
  name: string;
  outputName: string;
  headerMode: DistributionEntry["headerMode"];
  probabilityFormat: DistributionEntry["probabilityFormat"];
  includeHeaders: boolean;
  placement: Placement;
}

export function toXRangeOptions(settings: RangeSettings): XRangeOptions {
  const num = (s: string) => (s.trim() === "" ? undefined : Number(s));
  return {
    startMode: settings.startMode,
    customStart: num(settings.customStart),
    endMode: settings.endMode,
    customEnd: num(settings.customEnd),
  };
}

function firstErrors(parse: ParseResult, max = 3): string {
  const errors = parse.issues.filter((i) => i.severity === "error");
  const text = errors
    .slice(0, max)
    .map((e) => (e.cell ? `${e.cell}: ${e.message}` : e.message))
    .join(" ");
  return errors.length > max ? `${text} (and ${errors.length - max} more)` : text;
}

/** Compute all step coordinates up front so invalid settings fail before any write. */
function computePoints(state: AppState): { points: Map<string, Point[]>; xRange: { xStart: number; xEnd: number | null } } {
  const dists = state.distributions.map((e) => ({ outcomes: e.outcomes }));
  let xRange: { xStart: number; xEnd: number | null };
  try {
    xRange = resolveXRange(dists, toXRangeOptions(state.rangeSettings));
  } catch (error) {
    throw new UserError((error as Error).message);
  }
  const points = new Map<string, Point[]>();
  for (const e of state.distributions) {
    try {
      points.set(e.id, buildStepCoordinates({ outcomes: e.outcomes }, { ...xRange, tolerance: state.tolerance }));
    } catch (error) {
      throw new UserError(`"${e.name}": ${(error as Error).message}`);
    }
  }
  return { points, xRange };
}

/**
 * Rewrite every distribution's table and refresh the chart. With `reread`,
 * each source range is read and validated again first (for "Regenerate").
 */
export async function renderAll(input: AppState, options: { reread: boolean }): Promise<OperationResult> {
  const state: AppState = { ...input, distributions: input.distributions.map((e) => ({ ...e })) };
  const notes: string[] = [];

  if (options.reread) {
    for (const entry of state.distributions) {
      const snap = await readSource(entry.source);
      const parse = parseDistribution(snap.values, {
        headerMode: entry.headerMode,
        probabilityFormat: entry.probabilityFormat,
        tolerance: state.tolerance,
        origin: { rowIndex: snap.rowIndex, columnIndex: snap.columnIndex },
      });
      if (!parse.ok || !parse.distribution) {
        throw new UserError(`"${entry.name}" (${snap.displayAddress}) is no longer valid: ${firstErrors(parse)}`);
      }
      entry.outcomes = parse.distribution.outcomes;
      entry.source = snap.source;
      entry.xNumberFormat = snap.xNumberFormat;
    }
  }

  if (state.distributions.length === 0) {
    await Excel.run(async (context) => deleteChart(context, state.chart));
    state.chart = null;
    await saveState(state);
    return { state, notes };
  }

  const { points, xRange } = computePoints(state);

  try {
    await Excel.run(async (context) => {
      for (const entry of state.distributions) {
        entry.output = await writeBlock(context, entry, points.get(entry.id)!);
      }
      if (state.chartSettings.createChart) {
        const series: SeriesSource[] = state.distributions
          .filter((e) => e.includeInChart && e.output)
          .map((e) => ({ name: e.name, block: e.output! }));
        const result = await renderChart(context, state.chart, series, xRange, state.chartSettings);
        state.chart = result.chart;
        notes.push(...result.notes);
      }
    });
  } finally {
    // Persist whatever was written, even if a later step failed.
    await saveState(state);
  }
  return { state, notes };
}

/** Add the draft as a new distribution, or update the one with the same source range. */
export async function addOrUpdate(state: AppState, draft: Draft): Promise<OperationResult> {
  if (!draft.parse.ok || !draft.parse.distribution) throw new UserError(firstErrors(draft.parse));
  const name = draft.name.trim() || "Distribution";
  const existing = state.distributions.find(
    (e) => e.source.sheetId === draft.snapshot.source.sheetId && e.source.address === draft.snapshot.source.address,
  );

  let entry: DistributionEntry;
  if (existing) {
    const moved = JSON.stringify(existing.placement) !== JSON.stringify(draft.placement);
    if (moved && existing.output) {
      await Excel.run((context) => clearBlock(context, existing.output!));
    }
    entry = {
      ...existing,
      name,
      outputName: draft.outputName,
      headerMode: draft.headerMode,
      probabilityFormat: draft.probabilityFormat,
      outcomes: draft.parse.distribution.outcomes,
      includeHeaders: draft.includeHeaders,
      xNumberFormat: draft.snapshot.xNumberFormat,
      source: draft.snapshot.source,
      placement: draft.placement,
      output: moved ? null : existing.output,
    };
  } else {
    entry = {
      id: newId(),
      name,
      source: draft.snapshot.source,
      headerMode: draft.headerMode,
      probabilityFormat: draft.probabilityFormat,
      outcomes: draft.parse.distribution.outcomes,
      outputName: draft.outputName,
      includeHeaders: draft.includeHeaders,
      placement: draft.placement,
      output: null,
      includeInChart: true,
      xNumberFormat: draft.snapshot.xNumberFormat,
    };
  }
  const distributions = existing
    ? state.distributions.map((e) => (e.id === existing.id ? entry : e))
    : [...state.distributions, entry];
  const next = { ...state, distributions };
  // Validate the shared X range before writing anything.
  computePoints(next);
  return renderAll(next, { reread: false });
}

/** Remove a distribution: delete its output block and its chart series. */
export async function removeDistribution(state: AppState, id: string): Promise<OperationResult> {
  const entry = state.distributions.find((e) => e.id === id);
  if (entry?.output) await Excel.run((context) => clearBlock(context, entry.output!));
  return renderAll({ ...state, distributions: state.distributions.filter((e) => e.id !== id) }, { reread: false });
}

/** Rename / include-in-chart edits that only need the chart and labels refreshed. */
export async function updateEntry(
  state: AppState,
  id: string,
  patch: Partial<Pick<DistributionEntry, "name" | "includeInChart" | "outputName">>,
): Promise<OperationResult> {
  return renderAll(
    { ...state, distributions: state.distributions.map((e) => (e.id === id ? { ...e, ...patch } : e)) },
    { reread: false },
  );
}

/**
 * Persist the add-in state inside the workbook (Excel.SettingCollection,
 * ExcelApi 1.4) so distributions, output tables and chart options survive
 * closing the task pane or reopening the file.
 */
import { AppState, defaultState } from "./model";

const KEY = "ProbabilityStepChart.state";

export async function loadState(): Promise<AppState> {
  return Excel.run(async (context) => {
    const setting = context.workbook.settings.getItemOrNullObject(KEY);
    setting.load("value");
    await context.sync();
    if (setting.isNullObject) return defaultState();
    try {
      const raw = typeof setting.value === "string" ? JSON.parse(setting.value) : setting.value;
      return mergeState(raw);
    } catch {
      return defaultState();
    }
  });
}

export async function saveState(state: AppState): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.settings.add(KEY, JSON.stringify(state));
    await context.sync();
  });
}

/** Merge persisted data over defaults so new settings get sensible values. */
function mergeState(raw: Partial<AppState> | null | undefined): AppState {
  const base = defaultState();
  if (!raw || raw.version !== 1) return base;
  return {
    ...base,
    ...raw,
    distributions: Array.isArray(raw.distributions) ? raw.distributions : [],
    chartSettings: { ...base.chartSettings, ...raw.chartSettings },
    rangeSettings: { ...base.rangeSettings, ...raw.rangeSettings },
    outputSettings: { ...base.outputSettings, ...raw.outputSettings },
    analysisSettings: { ...base.analysisSettings, ...raw.analysisSettings },
  };
}

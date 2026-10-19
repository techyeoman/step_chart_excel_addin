/** Workbook-level state shared by the Excel layer and the task pane. */
import { EndMode, StartMode } from "../core/cdf";
import { Preference } from "../core/dominance";
import { HeaderMode, ProbabilityFormat } from "../core/parse";
import { DEFAULT_TOLERANCE, Outcome } from "../core/types";

export type OutputLocationKind = "active" | "newSheet" | "cell";

/** Where a distribution's output block is placed the first time it is written. */
export type Placement =
  | { kind: "active"; sheetId: string }
  | { kind: "newSheet"; sheetName: string }
  | { kind: "cell"; sheetId: string; row: number; col: number };

/** The rectangle written for one distribution (label row + table, or plain data). */
export interface OutputBlock {
  sheetId: string;
  sheetName: string;
  row: number;
  col: number;
  rows: number;
  cols: number;
  /** First row / count of the numeric X,Y data inside the block. */
  dataRow: number;
  dataCount: number;
  tableName: string | null;
  definedName: string | null;
}

export interface SourceRef {
  sheetId: string;
  sheetName: string;
  /** Address without the sheet, e.g. "A2:B5". */
  address: string;
}

export interface DistributionEntry {
  id: string;
  name: string;
  source: SourceRef;
  headerMode: HeaderMode;
  probabilityFormat: ProbabilityFormat;
  outcomes: Outcome[];
  /** Requested table (or defined) name for the generated coordinates. */
  outputName: string;
  includeHeaders: boolean;
  placement: Placement;
  output: OutputBlock | null;
  includeInChart: boolean;
  /** Number format of the source outcome cells, reused for the X column. */
  xNumberFormat: string;
}

export interface ChartSettings {
  createChart: boolean;
  title: string;
  xAxisTitle: string;
  yAxisTitle: string;
  yAsPercent: boolean;
  gridlines: boolean;
  markers: boolean;
}

export interface RangeSettings {
  startMode: StartMode;
  customStart: string;
  endMode: EndMode;
  customEnd: string;
}

export interface OutputSettings {
  location: OutputLocationKind;
  newSheetName: string;
  includeHeaders: boolean;
}

export interface AnalysisSettings {
  preference: Preference;
}

export interface ChartRef {
  sheetId: string;
  name: string;
}

export interface AppState {
  version: 1;
  distributions: DistributionEntry[];
  chart: ChartRef | null;
  chartSettings: ChartSettings;
  rangeSettings: RangeSettings;
  outputSettings: OutputSettings;
  analysisSettings: AnalysisSettings;
  tolerance: number;
}

export const CHART_NAME = "ProbabilityStepChart";

export function defaultState(): AppState {
  return {
    version: 1,
    distributions: [],
    chart: null,
    chartSettings: {
      createChart: true,
      title: "Cumulative Probability Distribution",
      xAxisTitle: "Outcome",
      yAxisTitle: "Cumulative Probability",
      yAsPercent: true,
      gridlines: true,
      markers: false,
    },
    rangeSettings: { startMode: "auto", customStart: "", endMode: "max", customEnd: "" },
    outputSettings: { location: "active", newSheetName: "CDF Output", includeHeaders: true },
    analysisSettings: { preference: "higher" },
    tolerance: DEFAULT_TOLERANCE,
  };
}

export function newId(): string {
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}

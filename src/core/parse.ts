/**
 * Parse and validate a two-column (outcome, probability) block of cell values
 * into a Distribution. Pure: works on the plain `values` matrix returned by
 * Excel, so it can be unit tested without Office.js.
 */
import { cellRef } from "./a1";
import { cleanFloat, fmt, stableSum } from "./numeric";
import { DEFAULT_TOLERANCE, Distribution, Outcome, ValidationIssue } from "./types";

export type CellValue = string | number | boolean | null | undefined;
export type HeaderMode = "auto" | "yes" | "no";
export type ProbabilityFormat = "auto" | "decimal" | "percent";

export interface ParseOptions {
  headerMode?: HeaderMode;
  /**
   * How to read probabilities. Cells formatted as Excel percentages already
   * arrive as decimals (40% -> 0.4), and text such as "40%" is always read as
   * a percentage. "percent" means plain numbers are percentage points (40 ->
   * 0.4); "auto" switches to that only when the numbers clearly sum to 100.
   */
  probabilityFormat?: ProbabilityFormat;
  /** Allowed |sum(p) - 1|. */
  tolerance?: number;
  /** Sheet position of values[0][0], used to label cells in messages. */
  origin?: { rowIndex: number; columnIndex: number };
}

export interface ParsedRow {
  x: number;
  p: number;
  /** A1 reference of the outcome cell. */
  cell: string;
}

export interface ParseResult {
  ok: boolean;
  /** Whether the first row was treated as a header row. */
  hasHeader: boolean;
  /** What automatic detection concluded (independent of the override). */
  headerDetected: boolean;
  headers: [string, string] | null;
  /** Data rows in sheet order, probabilities already scaled to [0, 1]. */
  rows: ParsedRow[];
  /** Sorted, duplicate-merged distribution; null when there are errors. */
  distribution: Distribution | null;
  probabilitySum: number;
  probabilityFormat: "decimal" | "percent";
  issues: ValidationIssue[];
}

type ParsedCell =
  | { kind: "blank" }
  | { kind: "number"; value: number; percentText: boolean }
  | { kind: "invalid"; text: string; excelError: boolean };

const NUMBER_RE = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;
const PERCENT_RE = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)\s*%$/i;
const EXCEL_ERROR_RE = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|PYTHON!)$/i;

export function parseCell(value: CellValue): ParsedCell {
  if (value === null || value === undefined) return { kind: "blank" };
  if (typeof value === "number") {
    return Number.isFinite(value)
      ? { kind: "number", value: Object.is(value, -0) ? 0 : value, percentText: false }
      : { kind: "invalid", text: String(value), excelError: false };
  }
  if (typeof value === "boolean") return { kind: "invalid", text: String(value).toUpperCase(), excelError: false };
  const s = value.trim();
  if (s === "") return { kind: "blank" };
  if (EXCEL_ERROR_RE.test(s)) return { kind: "invalid", text: s, excelError: true };
  const pct = PERCENT_RE.exec(s);
  if (pct) return { kind: "number", value: Number(pct[1]) / 100, percentText: true };
  if (NUMBER_RE.test(s)) return { kind: "number", value: Number(s), percentText: false };
  return { kind: "invalid", text: s, excelError: false };
}

/** Auto-detect a header row: the first row holds non-numeric text. */
export function detectHeader(values: CellValue[][]): boolean {
  if (values.length === 0) return false;
  const first = values[0].slice(0, 2).map(parseCell);
  return first.some((c) => c.kind === "invalid" && !c.excelError);
}

export function parseDistribution(values: CellValue[][], options: ParseOptions = {}): ParseResult {
  const tolerance = options.tolerance ?? DEFAULT_TOLERANCE;
  const origin = options.origin ?? { rowIndex: 0, columnIndex: 0 };
  const issues: ValidationIssue[] = [];
  const ref = (r: number, c: number) => cellRef(origin.rowIndex + r, origin.columnIndex + c);
  const fail = (message: string): ParseResult => ({
    ok: false,
    hasHeader: false,
    headerDetected: false,
    headers: null,
    rows: [],
    distribution: null,
    probabilitySum: 0,
    probabilityFormat: "decimal",
    issues: [...issues, { severity: "error", message }],
  });

  if (values.length === 0 || values[0].length === 0) return fail("The selection is empty.");
  const columnCount = values[0].length;
  if (columnCount !== 2) {
    return fail(
      `The selection has ${columnCount} column${columnCount === 1 ? "" : "s"}. Select exactly two columns: outcome (X) and probability.`,
    );
  }

  const headerDetected = detectHeader(values);
  const mode = options.headerMode ?? "auto";
  const hasHeader = mode === "auto" ? headerDetected : mode === "yes";
  const headers: [string, string] | null = hasHeader
    ? [String(values[0][0] ?? "").trim(), String(values[0][1] ?? "").trim()]
    : null;

  // Pass 1: read cells.
  interface RawRow {
    x: number;
    p: number;
    percentText: boolean;
    r: number;
  }
  const raw: RawRow[] = [];
  let blankRows = 0;
  for (let r = hasHeader ? 1 : 0; r < values.length; r++) {
    const cx = parseCell(values[r][0]);
    const cp = parseCell(values[r][1]);
    if (cx.kind === "blank" && cp.kind === "blank") {
      blankRows++;
      continue;
    }
    let rowOk = true;
    if (cx.kind === "blank") {
      issues.push({ severity: "error", cell: ref(r, 0), message: `Outcome is blank but its probability is filled in.` });
      rowOk = false;
    } else if (cx.kind === "invalid") {
      issues.push({
        severity: "error",
        cell: ref(r, 0),
        message: cx.excelError
          ? `Outcome cell contains the Excel error ${cx.text}.`
          : `Outcome "${cx.text}" is not a number.`,
      });
      rowOk = false;
    } else if (cx.percentText) {
      issues.push({ severity: "error", cell: ref(r, 0), message: `Outcome is entered as a percentage text; enter a number.` });
      rowOk = false;
    }
    if (cp.kind === "blank") {
      issues.push({ severity: "error", cell: ref(r, 1), message: `Probability is blank.` });
      rowOk = false;
    } else if (cp.kind === "invalid") {
      issues.push({
        severity: "error",
        cell: ref(r, 1),
        message: cp.excelError
          ? `Probability cell contains the Excel error ${cp.text}.`
          : `Probability "${cp.text}" is not a number.`,
      });
      rowOk = false;
    }
    if (rowOk && cx.kind === "number" && cp.kind === "number") {
      raw.push({ x: cx.value, p: cp.value, percentText: cp.percentText, r });
    }
  }

  if (raw.length === 0 && !issues.some((i) => i.severity === "error")) {
    return fail(hasHeader ? "No data rows were found below the header row." : "No data rows were found.");
  }
  if (blankRows > 0) {
    issues.push({ severity: "info", message: `${blankRows} blank row${blankRows === 1 ? " was" : "s were"} ignored.` });
  }

  // Pass 2: decide the probability scale.
  let probabilityFormat: "decimal" | "percent" = "decimal";
  const requested = options.probabilityFormat ?? "auto";
  const plain = raw.filter((row) => !row.percentText);
  if (requested === "percent") {
    probabilityFormat = "percent";
  } else if (requested === "auto" && plain.length > 0 && plain.some((row) => row.p > 1)) {
    const total = stableSum(raw.map((row) => (row.percentText ? row.p * 100 : row.p)));
    if (plain.every((row) => row.p >= 0 && row.p <= 100) && Math.abs(total - 100) <= tolerance * 100) {
      probabilityFormat = "percent";
      issues.push({
        severity: "info",
        message: "Probabilities look like percentage points (they sum to 100), so they were divided by 100.",
      });
    }
  }

  const rows: ParsedRow[] = raw.map((row) => ({
    x: row.x,
    p: probabilityFormat === "percent" && !row.percentText ? cleanFloat(row.p / 100, 15) : row.p,
    cell: ref(row.r, 0),
  }));

  // Pass 3: range checks.
  raw.forEach((row, i) => {
    const p = rows[i].p;
    if (p < 0) {
      issues.push({ severity: "error", cell: ref(row.r, 1), message: `Probability ${fmt(p)} is negative.` });
    } else if (p > 1 + tolerance) {
      issues.push({
        severity: "error",
        cell: ref(row.r, 1),
        message: `Probability ${fmt(p)} is greater than 1 (100%).`,
      });
    } else if (p === 0) {
      issues.push({
        severity: "warning",
        cell: ref(row.r, 0),
        message: `Outcome ${fmt(rows[i].x)} has zero probability; it is kept for statistics but adds no step.`,
      });
    }
  });

  const probabilitySum = stableSum(rows.map((r) => r.p));
  if (rows.length > 0 && Math.abs(probabilitySum - 1) > tolerance) {
    issues.push({
      severity: "error",
      message: `Probabilities sum to ${fmt(probabilitySum)}, not 1 (difference ${fmt(probabilitySum - 1)}; tolerance ${fmt(tolerance)}).`,
    });
  }

  // Pass 4: merge duplicates and sort.
  const byX = new Map<number, { p: number[]; cells: string[] }>();
  for (const row of rows) {
    const entry = byX.get(row.x) ?? { p: [], cells: [] };
    entry.p.push(row.p);
    entry.cells.push(row.cell);
    byX.set(row.x, entry);
  }
  for (const [x, entry] of byX) {
    if (entry.cells.length > 1) {
      issues.push({
        severity: "warning",
        message: `Outcome ${fmt(x)} appears ${entry.cells.length} times (${entry.cells.join(", ")}); its probabilities were combined.`,
      });
    }
  }
  const outcomes: Outcome[] = Array.from(byX, ([x, entry]) => ({ x, p: cleanFloat(stableSum(entry.p), 15) })).sort(
    (a, b) => a.x - b.x,
  );
  if (rows.some((r, i) => i > 0 && r.x < rows[i - 1].x)) {
    issues.push({ severity: "info", message: "Outcomes were sorted in ascending order." });
  }

  const ok = !issues.some((i) => i.severity === "error");
  return {
    ok,
    hasHeader,
    headerDetected,
    headers,
    rows,
    distribution: ok ? { outcomes } : null,
    probabilitySum,
    probabilityFormat,
    issues,
  };
}

/**
 * In-memory fake of the subset of the Excel JavaScript API used by
 * src/excel/*. It enforces Office.js's load/sync contract: reading a scalar
 * property (values, rowCount, id, name, isNullObject, items, …) throws unless
 * it was requested with load() (or is a null-object check) and a
 * context.sync() has happened since. This catches the most common Office.js
 * bug, reading proxy data before syncing, in plain unit tests.
 */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { qualify, rangeRef } from "../../src/core/a1";

type Cell = { value: string | number | boolean; format: string; bold?: boolean };

let pending = new Set<Loadable>();
let idCounter = 0;

export class PropertyNotLoaded extends Error {}

class Loadable {
  private _loaded = new Set<string>();
  private _requested = new Set<string>();
  /** Item properties requested through a collection, e.g. load("items/name"). */
  private _itemProps = new Set<string>();
  load(props?: string | string[]): this {
    const list = props === undefined ? ["*"] : Array.isArray(props) ? props : props.split(",");
    for (const p of list) {
      const [head, ...rest] = p.trim().split("/");
      this._requested.add(head);
      if (head === "items") this._itemProps.add(rest.length ? rest.join("/") : "*");
    }
    pending.add(this);
    return this;
  }
  /** Items of a collection (overridden by collections). */
  _children(): Loadable[] {
    return [];
  }
  /** Collection items as of the last sync that loaded them (Office.js semantics). */
  protected _snapshot: any[] = [];
  _markLoaded(prop: string) {
    this._loaded.add(prop);
  }
  /** Called by sync(). */
  _flush() {
    if (this._requested.has("items")) this._snapshot = [...this._children()];
    for (const p of this._requested) this._loaded.add(p);
    this._requested.clear();
    if (this._itemProps.size) {
      for (const child of this._children()) for (const prop of this._itemProps) child._markLoaded(prop);
      this._itemProps.clear();
    }
  }
  protected need(prop: string) {
    if (!this._loaded.has(prop) && !this._loaded.has("*")) {
      throw new PropertyNotLoaded(`Property '${prop}' was read before load('${prop}') + sync() on ${this.constructor.name}`);
    }
  }
  /** Null-object checks become readable after the next sync without an explicit load. */
  protected requestNullCheck() {
    this._requested.add("isNullObject");
    pending.add(this);
  }
}

function parseA1(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim());
  if (!m) throw new Error(`Bad cell reference ${ref}`);
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

// ------------------------------------------------------------------ workbook

export class FakeWorkbook extends Loadable {
  sheetsList: FakeSheet[] = [];
  activeSheet!: FakeSheet;
  selection: { sheet: FakeSheet; rects: { row: number; col: number; rows: number; cols: number }[] } | null = null;
  settingsMap = new Map<string, any>();
  tablesList: FakeTable[] = [];
  namesMap = new Map<string, FakeRange>();
  worksheets = new FakeWorksheetCollection(this);
  tables = new FakeTableCollection(this, null);
  names = new FakeNamedItemCollection(this);
  settings = new FakeSettingCollection(this);

  addSheet(name: string): FakeSheet {
    const sheet = new FakeSheet(this, name);
    this.sheetsList.push(sheet);
    if (!this.activeSheet) this.activeSheet = sheet;
    return sheet;
  }
  sheetByName(name: string) {
    return this.sheetsList.find((s) => s._name.toLowerCase() === name.toLowerCase()) ?? null;
  }
  select(sheet: FakeSheet, address: string) {
    this.activeSheet = sheet;
    this.selection = { sheet, rects: address.split(",").map((a) => sheet.rectOf(a)) };
  }
  getSelectedRange(): FakeRange {
    if (!this.selection) throw new Error("No selection");
    if (this.selection.rects.length > 1) throw new Error("InvalidSelection: multiple areas");
    const r = this.selection.rects[0];
    return new FakeRange(this.selection.sheet, r.row, r.col, r.rows, r.cols);
  }
  getSelectedRanges() {
    const sel = this.selection!;
    return new (class extends Loadable {
      get areaCount() {
        this.need("areaCount");
        return sel.rects.length;
      }
    })();
  }
}

class FakeWorksheetCollection extends Loadable {
  constructor(private wb: FakeWorkbook) {
    super();
  }
  _children() {
    return this.wb.sheetsList;
  }
  get items() {
    this.need("items");
    return this._snapshot as FakeSheet[];
  }
  getActiveWorksheet() {
    return this.wb.activeSheet;
  }
  getItem(key: string) {
    const s = this.wb.sheetsList.find((x) => x.id === key) ?? this.wb.sheetByName(key);
    if (!s) throw new Error(`ItemNotFound: worksheet ${key}`);
    return s;
  }
  getItemOrNullObject(key: string) {
    const s = this.wb.sheetsList.find((x) => x.id === key) ?? this.wb.sheetByName(key);
    return s ? s.asLookup() : FakeSheet.nullSheet(this.wb);
  }
  add(name: string) {
    if (this.wb.sheetByName(name)) throw new Error(`ItemAlreadyExists: ${name}`);
    return this.wb.addSheet(name);
  }
}

export class FakeSheet extends Loadable {
  readonly _id = `{sheet-${++idCounter}}`;
  cells = new Map<string, Cell>();
  chartsList: FakeChart[] = [];
  private _null = false;
  tables: FakeTableCollection;
  charts: FakeChartCollection;
  constructor(
    public wb: FakeWorkbook,
    public _name: string,
  ) {
    super();
    this.tables = new FakeTableCollection(wb, this);
    this.charts = new FakeChartCollection(this);
  }
  static nullSheet(wb: FakeWorkbook) {
    const s = new FakeSheet(wb, "<null>");
    s._null = true;
    s.requestNullCheck();
    return s;
  }
  asLookup() {
    this.requestNullCheck();
    return this;
  }
  get isNullObject() {
    this.need("isNullObject");
    return this._null;
  }
  get id() {
    this.need("id");
    return this._id;
  }
  get name() {
    this.need("name");
    return this._name;
  }
  rectOf(address: string) {
    const [a, b] = address.split(":");
    const s = parseA1(a);
    const e = b ? parseA1(b) : s;
    return { row: s.row, col: s.col, rows: e.row - s.row + 1, cols: e.col - s.col + 1 };
  }
  getRange(address: string) {
    const r = this.rectOf(address);
    return new FakeRange(this, r.row, r.col, r.rows, r.cols);
  }
  getRangeByIndexes(row: number, col: number, rows: number, cols: number) {
    return new FakeRange(this, row, col, rows, cols);
  }
  getCell(row: number, col: number) {
    return new FakeRange(this, row, col, 1, 1);
  }
  getUsedRangeOrNullObject(valuesOnly = false) {
    let minR = Infinity,
      minC = Infinity,
      maxR = -1,
      maxC = -1;
    for (const [key, cell] of this.cells) {
      if (valuesOnly && cell.value === "") continue;
      const [r, c] = key.split(",").map(Number);
      minR = Math.min(minR, r);
      minC = Math.min(minC, c);
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
    const range = maxR < 0 ? FakeRange.nullRange(this) : new FakeRange(this, minR, minC, maxR - minR + 1, maxC - minC + 1);
    range.markLookup();
    return range;
  }
  activate() {
    this.wb.activeSheet = this;
  }
  // test helpers
  set(address: string, rows: (string | number)[][]) {
    const r = this.rectOf(address);
    rows.forEach((row, i) => row.forEach((v, j) => this.put(r.row + i, r.col + j, v)));
  }
  put(row: number, col: number, value: string | number | boolean) {
    const key = `${row},${col}`;
    const existing = this.cells.get(key);
    if (value === "" && !existing) return;
    this.cells.set(key, { format: "General", ...existing, value });
  }
  read(row: number, col: number) {
    return this.cells.get(`${row},${col}`)?.value ?? "";
  }
  readBlock(address: string) {
    const r = this.rectOf(address);
    return Array.from({ length: r.rows }, (_, i) => Array.from({ length: r.cols }, (_, j) => this.read(r.row + i, r.col + j)));
  }
}

// ------------------------------------------------------------------ ranges

export class FakeRange extends Loadable {
  private _null = false;
  constructor(
    public sheet: FakeSheet,
    public r: number,
    public c: number,
    public nr: number,
    public nc: number,
  ) {
    super();
  }
  static nullRange(sheet: FakeSheet) {
    const x = new FakeRange(sheet, 0, 0, 0, 0);
    x._null = true;
    return x;
  }
  markLookup() {
    this.requestNullCheck();
  }
  get isNullObject() {
    this.need("isNullObject");
    return this._null;
  }
  get worksheet() {
    return this.sheet;
  }
  get rowIndex() {
    this.need("rowIndex");
    return this.r;
  }
  get columnIndex() {
    this.need("columnIndex");
    return this.c;
  }
  get rowCount() {
    this.need("rowCount");
    return this.nr;
  }
  get columnCount() {
    this.need("columnCount");
    return this.nc;
  }
  get address() {
    this.need("address");
    return qualify(this.sheet._name, rangeRef(this.r, this.c, this.nr, this.nc));
  }
  get left() {
    this.need("left");
    return this.c * 64;
  }
  get top() {
    this.need("top");
    return this.r * 20;
  }
  get width() {
    this.need("width");
    return this.nc * 64;
  }
  get height() {
    this.need("height");
    return this.nr * 20;
  }
  private grid<T>(f: (row: number, col: number) => T): T[][] {
    return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => f(this.r + i, this.c + j)));
  }
  get values() {
    this.need("values");
    return this.grid((r, c) => this.sheet.read(r, c));
  }
  set values(v: any[][]) {
    if (v.length !== this.nr || v.some((row) => row.length !== this.nc)) {
      throw new Error(`InvalidArgument: values are ${v.length}x${v[0]?.length}, range is ${this.nr}x${this.nc}`);
    }
    v.forEach((row, i) => row.forEach((val, j) => this.sheet.put(this.r + i, this.c + j, val)));
  }
  get formulas() {
    this.need("formulas");
    return this.grid((r, c) => this.sheet.read(r, c));
  }
  get numberFormat() {
    this.need("numberFormat");
    return this.grid((r, c) => this.sheet.cells.get(`${r},${c}`)?.format ?? "General");
  }
  set numberFormat(v: any[][]) {
    if (v.length !== this.nr) throw new Error("InvalidArgument: numberFormat dimensions");
    v.forEach((row, i) =>
      row.forEach((f, j) => {
        const key = `${this.r + i},${this.c + j}`;
        this.sheet.cells.set(key, { value: "", ...this.sheet.cells.get(key), format: f });
      }),
    );
  }
  get format() {
    const self = this;
    return {
      font: {
        set bold(b: boolean) {
          for (let i = 0; i < self.nr; i++)
            for (let j = 0; j < self.nc; j++) {
              const key = `${self.r + i},${self.c + j}`;
              self.sheet.cells.set(key, { value: "", format: "General", ...self.sheet.cells.get(key), bold: b });
            }
        },
        size: 11,
      },
      autofitColumns() {},
      columnWidth: 64,
    };
  }
  getIntersectionOrNullObject(other: FakeRange) {
    const r1 = Math.max(this.r, other.r);
    const c1 = Math.max(this.c, other.c);
    const r2 = Math.min(this.r + this.nr, other.r + other.nr);
    const c2 = Math.min(this.c + this.nc, other.c + other.nc);
    const out = r2 <= r1 || c2 <= c1 ? FakeRange.nullRange(this.sheet) : new FakeRange(this.sheet, r1, c1, r2 - r1, c2 - c1);
    out.markLookup();
    return out;
  }
  getSurroundingRegion() {
    let { r, c, nr, nc } = this;
    const filled = (row: number, col: number) => row >= 0 && col >= 0 && this.sheet.read(row, col) !== "";
    let grew = true;
    while (grew) {
      grew = false;
      const span = (fixed: number, from: number, to: number, isRow: boolean) => {
        for (let k = from; k <= to; k++) if (isRow ? filled(fixed, k) : filled(k, fixed)) return true;
        return false;
      };
      if (span(r - 1, c - 1, c + nc, true)) (r--, nr++, (grew = true));
      if (span(r + nr, c - 1, c + nc, true)) (nr++, (grew = true));
      if (span(c - 1, r - 1, r + nr, false)) (c--, nc++, (grew = true));
      if (span(c + nc, r - 1, r + nr, false)) (nc++, (grew = true));
    }
    return new FakeRange(this.sheet, r, c, nr, nc);
  }
  getOffsetRange(dr: number, dc: number) {
    return new FakeRange(this.sheet, this.r + dr, this.c + dc, this.nr, this.nc);
  }
  getResizedRange(dr: number, dc: number) {
    return new FakeRange(this.sheet, this.r, this.c, this.nr + dr, this.nc + dc);
  }
  clear() {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.cells.delete(`${this.r + i},${this.c + j}`);
  }
  select() {
    this.sheet.wb.select(this.sheet, rangeRef(this.r, this.c, this.nr, this.nc));
  }
}

// ------------------------------------------------------------------ tables & names

class FakeTableCollection extends Loadable {
  constructor(
    private wb: FakeWorkbook,
    private sheet: FakeSheet | null,
  ) {
    super();
  }
  _children() {
    return this.wb.tablesList.filter((t) => !this.sheet || t.sheet === this.sheet);
  }
  get items() {
    this.need("items");
    return this._snapshot as FakeTable[];
  }
  add(range: FakeRange, hasHeaders: boolean) {
    if (!hasHeaders) throw new Error("fake supports header tables only");
    for (const t of this.wb.tablesList) {
      if (t.sheet === range.sheet && overlaps(t.range, range)) {
        throw new Error("InvalidOperation: overlaps an existing table");
      }
    }
    const t = new FakeTable(this.wb, range.sheet, range, `Table${++idCounter}`);
    this.wb.tablesList.push(t);
    return t;
  }
  getItemOrNullObject(name: string) {
    const t = this.wb.tablesList.find((x) => x._name.toLowerCase() === name.toLowerCase());
    const out = t ?? FakeTable.nullTable(this.wb);
    out.lookup();
    return out;
  }
}

function overlaps(a: FakeRange, b: FakeRange) {
  return a.r < b.r + b.nr && b.r < a.r + a.nr && a.c < b.c + b.nc && b.c < a.c + a.nc;
}

export class FakeTable extends Loadable {
  private _null = false;
  showBandedRows = true;
  constructor(
    private wb: FakeWorkbook,
    public sheet: FakeSheet,
    public range: FakeRange,
    public _name: string,
  ) {
    super();
  }
  static nullTable(wb: FakeWorkbook) {
    const t = new FakeTable(wb, wb.sheetsList[0], FakeRange.nullRange(wb.sheetsList[0]), "<null>");
    t._null = true;
    return t;
  }
  lookup() {
    this.requestNullCheck();
  }
  get isNullObject() {
    this.need("isNullObject");
    return this._null;
  }
  get name() {
    this.need("name");
    return this._name;
  }
  set name(n: string) {
    const clash = this.wb.tablesList.some((t) => t !== this && t._name.toLowerCase() === n.toLowerCase()) || this.wb.namesMap.has(n.toLowerCase());
    if (clash) throw new Error(`InvalidArgument: name ${n} already exists`);
    this._name = n;
  }
  get worksheet() {
    return this.sheet;
  }
  getRange() {
    return new FakeRange(this.sheet, this.range.r, this.range.c, this.range.nr, this.range.nc);
  }
  delete() {
    this.wb.tablesList = this.wb.tablesList.filter((t) => t !== this);
    this.range.clear();
  }
}

class FakeNamedItemCollection extends Loadable {
  constructor(private wb: FakeWorkbook) {
    super();
  }
  get items() {
    this.need("items");
    return Array.from(this.wb.namesMap.keys()).map((name) => ({ name }));
  }
  add(name: string, range: FakeRange) {
    if (this.wb.namesMap.has(name.toLowerCase()) || this.wb.tablesList.some((t) => t._name.toLowerCase() === name.toLowerCase())) {
      throw new Error(`InvalidArgument: name ${name} exists`);
    }
    this.wb.namesMap.set(name.toLowerCase(), range);
    return {};
  }
  getItemOrNullObject(name: string) {
    const wb = this.wb;
    const exists = wb.namesMap.has(name.toLowerCase());
    const item = new (class extends Loadable {
      get isNullObject() {
        this.need("isNullObject");
        return !exists;
      }
      delete() {
        wb.namesMap.delete(name.toLowerCase());
      }
    })();
    (item as any).requestNullCheck();
    return item;
  }
}

class FakeSettingCollection {
  constructor(private wb: FakeWorkbook) {}
  add(key: string, value: any) {
    this.wb.settingsMap.set(key, value);
  }
  getItemOrNullObject(key: string) {
    const wb = this.wb;
    const item = new (class extends Loadable {
      get isNullObject() {
        this.need("isNullObject");
        return !wb.settingsMap.has(key);
      }
      get value() {
        this.need("value");
        return wb.settingsMap.get(key);
      }
    })();
    (item as any).requestNullCheck();
    return item;
  }
}

// ------------------------------------------------------------------ charts

class FakeChartCollection extends Loadable {
  constructor(private sheet: FakeSheet) {
    super();
  }
  _children() {
    return this.sheet.chartsList;
  }
  get items() {
    this.need("items");
    return this._snapshot as FakeChart[];
  }
  add(type: string, source: FakeRange) {
    const chart = new FakeChart(this.sheet, type);
    chart.seriesList.push(new FakeSeries(chart, "Series1", type, null, source));
    this.sheet.chartsList.push(chart);
    return chart;
  }
  getItemOrNullObject(name: string) {
    const c = this.sheet.chartsList.find((x) => x._name === name) ?? FakeChart.nullChart(this.sheet);
    c.lookup();
    return c;
  }
}

class FakeAxis {
  minimum: number | undefined;
  maximum: number | undefined;
  majorUnit: number | undefined;
  numberFormat = "General";
  position = "Automatic";
  title = { text: "", visible: false };
  majorGridlines = { visible: false };
}

export class FakeChart extends Loadable {
  _name = "Chart 1";
  private _null = false;
  seriesList: FakeSeries[] = [];
  title = { text: "", visible: false };
  legend = { visible: true, position: "Right" };
  axes = { valueAxis: new FakeAxis(), categoryAxis: new FakeAxis() };
  position: [string, string?] | null = null;
  private _left = 0;
  _top = 0;
  _width = 360;
  _height = 216;
  series: any;
  constructor(
    public sheet: FakeSheet,
    public type: string,
  ) {
    super();
    const chart = this;
    this.series = new (class extends Loadable {
      _children() {
        return chart.seriesList;
      }
      get items() {
        this.need("items");
        return this._snapshot as FakeSeries[];
      }
      getItemAt(i: number) {
        if (!chart.seriesList[i]) throw new Error(`ItemNotFound: series ${i}`);
        return chart.seriesList[i];
      }
      add(name: string) {
        const s = new FakeSeries(chart, name, chart.type, null, null);
        chart.seriesList.push(s);
        return s;
      }
    })();
  }
  static nullChart(sheet: FakeSheet) {
    const c = new FakeChart(sheet, "");
    c._null = true;
    return c;
  }
  lookup() {
    this.requestNullCheck();
  }
  get isNullObject() {
    this.need("isNullObject");
    return this._null;
  }
  get name() {
    this.need("name");
    return this._name;
  }
  set name(n: string) {
    this._name = n;
  }
  get worksheet() {
    return this.sheet;
  }
  get left() {
    this.need("left");
    return this._left;
  }
  set left(v: number) {
    this._left = v;
  }
  get top() {
    this.need("top");
    return this._top;
  }
  get width() {
    this.need("width");
    return this._width;
  }
  get height() {
    this.need("height");
    return this._height;
  }
  setPosition(start: string, end?: string) {
    this.position = [start, end];
    const s = this.sheet.rectOf(start);
    this._left = s.col * 64;
    this._top = s.row * 20;
  }
  delete() {
    this.sheet.chartsList = this.sheet.chartsList.filter((c) => c !== this);
  }
}

export class FakeSeries extends Loadable {
  smooth = true;
  markerStyle = "Automatic";
  format = { line: { weight: 0.75 } };
  constructor(
    private chart: FakeChart,
    public _name: string,
    public chartType: string,
    public x: FakeRange | null,
    public y: FakeRange | null,
  ) {
    super();
  }
  get name() {
    return this._name;
  }
  set name(n: string) {
    this._name = n;
  }
  setXAxisValues(r: FakeRange) {
    this.x = r;
  }
  setValues(r: FakeRange) {
    this.y = r;
  }
  delete() {
    this.chart.seriesList = this.chart.seriesList.filter((s) => s !== this);
  }
}

// ------------------------------------------------------------------ globals

export interface FakeEnv {
  workbook: FakeWorkbook;
  syncCount: () => number;
}

/** Install fake `Excel`, `Office` and `OfficeExtension` globals. */
export function installFakeExcel(options: { apiVersion?: string } = {}): FakeEnv {
  pending = new Set();
  const workbook = new FakeWorkbook();
  let syncs = 0;
  const maxMinor = Number((options.apiVersion ?? "1.19").split(".")[1]);
  const context = {
    workbook,
    sync: async () => {
      syncs++;
      const batch = pending;
      pending = new Set();
      batch.forEach((o) => o._flush());
    },
  };
  const g = globalThis as any;
  g.Excel = {
    run: async (fn: (ctx: any) => Promise<any>) => fn(context),
    ChartType: { xyscatterLines: "XYScatterLines", xyscatterLinesNoMarkers: "XYScatterLinesNoMarkers" },
    ChartSeriesBy: { columns: "Columns" },
    ClearApplyTo: { all: "All" },
    ChartMarkerStyle: { none: "None", automatic: "Automatic" },
    ChartLegendPosition: { bottom: "Bottom" },
    ChartAxisPosition: { minimum: "Minimum" },
  };
  g.OfficeExtension = { Error: class OfficeError extends Error {} };
  g.Office = {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) => name === "ExcelApi" && Number(version.split(".")[1]) <= maxMinor,
      },
    },
  };
  return { workbook, syncCount: () => syncs };
}

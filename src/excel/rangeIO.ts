/** Reading input ranges from the workbook. */
import { cellRef, qualify, splitAddress } from "../core/a1";
import { CellValue } from "../core/parse";
import { SourceRef } from "./model";
import { excelApi, UserError } from "./support";

/** Hard cap so an accidental whole-sheet selection cannot freeze the pane. */
const MAX_ROWS = 50000;

export interface RangeSnapshot {
  source: SourceRef;
  /** Sheet-qualified address for display. */
  displayAddress: string;
  rowIndex: number;
  columnIndex: number;
  values: CellValue[][];
  /** Number format of the first outcome cell below any header. */
  xNumberFormat: string;
  /** Text in the cell directly above the range (a likely distribution name). */
  labelAbove: string | null;
  /** True when a single selected cell was expanded to its surrounding region. */
  expanded: boolean;
}

export interface CellSnapshot {
  sheetId: string;
  sheetName: string;
  row: number;
  col: number;
  displayAddress: string;
}

async function snapshot(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  range: Excel.Range,
  expanded: boolean,
): Promise<RangeSnapshot> {
  // Trim whole-column / whole-sheet selections to the used area.
  const used = sheet.getUsedRangeOrNullObject(true);
  await context.sync();
  if (used.isNullObject) throw new UserError("The selected sheet is empty.");
  const trimmed = range.getIntersectionOrNullObject(used);
  trimmed.load("rowCount,columnCount");
  await context.sync();
  if (trimmed.isNullObject) throw new UserError("The selected range contains no data.");
  if (trimmed.rowCount > MAX_ROWS) {
    throw new UserError(`The selection has ${trimmed.rowCount.toLocaleString()} rows; the limit is ${MAX_ROWS.toLocaleString()}.`);
  }

  trimmed.load("values,numberFormat,address,rowIndex,columnIndex,rowCount,columnCount");
  sheet.load("id,name");
  await context.sync();

  let labelAbove: string | null = null;
  if (trimmed.rowIndex > 0) {
    const above = sheet.getCell(trimmed.rowIndex - 1, trimmed.columnIndex);
    above.load("values");
    await context.sync();
    const v = above.values[0][0];
    if (typeof v === "string" && v.trim()) labelAbove = v.trim();
  }

  const values = trimmed.values as CellValue[][];
  const firstNumericRow = values.findIndex((row) => typeof row[0] === "number");
  const xNumberFormat = firstNumericRow >= 0 ? String(trimmed.numberFormat[firstNumericRow][0] ?? "General") : "General";
  const local = splitAddress(trimmed.address).local;

  return {
    source: { sheetId: sheet.id, sheetName: sheet.name, address: local },
    displayAddress: qualify(sheet.name, local),
    rowIndex: trimmed.rowIndex,
    columnIndex: trimmed.columnIndex,
    values,
    xNumberFormat,
    labelAbove,
    expanded,
  };
}

/**
 * Read the active selection. A single selected cell is expanded to its
 * surrounding region (the contiguous block around it).
 */
export async function readSelection(): Promise<RangeSnapshot> {
  return Excel.run(async (context) => {
    if (excelApi("1.9")) {
      const areas = context.workbook.getSelectedRanges();
      areas.load("areaCount");
      await context.sync();
      if (areas.areaCount > 1) {
        throw new UserError("Select a single contiguous two-column range (multiple areas are selected).");
      }
    }
    let range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    await context.sync();
    let expanded = false;
    if (range.rowCount === 1 && range.columnCount === 1) {
      range = range.getSurroundingRegion();
      expanded = true;
      // A title row ("Make Component" with an empty second cell) above the
      // data is not part of the distribution; it becomes the suggested name.
      range.load("values,rowCount,columnCount");
      await context.sync();
      const first = range.values[0];
      if (range.rowCount > 2 && range.columnCount === 2 && typeof first[0] === "string" && first[0].trim() && first[1] === "") {
        range = range.getOffsetRange(1, 0).getResizedRange(-1, 0);
      }
    }
    return snapshot(context, sheet, range, expanded);
  });
}

/** Re-read a stored source range (by worksheet id, so renamed sheets still work). */
export async function readSource(source: SourceRef): Promise<RangeSnapshot> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(source.sheetId);
    await context.sync();
    if (sheet.isNullObject) {
      throw new UserError(`The source worksheet "${source.sheetName}" no longer exists.`);
    }
    return snapshot(context, sheet, sheet.getRange(source.address), false);
  });
}

/** The top-left cell of the current selection (for "user-selected cell" output). */
export async function readSelectedCell(): Promise<CellSnapshot> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("rowIndex,columnIndex");
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("id,name");
    await context.sync();
    return {
      sheetId: sheet.id,
      sheetName: sheet.name,
      row: range.rowIndex,
      col: range.columnIndex,
      displayAddress: qualify(sheet.name, cellRef(range.rowIndex, range.columnIndex)),
    };
  });
}

export async function getActiveSheet(): Promise<{ id: string; name: string }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("id,name");
    await context.sync();
    return { id: sheet.id, name: sheet.name };
  });
}

/** Activate the sheet and select a stored range. */
export async function selectRange(sheetId: string, address: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetId);
    await context.sync();
    if (sheet.isNullObject) throw new UserError("That worksheet no longer exists.");
    sheet.activate();
    sheet.getRange(address).select();
    await context.sync();
  });
}

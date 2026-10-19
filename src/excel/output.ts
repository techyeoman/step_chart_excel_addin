/** Writing CDF coordinate tables to the workbook. */
import { cellRef, qualify, Rect, rectContains, rangeRef, sanitizeExcelName, sanitizeSheetName, uniqueName } from "../core/a1";
import { Point } from "../core/types";
import { DistributionEntry, OutputBlock, Placement } from "./model";
import { UserError } from "./support";

const MAX_ROW = 1048576;
const MAX_COL = 16384;
export const X_HEADER = "X (Outcome)";
export const Y_HEADER = "Y (Cumulative Probability)";

/** Prevent text that starts like a formula from being evaluated. */
export function asText(s: string): string {
  return /^[=+\-@]/.test(s) ? `'${s}` : s;
}

async function nextFreeAnchor(context: Excel.RequestContext, sheet: Excel.Worksheet): Promise<{ row: number; col: number }> {
  const used = sheet.getUsedRangeOrNullObject(false);
  used.load("rowIndex,columnIndex,columnCount");
  await context.sync();
  if (used.isNullObject) return { row: 0, col: 0 };
  // Leave one blank column between existing content and the new block.
  return { row: used.rowIndex, col: used.columnIndex + used.columnCount + 1 };
}

async function resolvePlacement(
  context: Excel.RequestContext,
  placement: Placement,
): Promise<{ sheet: Excel.Worksheet; row: number; col: number }> {
  const sheets = context.workbook.worksheets;
  if (placement.kind === "newSheet") {
    const name = sanitizeSheetName(placement.sheetName);
    let sheet = sheets.getItemOrNullObject(name);
    await context.sync();
    if (sheet.isNullObject) {
      sheet = sheets.add(name);
      await context.sync();
      return { sheet, row: 0, col: 0 };
    }
    const anchor = await nextFreeAnchor(context, sheet);
    return { sheet, row: 0, col: anchor.col };
  }

  const sheet = sheets.getItemOrNullObject(placement.sheetId);
  await context.sync();
  if (sheet.isNullObject) {
    if (placement.kind === "cell") throw new UserError("The worksheet chosen for the output cell no longer exists.");
    const active = sheets.getActiveWorksheet();
    return { sheet: active, ...(await nextFreeAnchor(context, active)) };
  }
  if (placement.kind === "cell") return { sheet, row: placement.row, col: placement.col };
  return { sheet, ...(await nextFreeAnchor(context, sheet)) };
}

/** Remove a previously written block (its table / defined name and cell contents). */
export async function clearBlock(context: Excel.RequestContext, block: OutputBlock): Promise<void> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(block.sheetId);
  await context.sync();
  if (sheet.isNullObject) return; // the sheet (and everything on it) is already gone

  if (block.tableName) {
    const table = context.workbook.tables.getItemOrNullObject(block.tableName);
    await context.sync();
    if (!table.isNullObject) {
      // Only delete the table if it is still the one we created (same sheet and anchor).
      table.worksheet.load("id");
      const r = table.getRange();
      r.load("rowIndex,columnIndex");
      await context.sync();
      const headerRow = block.dataRow - 1;
      if (table.worksheet.id === block.sheetId && r.rowIndex === headerRow && r.columnIndex === block.col) {
        table.delete();
      }
    }
  }
  if (block.definedName) {
    const named = context.workbook.names.getItemOrNullObject(block.definedName);
    await context.sync();
    if (!named.isNullObject) named.delete();
  }
  sheet.getRangeByIndexes(block.row, block.col, block.rows, block.cols).clear(Excel.ClearApplyTo.all);
  await context.sync();
}

async function existingNames(context: Excel.RequestContext): Promise<string[]> {
  const tables = context.workbook.tables;
  const names = context.workbook.names;
  tables.load("items/name");
  names.load("items/name");
  await context.sync();
  return [...tables.items.map((t) => t.name), ...names.items.map((n) => n.name)];
}

/**
 * Write (or rewrite in place) the coordinate block for one distribution.
 *
 * Layout with headers:      Layout without headers:
 *   row 0  <name>             X  Y   (defined name covers the data)
 *   row 1  X (Outcome) | Y (Cumulative Probability)   <- Excel table header
 *   row 2+ x | y
 */
export async function writeBlock(
  context: Excel.RequestContext,
  entry: DistributionEntry,
  points: Point[],
): Promise<OutputBlock> {
  const headers = entry.includeHeaders;
  const rows = points.length + (headers ? 2 : 0);
  const cols = 2;

  // 1. Where does the block go? Reuse the previous location when possible.
  let sheet: Excel.Worksheet | null = null;
  let row = 0;
  let col = 0;
  let oldRect: Rect | null = null;
  if (entry.output) {
    const previous = context.workbook.worksheets.getItemOrNullObject(entry.output.sheetId);
    await context.sync();
    if (!previous.isNullObject) {
      sheet = previous;
      row = entry.output.row;
      col = entry.output.col;
      oldRect = { row, col, rows: entry.output.rows, cols: entry.output.cols };
    }
  }
  if (!sheet) ({ sheet, row, col } = await resolvePlacement(context, entry.placement));
  sheet.load("id,name");
  await context.sync();

  if (row + rows > MAX_ROW || col + cols > MAX_COL) {
    throw new UserError(`There is not enough room on "${sheet.name}" for the output of "${entry.name}".`);
  }

  // 2. Refuse to overwrite anything that is not our own previous output.
  const target = sheet.getRangeByIndexes(row, col, rows, cols);
  target.load("formulas");
  await context.sync();
  const conflicts: string[] = [];
  target.formulas.forEach((r, i) =>
    r.forEach((v, j) => {
      const occupied = v !== "" && v !== null;
      if (occupied && !(oldRect && rectContains(oldRect, row + i, col + j))) conflicts.push(cellRef(row + i, col + j));
    }),
  );
  if (conflicts.length) {
    const list = conflicts.slice(0, 5).join(", ") + (conflicts.length > 5 ? ", …" : "");
    throw new UserError(
      `Output for "${entry.name}" needs ${qualify(sheet.name, rangeRef(row, col, rows, cols))}, but ${list} ` +
        `${conflicts.length === 1 ? "is" : "are"} not empty. Choose another output location or clear those cells.`,
    );
  }

  // 3. Remove the old block, then write the new one.
  if (entry.output) await clearBlock(context, entry.output);
  const ownNames = new Set([entry.output?.tableName, entry.output?.definedName].filter(Boolean) as string[]);
  const taken = (await existingNames(context)).filter((n) => !ownNames.has(n));
  const name = uniqueName(sanitizeExcelName(entry.outputName || `CDF_${entry.name}`), taken);

  const values: (string | number)[][] = [];
  if (headers) {
    values.push([asText(entry.name), ""]);
    values.push([X_HEADER, Y_HEADER]);
  }
  for (const p of points) values.push([p.x, p.y]);
  target.values = values;

  const dataRow = row + (headers ? 2 : 0);
  const data = sheet.getRangeByIndexes(dataRow, col, points.length, cols);
  if (entry.xNumberFormat && entry.xNumberFormat !== "General") {
    sheet.getRangeByIndexes(dataRow, col, points.length, 1).numberFormat = points.map(() => [entry.xNumberFormat]);
  }

  let tableName: string | null = null;
  let definedName: string | null = null;
  if (headers) {
    sheet.getRangeByIndexes(row, col, 1, 1).format.font.bold = true;
    const table = sheet.tables.add(sheet.getRangeByIndexes(row + 1, col, points.length + 1, cols), true);
    table.name = name;
    table.showBandedRows = true;
    tableName = name;
  } else {
    context.workbook.names.add(name, data, `CDF step coordinates for ${entry.name}`);
    definedName = name;
  }
  target.format.autofitColumns();
  await context.sync();

  return {
    sheetId: sheet.id,
    sheetName: sheet.name,
    row,
    col,
    rows,
    cols,
    dataRow,
    dataCount: points.length,
    tableName,
    definedName,
  };
}

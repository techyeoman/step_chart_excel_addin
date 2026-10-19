/** Pure A1-reference and Excel naming helpers (no Office.js dependency). */

/** 0-based column index -> column letters (0 -> A, 25 -> Z, 26 -> AA). */
export function columnLetter(index: number): string {
  if (!Number.isInteger(index) || index < 0) throw new RangeError(`Invalid column index ${index}`);
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 0-based row/column -> "B7". */
export function cellRef(rowIndex: number, columnIndex: number): string {
  return `${columnLetter(columnIndex)}${rowIndex + 1}`;
}

/** 0-based rectangle -> "B7:C12" (single cell -> "B7"). */
export function rangeRef(rowIndex: number, columnIndex: number, rowCount: number, columnCount: number): string {
  const start = cellRef(rowIndex, columnIndex);
  if (rowCount === 1 && columnCount === 1) return start;
  return `${start}:${cellRef(rowIndex + rowCount - 1, columnIndex + columnCount - 1)}`;
}

/** Split "'My Sheet'!A1:B4" into sheet name and local address. */
export function splitAddress(address: string): { sheet: string | null; local: string } {
  const bang = address.lastIndexOf("!");
  if (bang < 0) return { sheet: null, local: address };
  let sheet = address.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, local: address.slice(bang + 1) };
}

/** Quote a sheet name for use in an address when needed. */
export function qualify(sheet: string, local: string): string {
  const needsQuotes = !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet);
  const s = needsQuotes ? `'${sheet.replace(/'/g, "''")}'` : sheet;
  return `${s}!${local}`;
}

export interface Rect {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

export function rectContains(r: Rect, row: number, col: number): boolean {
  return row >= r.row && row < r.row + r.rows && col >= r.col && col < r.col + r.cols;
}

/**
 * Turn arbitrary text into a legal Excel table / defined name:
 * letters, digits, "_" and "."; must start with a letter or "_"; must not look
 * like a cell reference (A1 or R1C1 style); at most 255 characters.
 */
export function sanitizeExcelName(raw: string, fallback = "CDF_Table"): string {
  let name = raw.trim().replace(/[^A-Za-z0-9_.]+/g, "_").replace(/_+/g, "_");
  name = name.replace(/^_+(?=[A-Za-z])/, "");
  const trimmed = name.replace(/_+$/, "");
  if (trimmed) name = trimmed;
  if (!name) name = fallback;
  if (!/^[A-Za-z_]/.test(name)) name = `_${name}`;
  if (looksLikeReference(name)) name = `${name}_`;
  return name.slice(0, 255);
}

function looksLikeReference(name: string): boolean {
  return (
    /^[A-Za-z]{1,3}[0-9]+$/.test(name) || // A1, XFD1048576
    /^[RrCc]$/.test(name) ||
    /^[Rr][0-9]*[Cc][0-9]*$/.test(name) || // R1C1, RC
    /^[Rr][0-9]+$/.test(name) ||
    /^[Cc][0-9]+$/.test(name)
  );
}

/** Make `base` unique against a set of existing names (case-insensitive). */
export function uniqueName(base: string, existing: Iterable<string>): string {
  const taken = new Set(Array.from(existing, (n) => n.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) {
    const candidate = `${base}_${i}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}

/** Excel worksheet names: max 31 chars, none of : \ / ? * [ ], not blank. */
export function sanitizeSheetName(raw: string, fallback = "CDF Output"): string {
  const cleaned = raw.replace(/[:\\/?*[\]]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31).trim();
  return cleaned || fallback;
}

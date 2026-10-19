/**
 * Round to a fixed number of decimal places to strip binary floating-point
 * noise (e.g. 0.1 + 0.2 = 0.30000000000000004 -> 0.3) without changing any
 * value a user could have typed.
 */
export function cleanFloat(value: number, decimals = 12): number {
  if (!Number.isFinite(value)) return value;
  const rounded = Number(value.toFixed(decimals));
  return Object.is(rounded, -0) ? 0 : rounded;
}

/** Round to significant digits (scale-independent noise removal). */
export function cleanSig(value: number, digits = 14): number {
  if (!Number.isFinite(value) || value === 0) return value === 0 ? 0 : value;
  return Number(value.toPrecision(digits));
}

/** Sum with Neumaier compensation, which keeps long probability lists accurate. */
export function stableSum(values: readonly number[]): number {
  let sum = 0;
  let compensation = 0;
  for (const v of values) {
    const t = sum + v;
    if (Math.abs(sum) >= Math.abs(v)) {
      compensation += sum - t + v;
    } else {
      compensation += v - t + sum;
    }
    sum = t;
  }
  return sum + compensation;
}

export function approxEqual(a: number, b: number, tolerance: number): boolean {
  return Math.abs(a - b) <= tolerance;
}

/** Format a number compactly for messages (max 10 significant digits). */
export function fmt(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  return String(Number(value.toPrecision(10)));
}

/**
 * Largest "nice" number (multiple of a power of ten matched to the data
 * scale) that is strictly below `min`. Used as an automatic CDF starting
 * point when outcomes are negative.
 */
export function niceFloorBelow(min: number, max: number): number {
  const span = max - min;
  const scale = span > 0 ? span : Math.abs(min) > 0 ? Math.abs(min) : 1;
  const step = Math.pow(10, Math.floor(Math.log10(scale)));
  const pad = span > 0 ? span * 0.05 : scale * 0.1;
  let candidate = cleanFloat(Math.floor((min - pad) / step) * step);
  if (candidate >= min) candidate = cleanFloat(candidate - step);
  return candidate;
}

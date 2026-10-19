/**
 * CDF step-coordinate generation. Converts a discrete distribution into the
 * duplicated X/Y points that an XY scatter chart with straight lines draws as
 * a right-continuous step function: horizontal segments with vertical jumps
 * at every outcome.
 */
import { cleanFloat, fmt, niceFloorBelow } from "./numeric";
import { DEFAULT_TOLERANCE, Distribution, Point } from "./types";

export type StartMode = "auto" | "zero" | "min" | "custom";
export type EndMode = "max" | "none" | "custom";

export interface XRangeOptions {
  startMode: StartMode;
  customStart?: number;
  endMode: EndMode;
  customEnd?: number;
}

export interface StepOptions {
  /** X of the initial (xStart, 0) point. Must be <= the smallest outcome. */
  xStart: number;
  /** Optional X to extend the final horizontal segment at cumulative probability 1. */
  xEnd?: number | null;
  tolerance?: number;
}

/** Cumulative probability just after each outcome: F(x) = P(X <= x). */
export function cumulativeProbabilities(dist: Distribution, tolerance = DEFAULT_TOLERANCE): { x: number; F: number }[] {
  let cum = 0;
  const out = dist.outcomes.map((o) => {
    cum = cleanFloat(cum + o.p);
    return { x: o.x, F: cum };
  });
  const last = out[out.length - 1];
  if (last && Math.abs(last.F - 1) <= tolerance) {
    // Snap to exactly 1 and keep the sequence monotone after snapping.
    last.F = 1;
    for (let i = out.length - 2; i >= 0 && out[i].F > 1; i--) out[i].F = 1;
  }
  return out;
}

/**
 * Build step coordinates:
 *   (xStart, 0), then for each outcome x with probability p > 0:
 *   (x, F before the jump), (x, F after the jump),
 *   and optionally (xEnd, 1).
 * Consecutive identical points (e.g. when xStart equals the first outcome)
 * are dropped; zero-probability outcomes add no step.
 */
export function buildStepCoordinates(dist: Distribution, options: StepOptions): Point[] {
  const tolerance = options.tolerance ?? DEFAULT_TOLERANCE;
  const positive = dist.outcomes.filter((o) => o.p > 0);
  if (positive.length === 0) throw new Error("The distribution has no outcome with positive probability.");

  const { xStart } = options;
  if (!Number.isFinite(xStart)) throw new RangeError("The starting X value must be a finite number.");
  if (xStart > positive[0].x) {
    throw new RangeError(
      `The starting X value (${fmt(xStart)}) is greater than the smallest outcome (${fmt(positive[0].x)}). Choose a value at or below it.`,
    );
  }

  const cumulative = cumulativeProbabilities({ outcomes: positive }, tolerance);
  const final = cumulative[cumulative.length - 1].F;
  if (Math.abs(final - 1) > tolerance) {
    throw new RangeError(`The cumulative probability ends at ${fmt(final)} instead of 1.`);
  }

  const points: Point[] = [{ x: xStart, y: 0 }];
  let before = 0;
  for (const { x, F } of cumulative) {
    points.push({ x, y: before }, { x, y: F });
    before = F;
  }

  const lastX = positive[positive.length - 1].x;
  if (options.xEnd !== undefined && options.xEnd !== null) {
    if (!Number.isFinite(options.xEnd)) throw new RangeError("The ending X value must be a finite number.");
    if (options.xEnd < lastX) {
      throw new RangeError(
        `The ending X value (${fmt(options.xEnd)}) is less than the largest outcome (${fmt(lastX)}).`,
      );
    }
    points.push({ x: options.xEnd, y: 1 });
  }

  return points.filter((pt, i) => i === 0 || pt.x !== points[i - 1].x || pt.y !== points[i - 1].y);
}

/** Default start: 0 for non-negative data, otherwise a round number below the minimum. */
export function automaticStart(minX: number, maxX: number): number {
  return minX >= 0 ? 0 : niceFloorBelow(minX, maxX);
}

/**
 * Resolve the shared X range for one or more distributions drawn on the same
 * chart, so every CDF starts and ends at the same X.
 */
export function resolveXRange(dists: readonly Distribution[], options: XRangeOptions): { xStart: number; xEnd: number | null } {
  const xs = dists.flatMap((d) => d.outcomes.map((o) => o.x));
  const support = dists.flatMap((d) => d.outcomes.filter((o) => o.p > 0).map((o) => o.x));
  if (xs.length === 0 || support.length === 0) throw new Error("No distribution data to chart.");
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minSupport = Math.min(...support);
  const maxSupport = Math.max(...support);

  let xStart: number;
  switch (options.startMode) {
    case "zero":
      xStart = 0;
      break;
    case "min":
      xStart = minX;
      break;
    case "custom":
      if (options.customStart === undefined || !Number.isFinite(options.customStart)) {
        throw new RangeError("Enter a numeric starting X value.");
      }
      xStart = options.customStart;
      break;
    default:
      xStart = automaticStart(minX, maxX);
  }
  if (xStart > minSupport) {
    throw new RangeError(
      `The starting X value (${fmt(xStart)}) is greater than the smallest outcome (${fmt(minSupport)}). ` +
        `Use "Automatic" or a value at or below ${fmt(minSupport)}.`,
    );
  }

  let xEnd: number | null = null;
  if (options.endMode === "max") {
    xEnd = maxX;
  } else if (options.endMode === "custom") {
    if (options.customEnd === undefined || !Number.isFinite(options.customEnd)) {
      throw new RangeError("Enter a numeric ending X value.");
    }
    if (options.customEnd < maxSupport) {
      throw new RangeError(
        `The ending X value (${fmt(options.customEnd)}) is less than the largest outcome (${fmt(maxSupport)}).`,
      );
    }
    xEnd = options.customEnd;
  }
  return { xStart, xEnd };
}

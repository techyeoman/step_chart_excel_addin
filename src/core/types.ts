/** A single discrete outcome and its probability mass. */
export interface Outcome {
  x: number;
  p: number;
}

/**
 * A validated discrete distribution: outcomes are sorted ascending, unique,
 * every p is in [0, 1], and the probabilities sum to 1 within tolerance.
 */
export interface Distribution {
  outcomes: Outcome[];
}

/** One (x, y) coordinate of the step chart. */
export interface Point {
  x: number;
  y: number;
}

export type Severity = "error" | "warning" | "info";

export interface ValidationIssue {
  severity: Severity;
  message: string;
  /** A1-style cell reference (without sheet) when the issue concerns one cell. */
  cell?: string;
}

/** Default tolerance for "probabilities sum to 1". */
export const DEFAULT_TOLERANCE = 1e-6;

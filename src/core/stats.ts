/** Summary statistics of a discrete distribution. */
import { cumulativeProbabilities } from "./cdf";
import { cleanSig, stableSum } from "./numeric";
import { Distribution } from "./types";

export interface DistributionStats {
  outcomes: number;
  expectedValue: number;
  variance: number;
  standardDeviation: number;
  min: number;
  max: number;
  /** Smallest x with P(X <= x) >= 0.5. */
  median: number;
}

/**
 * EV = sum(x * p); Var = sum(p * (x - EV)^2) (population variance of the
 * discrete distribution). Probabilities are normalised by their sum so that
 * tiny rounding in the inputs (within the validation tolerance) cannot leak
 * into the results.
 */
export function computeStats(dist: Distribution): DistributionStats {
  const support = dist.outcomes.filter((o) => o.p > 0);
  if (support.length === 0) throw new Error("The distribution has no outcome with positive probability.");
  const total = stableSum(support.map((o) => o.p));
  const mean = stableSum(support.map((o) => (o.x * o.p) / total));
  const variance = Math.max(0, stableSum(support.map((o) => ((o.x - mean) ** 2 * o.p) / total)));
  const cdf = cumulativeProbabilities({ outcomes: support });
  const median = (cdf.find((c) => c.F >= 0.5 - 1e-12) ?? cdf[cdf.length - 1]).x;
  return {
    outcomes: support.length,
    expectedValue: cleanSig(mean),
    variance: cleanSig(variance),
    standardDeviation: cleanSig(Math.sqrt(variance)),
    min: support[0].x,
    max: support[support.length - 1].x,
    median,
  };
}

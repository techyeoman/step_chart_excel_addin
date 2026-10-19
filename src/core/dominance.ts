/**
 * Exact first- and second-order stochastic dominance tests for discrete
 * distributions.
 *
 * Both CDFs are right-continuous step functions that only change at outcome
 * values, so F_A - F_B is constant on every interval [z_k, z_{k+1}) of the
 * sorted union z_1 < ... < z_n of all outcomes, and zero below z_1. Checking
 * F at the union points therefore checks every x.
 *
 * The integrated CDF G(x) = integral from -inf to x of F(t) dt is piecewise
 * linear with breakpoints at the same z_k, so G_A - G_B attains its extremes
 * at those points (beyond z_n its slope is F_A - F_B = 1 - 1 = 0). Checking G
 * at the union points is therefore exact as well; no grid or sampling is used.
 *
 * "lower" preference (e.g. costs) is handled by testing -X, which turns
 * "lower is better" into the standard "higher is better" conditions.
 */
import { cleanFloat, fmt } from "./numeric";
import { DEFAULT_TOLERANCE, Distribution } from "./types";

export type Preference = "higher" | "lower";

/** Which distribution dominates: "A", "B", "neither", or "identical". */
export type Verdict = "A" | "B" | "neither" | "identical";

export interface ComparisonRow {
  x: number;
  /** CDF values P(X <= x). */
  FA: number;
  FB: number;
  /**
   * Integrated values used in the SSD test.
   * higher: integral_{-inf}^{x} F(t) dt;  lower: integral_{x}^{inf} (1 - F(t)) dt.
   */
  IA: number;
  IB: number;
}

export interface DominanceResult {
  preference: Preference;
  fsd: Verdict;
  ssd: Verdict;
  rows: ComparisonRow[];
  /** X values where the sign of F_A - F_B changes (the CDFs cross). */
  cdfCrossings: number[];
  /** X values where the sign of I_A - I_B changes. */
  integralCrossings: number[];
}

export interface DominanceOptions {
  preference?: Preference;
  /** Tolerance for comparing CDF values; scaled by the outcome span for integrals. */
  tolerance?: number;
}

interface Evaluated {
  xs: number[];
  FA: number[];
  FB: number[];
  GA: number[];
  GB: number[];
}

function cdfAt(dist: Distribution, xs: number[]): number[] {
  const out: number[] = [];
  let i = 0;
  let cum = 0;
  for (const x of xs) {
    while (i < dist.outcomes.length && dist.outcomes[i].x <= x) {
      cum = cleanFloat(cum + dist.outcomes[i].p);
      i++;
    }
    out.push(cum);
  }
  return out;
}

function integrate(xs: number[], F: number[]): number[] {
  const G = [0];
  for (let k = 1; k < xs.length; k++) G.push(G[k - 1] + F[k - 1] * (xs[k] - xs[k - 1]));
  return G;
}

function sorted(dist: Distribution): Distribution {
  return { outcomes: [...dist.outcomes].sort((a, b) => a.x - b.x) };
}

function negate(dist: Distribution): Distribution {
  return sorted({ outcomes: dist.outcomes.map((o) => ({ x: o.x === 0 ? 0 : -o.x, p: o.p })) });
}

function evaluate(a: Distribution, b: Distribution): Evaluated {
  const xs = Array.from(new Set([...a.outcomes, ...b.outcomes].map((o) => o.x))).sort((p, q) => p - q);
  const FA = cdfAt(a, xs);
  const FB = cdfAt(b, xs);
  return { xs, FA, FB, GA: integrate(xs, FA), GB: integrate(xs, FB) };
}

/** A dominates when its curve is <= B's everywhere and strictly below somewhere. */
export function verdictFromDifferences(diffs: readonly number[], tolerance: number): Verdict {
  let anyPositive = false;
  let anyNegative = false;
  for (const d of diffs) {
    if (d > tolerance) anyPositive = true;
    else if (d < -tolerance) anyNegative = true;
  }
  if (!anyPositive && !anyNegative) return "identical";
  if (!anyPositive) return "A";
  if (!anyNegative) return "B";
  return "neither";
}

function signChanges(xs: readonly number[], diffs: readonly number[], tolerance: number): number[] {
  const out: number[] = [];
  let lastSign = 0;
  diffs.forEach((d, k) => {
    const sign = d > tolerance ? 1 : d < -tolerance ? -1 : 0;
    if (sign !== 0) {
      if (lastSign !== 0 && sign !== lastSign) out.push(xs[k]);
      lastSign = sign;
    }
  });
  return out;
}

export function compareDistributions(a: Distribution, b: Distribution, options: DominanceOptions = {}): DominanceResult {
  const preference = options.preference ?? "higher";
  const tolerance = options.tolerance ?? DEFAULT_TOLERANCE;
  if (a.outcomes.length === 0 || b.outcomes.length === 0) throw new Error("Both distributions need at least one outcome.");

  const A = sorted(a);
  const B = sorted(b);
  const original = evaluate(A, B);
  const test = preference === "higher" ? original : evaluate(negate(A), negate(B));
  const span = original.xs[original.xs.length - 1] - original.xs[0];
  const integralTolerance = tolerance * Math.max(1, Math.abs(span));

  const fsdDiffs = test.FA.map((f, k) => f - test.FB[k]);
  const ssdDiffs = test.GA.map((g, k) => g - test.GB[k]);
  let fsd = verdictFromDifferences(fsdDiffs, tolerance);
  let ssd = verdictFromDifferences(ssdDiffs, integralTolerance);
  // FSD implies SSD; keep the two verdicts consistent under rounding.
  if (fsd === "A" || fsd === "B") ssd = fsd;
  if (fsd === "identical") ssd = "identical";

  // Integrated values in original-x order for display.
  const IA = preference === "higher" ? original.GA : test.GA.slice().reverse();
  const IB = preference === "higher" ? original.GB : test.GB.slice().reverse();
  const rows: ComparisonRow[] = original.xs.map((x, k) => ({
    x,
    FA: original.FA[k],
    FB: original.FB[k],
    IA: cleanFloat(IA[k], 10),
    IB: cleanFloat(IB[k], 10),
  }));

  const cdfDiffs = rows.map((r) => r.FA - r.FB);
  const integralDiffs = rows.map((r) => r.IA - r.IB);
  return {
    preference,
    fsd,
    ssd,
    rows,
    cdfCrossings: signChanges(original.xs, cdfDiffs, tolerance),
    integralCrossings: signChanges(original.xs, integralDiffs, integralTolerance),
  };
}

export interface DominanceExplanation {
  fsdHeadline: string;
  fsdDetail: string;
  ssdHeadline: string;
  ssdDetail: string;
  notes: string[];
}

/** Plain-language interpretation of a DominanceResult. */
export function explainDominance(
  result: DominanceResult,
  nameA: string,
  nameB: string,
  means?: { a: number; b: number },
): DominanceExplanation {
  const higher = result.preference === "higher";
  const name = (v: "A" | "B") => (v === "A" ? nameA : nameB);
  const other = (v: "A" | "B") => (v === "A" ? nameB : nameA);
  const ineq = higher ? "≤" : "≥";
  const prefers = higher ? "prefers more to less (any increasing utility function)" : "prefers less to more (any decreasing utility function)";
  const riskAverse = higher
    ? "risk-averse decision maker who prefers more to less (increasing, concave utility)"
    : "risk-averse decision maker who prefers less to more (decreasing, concave utility)";
  const integral = higher ? "∫₋∞ˣ F(t) dt" : "∫ₓ^∞ (1 − F(t)) dt";
  const crossingText = (xs: number[]) =>
    xs.length ? ` They cross near x = ${xs.slice(0, 5).map(fmt).join(", ")}${xs.length > 5 ? ", …" : ""}.` : "";

  let fsdHeadline: string;
  let fsdDetail: string;
  switch (result.fsd) {
    case "identical":
      fsdHeadline = "The distributions are identical.";
      fsdDetail = "Their CDFs coincide at every outcome, so no decision maker can prefer one over the other.";
      break;
    case "neither":
      fsdHeadline = "Neither distribution first-order dominates the other.";
      fsdDetail =
        `The CDFs cross: each is strictly ${higher ? "lower" : "higher"} than the other somewhere.` +
        crossingText(result.cdfCrossings) +
        ` Some decision makers who ${higher ? "prefer more" : "prefer less"} will choose ${nameA}, others ${nameB}.`;
      break;
    default:
      fsdHeadline = `${name(result.fsd)} first-order stochastically dominates ${other(result.fsd)}.`;
      fsdDetail =
        `F_${name(result.fsd)}(x) ${ineq} F_${other(result.fsd)}(x) at every x, with strict inequality for at least one x. ` +
        `Every decision maker who ${prefers} weakly prefers ${name(result.fsd)}.`;
  }

  let ssdHeadline: string;
  let ssdDetail: string;
  switch (result.ssd) {
    case "identical":
      ssdHeadline = "Identical distributions (no second-order preference).";
      ssdDetail = "The integrated CDFs coincide.";
      break;
    case "neither":
      ssdHeadline = "Neither distribution second-order dominates the other.";
      ssdDetail =
        `The integrated CDFs ${integral} cross.` +
        crossingText(result.integralCrossings) +
        ` Different risk-averse decision makers can rank ${nameA} and ${nameB} differently.`;
      break;
    default:
      ssdHeadline = `${name(result.ssd)} second-order stochastically dominates ${other(result.ssd)}.`;
      ssdDetail =
        `${integral} for ${name(result.ssd)} is ${ineq} that of ${other(result.ssd)} at every x, strictly somewhere. ` +
        `Every ${riskAverse} weakly prefers ${name(result.ssd)}.` +
        (result.fsd === result.ssd ? " (Implied by first-order dominance.)" : "");
  }

  const notes: string[] = [
    "Dominance is decided from the full CDFs, evaluated exactly at every outcome of either distribution. " +
      "Expected value and variance are shown for information only and are not used to classify dominance.",
  ];
  if (means && result.ssd === "neither") {
    const better = higher ? (means.a > means.b ? nameA : nameB) : means.a < means.b ? nameA : nameB;
    if (means.a !== means.b) {
      notes.push(
        `${better} has the more favourable expected value, but that alone does not make it preferred by every risk-averse decision maker.`,
      );
    }
  }
  if (result.fsd === "neither" && (result.ssd === "A" || result.ssd === "B")) {
    notes.push(
      `Because the CDFs cross, the ranking holds for risk-averse decision makers only; a risk-seeking decision maker may prefer ${other(result.ssd)}.`,
    );
  }
  return { fsdHeadline, fsdDetail, ssdHeadline, ssdDetail, notes };
}

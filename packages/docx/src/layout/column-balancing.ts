export interface ColumnBalanceProbe {
  readonly fits: boolean;
  readonly requiredTargetPt: number;
  readonly thresholdsPt: readonly number[];
}

/**
 * Explore the finite set of measured line, row, and block boundaries exposed by
 * the real paginator. Feasibility is not numerically monotone: lowering a cap can
 * produce a more even legal partition after an immediately smaller value fails.
 * Closing over exact observed thresholds avoids both that false monotonicity and
 * an empirical epsilon.
 */
export function minimumColumnBalanceTarget(
  maximumTargetPt: number,
  probe: (targetPt: number) => ColumnBalanceProbe,
): number {
  if (!Number.isFinite(maximumTargetPt) || maximumTargetPt < 0) {
    throw new RangeError('Column balance maximum must be finite and non-negative');
  }
  const pending = [maximumTargetPt];
  const observed = new Set<number>();
  let minimum: number | null = null;
  while (pending.length > 0) {
    const targetPt = pending.pop()!;
    if (observed.has(targetPt)) continue;
    observed.add(targetPt);
    const result = probe(targetPt);
    if (!Number.isFinite(result.requiredTargetPt)
      || result.requiredTargetPt < 0
      || result.thresholdsPt.some((threshold) =>
        !Number.isFinite(threshold) || threshold < 0 || threshold > maximumTargetPt)) {
      throw new Error('Column balance probe returned an invalid occupied boundary');
    }
    if (result.fits) {
      if (result.requiredTargetPt > targetPt) {
        throw new Error('Column balance probe returned an invalid occupied boundary');
      }
      minimum = minimum === null ? targetPt : Math.min(minimum, targetPt);
    }
    for (const threshold of [result.requiredTargetPt, ...result.thresholdsPt]) {
      if (!observed.has(threshold)) pending.push(threshold);
    }
  }
  if (minimum === null) {
    throw new Error('The physical column extent must admit its unbalanced content');
  }
  return minimum;
}

// Excel-style "nice" value-axis scaling. Pure math (no canvas), extracted so it
// can be unit-tested and reused independently of the chart renderer.

/** A round major-unit step that yields roughly `targetSteps` gridlines across
 *  `range` (1 / 2 / 5 × 10ⁿ — Excel's default ladder). */
export function niceStep(range: number, targetSteps = 5): number {
  if (range === 0) return 1;
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
}

/** Excel / PowerPoint automatic value-axis maximum. Microsoft's documented
 *  algorithm (per Peltier Tech) is "the first major unit above
 *  `Ymax + (Ymax − Ymin)/20`": ~5% of the data range is added as headroom so the
 *  tallest series sits just below the top gridline rather than flush against it,
 *  then the result is rounded up to the next major unit. `dataMin` is the axis
 *  minimum (0 for bar/column charts; the data minimum otherwise).
 *
 *  The major unit itself is Excel-proprietary (it varies with plot size, tick
 *  font, etc. and is not documented), so we approximate it with `niceStep`; the
 *  computed max can therefore differ from PowerPoint by one major unit on some
 *  charts. */
export function niceAxisMax(dataMax: number, step: number, dataMin = 0): number {
  if (dataMax <= 0) return step;
  const withHeadroom = dataMax + (dataMax - dataMin) / 20;
  return Math.ceil(withHeadroom / step) * step;
}

/** Axis minimum for data that dips below zero: the largest major-unit multiple
 *  <= dataMin, dropping one extra step when the data sits exactly on a
 *  gridline so the lowest point isn't flush against the axis. Non-negative data
 *  anchors the axis at 0. */
export function niceAxisMin(dataMin: number, step: number): number {
  if (dataMin >= 0) return 0;
  const ax = Math.floor(dataMin / step) * step;
  return Math.abs(ax - dataMin) < step * 1e-9 ? ax - step : ax;
}

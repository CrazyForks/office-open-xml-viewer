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

/** Excel-style auto value-axis bounds + major unit. ONE `niceStep` (of the data
 *  range) drives the rounded min, max AND the gridline step, so they can never
 *  desync. Explicit `<c:valAx><c:scaling><c:min/max>` wins. The auto major unit
 *  is Excel-proprietary (not in ECMA-376); niceStep approximates it. */
/** Target gridline spacing in POINTS. Excel's auto major unit is not a fixed
 *  gridline count — it targets a roughly constant on-screen spacing, so a long
 *  axis (e.g. a horizontal bar chart's wide value axis) gets MORE, finer
 *  gridlines than a short one of the same data range. Empirically ~one major
 *  gridline per this many points reproduces PowerPoint across sample-14's
 *  column / area / horizontal-bar / secondary-axis charts. This is a runtime
 *  Excel behavior (not in ECMA-376); the constant is the one tunable. */
const GRIDLINE_SPACING_PT = 40;

/** Pick the `niceStep` target-gridline count for an axis of `axisLenPt` points.
 *  Falls back to 5 (the legacy fixed target) when the length is unknown. */
function targetStepsForAxis(axisLenPt?: number): number {
  if (axisLenPt == null || !isFinite(axisLenPt) || axisLenPt <= 0) return 5;
  return Math.min(15, Math.max(3, Math.round(axisLenPt / GRIDLINE_SPACING_PT)));
}

export function valueAxisScale(
  dataMin: number, dataMax: number,
  explicitMin?: number | null, explicitMax?: number | null,
  axisLenPt?: number,
): { min: number; max: number; step: number } {
  const step = niceStep(dataMax - dataMin, targetStepsForAxis(axisLenPt));
  const min = explicitMin ?? niceAxisMin(dataMin, step);
  const max = explicitMax ?? niceAxisMax(dataMax, step, min);
  return { min, max, step };
}

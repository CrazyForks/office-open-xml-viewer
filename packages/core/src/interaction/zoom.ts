/** Ctrl/⌘ + wheel (and trackpad pinch) zoom sensitivity. `deltaY` is multiplied
 *  by this before `exp()`. Purely an interaction-feel constant (no ECMA-376
 *  bearing); lower = gentler. */
const ZOOM_WHEEL_SENSITIVITY = 0.01;

/**
 * New scale for one wheel/pinch zoom step. The step is *exponential* in
 * `deltaY` rather than a fixed increment, which fixes two problems with a
 * sign-only `scale ± 0.1`:
 *
 *  - A trackpad pinch arrives as a high-frequency stream of small-`deltaY`
 *    wheel events; a fixed per-event increment compounds across dozens of
 *    events per gesture and zooms wildly. Because `exp(-k·a)·exp(-k·b) =
 *    exp(-k·(a+b))`, the total zoom here depends only on the summed `deltaY`
 *    of the gesture, not on how many events the OS chops it into — so a pinch
 *    and a mouse wheel covering the same distance zoom by the same amount.
 *  - It is multiplicative, so a step feels proportional at every zoom level
 *    (the old additive `+0.1` was huge at 20% and tiny at 400%), and exactly
 *    symmetric: zooming in then out by the same delta returns to the start.
 *
 * Negative `deltaY` (scroll up / pinch out) zooms in. The result is unclamped
 * and unsnapped; the caller clamps to its `[zoomMin, zoomMax]` range (and, for
 * XlsxViewer, snaps to whole percent). Shared by XlsxViewer and the two
 * continuous-scroll viewers (design §5.2).
 */
export function zoomStepScale(currentScale: number, deltaY: number): number {
  return currentScale * Math.exp(-deltaY * ZOOM_WHEEL_SENSITIVITY);
}

/**
 * Turn a matched run-slice (from {@link findMatches}) into the horizontal extent
 * of the highlight box within its run, using a caller-supplied text-measurer.
 *
 * A run's glyphs are laid left-to-right from the run's own origin. A slice
 * `[start, end)` of the run text therefore starts at the advance width of
 * `runText[0..start)` and ends at the advance width of `runText[0..end)`. The
 * caller supplies `measure(s)` — a `CanvasRenderingContext2D.measureText(s).width`
 * closure already primed with the run's font (`ctx.font = run.font`) — so this
 * stays a pure arithmetic wrapper with no canvas/DOM dependency of its own and is
 * shared by the docx and pptx highlight overlays (xlsx measures inside the cell
 * rect the same way). The vertical extent (top / height) is the run's line box,
 * owned by each renderer's overlay, so it is not computed here.
 *
 * Measuring the two prefixes (rather than measuring the slice text alone and
 * summing) is deliberate: it accounts for kerning between the run's leading
 * glyphs and keeps the highlight edges flush with where the canvas actually drew
 * those characters, exactly as the selection overlay relies on `measureText`
 * matching `fillText`.
 *
 * @param runText   the full text of the run the slice belongs to.
 * @param start     slice start offset within `runText` (inclusive).
 * @param end       slice end offset within `runText` (exclusive).
 * @param measure   advance width in px of a substring, in the run's font.
 * @returns `x` (left offset from the run origin, px) and `width` (px).
 */
export function sliceHorizontalExtent(
  runText: string,
  start: number,
  end: number,
  measure: (s: string) => number,
): { x: number; width: number } {
  const x = start <= 0 ? 0 : measure(runText.slice(0, start));
  const endX = end >= runText.length ? measure(runText) : measure(runText.slice(0, end));
  return { x, width: Math.max(0, endX - x) };
}

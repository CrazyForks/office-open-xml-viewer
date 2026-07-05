/**
 * IX2 active-match cursor — the pure arithmetic behind `findNext` / `findPrev`.
 *
 * A viewer keeps one "active" match highlighted distinctly and cycles it with
 * wrap-around, exactly like a browser's find bar: `findNext` advances (last →
 * first), `findPrev` steps back (first → last). The convention for the active
 * index is `-1` = "no active match yet" (nothing highlighted). From that state
 * `findNext` lands on the first match and `findPrev` on the last, matching what
 * a user expects the first Enter / Shift+Enter to do.
 *
 * This is the only piece of find-cursor logic shared by all three viewers, and
 * it is pure index math (no DOM, no geometry), so it lives in core rather than
 * being re-derived — with off-by-one wrap bugs — in each package.
 */

/**
 * The next active-match index after `findNext`, with wrap-around.
 *
 * @param active current active index, or `-1` for "none yet".
 * @param count  total number of matches.
 * @returns the new active index in `[0, count)`, or `-1` when `count === 0`.
 */
export function nextActive(active: number, count: number): number {
  if (count <= 0) return -1;
  if (active < 0) return 0; // no active match yet → first
  return (active + 1) % count;
}

/**
 * The previous active-match index after `findPrev`, with wrap-around.
 *
 * @param active current active index, or `-1` for "none yet".
 * @param count  total number of matches.
 * @returns the new active index in `[0, count)`, or `-1` when `count === 0`.
 */
export function prevActive(active: number, count: number): number {
  if (count <= 0) return -1;
  if (active < 0) return count - 1; // no active match yet → last
  return (active - 1 + count) % count;
}

/**
 * Normalize an active index against the current match count: an in-range index
 * passes through, anything else (negative, or ≥ count because the match set
 * shrank on a new query) collapses to `-1` (no active match). Used when a
 * viewer recomputes matches and must decide whether the old active index is
 * still meaningful.
 */
export function clampActive(active: number, count: number): number {
  if (count <= 0) return -1;
  if (active < 0 || active >= count) return -1;
  return active;
}

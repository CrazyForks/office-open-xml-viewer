/**
 * Next visible (non-hidden) absolute index in `dir` (+1 forward, -1 back),
 * searching strictly AFTER `from`. Shared by viewers whose sequential-navigation
 * ("skip") mode jumps over hidden items (pptx hidden slides, xlsx hidden sheets).
 * Returns the found index, or `from` itself when no other visible item exists in
 * that direction (caller then stays put).
 */
export function nextVisibleIndex(
  from: number,
  dir: 1 | -1,
  isHidden: (i: number) => boolean,
  count: number,
): number {
  for (let i = from + dir; i >= 0 && i < count; i += dir) {
    if (!isHidden(i)) return i;
  }
  return from;
}

/**
 * Resolve which item to show on initial load or when entering "skip" mode:
 * keep `current` if visible; else the nearest visible item AFTER it; else the
 * nearest visible item BEFORE it; else `current` (every item hidden — skip
 * degenerates, so show it anyway). All indices are absolute.
 */
export function resolveVisibleIndex(
  current: number,
  isHidden: (i: number) => boolean,
  count: number,
): number {
  if (count === 0 || !isHidden(current)) return current;
  const fwd = nextVisibleIndex(current, 1, isHidden, count);
  if (fwd !== current) return fwd;
  return nextVisibleIndex(current, -1, isHidden, count);
}

/**
 * Count how many of `count` items are visible (not hidden), per the same
 * `isHidden` callback the skip-navigation helpers use. Shared by viewers'
 * `visible*Count` getters (pptx visibleSlideCount, xlsx visibleSheetCount).
 */
export function countVisible(isHidden: (i: number) => boolean, count: number): number {
  let n = 0;
  for (let i = 0; i < count; i++) if (!isHidden(i)) n++;
  return n;
}

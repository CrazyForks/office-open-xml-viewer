import type { Slide } from './types';

/**
 * Pure core of {@link PptxPresentation.isHidden}: whether the slide at
 * `slideIndex` (0-based, absolute) is marked hidden. Like `selectNotes`, the
 * index is NOT clamped — out-of-range / non-integer ⇒ `false` ("no slide here,
 * so not hidden") rather than the nearest slide's flag.
 */
export function selectHidden(slides: readonly Slide[], slideIndex: number): boolean {
  if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= slides.length) {
    return false;
  }
  return slides[slideIndex].hidden ?? false;
}

/**
 * Next visible (non-hidden) absolute slide index in `dir` (+1 forward, -1 back),
 * searching strictly AFTER `from`. Used by {@link PptxViewer}'s `'skip'` mode for
 * sequential navigation. Returns the found index, or `from` itself when no other
 * visible slide exists in that direction (caller then stays put).
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
 * Resolve which slide to show on initial load or when entering `'skip'` mode:
 * keep `current` if visible; else the nearest visible slide AFTER it; else the
 * nearest visible slide BEFORE it; else `current` (every slide hidden — skip
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

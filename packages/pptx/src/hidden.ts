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

// Generic index-navigation helpers live in core (shared with xlsx hidden sheets).
export { nextVisibleIndex, resolveVisibleIndex, countVisible } from '@silurus/ooxml-core';

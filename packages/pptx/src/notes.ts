import type { Slide } from './types';

/**
 * Pure core of {@link PptxPresentation.getNotes}: return the speaker-notes text
 * for the slide at `slideIndex` (0-based), or `null` when the index is out of
 * range / not an integer, or when the slide has no notes part.
 *
 * Unlike navigation methods (`goToSlide`) the index is *not* clamped — an
 * out-of-range request is a "no notes here" answer, returned as `null` rather
 * than the notes of the nearest slide, which would be misleading for a tool
 * reading notes by index.
 */
export function selectNotes(slides: readonly Slide[], slideIndex: number): string | null {
  if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= slides.length) {
    return null;
  }
  return slides[slideIndex].notes ?? null;
}

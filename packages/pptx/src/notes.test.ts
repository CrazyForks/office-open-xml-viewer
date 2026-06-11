import { describe, it, expect } from 'vitest';
import { selectNotes } from './notes.js';
import type { Slide } from './types.js';
import type { PptxPresentation } from './presentation.js';
import type { PptxViewer } from './viewer.js';

/**
 * Compile-time API-surface assertions (erased at runtime, enforced by
 * `pnpm typecheck`): both the headless engine and the viewer expose
 * `getNotes(slideIndex: number): string | null`.
 */
type Expect<T extends true> = T;
type IsGetNotes<T> = T extends (slideIndex: number) => string | null ? true : false;
type _PresentationHasGetNotes = Expect<IsGetNotes<PptxPresentation['getNotes']>>;
type _ViewerHasGetNotes = Expect<IsGetNotes<PptxViewer['getNotes']>>;

/**
 * Build a minimal slide list. Only the fields `selectNotes` reads matter; the
 * rest are filled with inert defaults so the value type-checks as a `Slide`.
 */
function slidesWithNotes(notes: Array<string | undefined>): Slide[] {
  return notes.map((n, i) => ({
    index: i,
    slideNumber: i + 1,
    background: null,
    elements: [],
    ...(n === undefined ? {} : { notes: n }),
  }));
}

describe('selectNotes (PptxPresentation.getNotes core)', () => {
  it('returns the notes string for an in-range slide', () => {
    const slides = slidesWithNotes(['first', 'second', 'third']);
    expect(selectNotes(slides, 0)).toBe('first');
    expect(selectNotes(slides, 2)).toBe('third');
  });

  it('returns null when the slide has no notes part', () => {
    const slides = slidesWithNotes(['has notes', undefined]);
    expect(selectNotes(slides, 1)).toBeNull();
  });

  it('returns null for an out-of-range index (no clamping, matches null contract)', () => {
    const slides = slidesWithNotes(['only']);
    expect(selectNotes(slides, -1)).toBeNull();
    expect(selectNotes(slides, 1)).toBeNull();
    expect(selectNotes(slides, 99)).toBeNull();
  });

  it('returns null for a non-integer index', () => {
    const slides = slidesWithNotes(['a', 'b']);
    expect(selectNotes(slides, 1.5)).toBeNull();
  });

  it('returns null on an empty deck', () => {
    expect(selectNotes([], 0)).toBeNull();
  });
});

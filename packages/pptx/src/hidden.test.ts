import { describe, it, expect } from 'vitest';
import { selectHidden } from './hidden.js';
import type { Slide } from './types.js';

/** Build a slide list where `hiddenFlags[i]` toggles `Slide.hidden`. */
function slides(hiddenFlags: boolean[]): Slide[] {
  return hiddenFlags.map((h, i) => ({
    index: i,
    slideNumber: i + 1,
    background: null,
    elements: [],
    ...(h ? { hidden: true } : {}),
  }));
}

describe('selectHidden (PptxPresentation.isHidden core)', () => {
  it('reports the hidden flag for an in-range slide', () => {
    const s = slides([false, true, false]);
    expect(selectHidden(s, 0)).toBe(false);
    expect(selectHidden(s, 1)).toBe(true);
  });
  it('treats absent hidden as false', () => {
    expect(selectHidden(slides([false]), 0)).toBe(false);
  });
  it('returns false for out-of-range / non-integer (non-clamped, like getNotes)', () => {
    const s = slides([true]);
    expect(selectHidden(s, -1)).toBe(false);
    expect(selectHidden(s, 1)).toBe(false);
    expect(selectHidden(s, 0.5)).toBe(false);
    expect(selectHidden([], 0)).toBe(false);
  });
});

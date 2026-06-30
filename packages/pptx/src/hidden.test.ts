import { describe, it, expect } from 'vitest';
import { selectHidden, nextVisibleIndex, resolveVisibleIndex } from './hidden.js';
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

describe('nextVisibleIndex (skip-mode sequential nav)', () => {
  // hidden at 1 and 2; visible 0,3,4
  const isHidden = (i: number) => i === 1 || i === 2;
  const count = 5;
  it('skips hidden slides going forward', () => {
    expect(nextVisibleIndex(0, 1, isHidden, count)).toBe(3);
  });
  it('skips hidden slides going backward', () => {
    expect(nextVisibleIndex(3, -1, isHidden, count)).toBe(0);
  });
  it('stays put when no visible slide exists in that direction', () => {
    expect(nextVisibleIndex(4, 1, isHidden, count)).toBe(4); // nothing after 4
    expect(nextVisibleIndex(0, -1, isHidden, count)).toBe(0); // nothing before 0
  });
});

describe('resolveVisibleIndex (initial load / entering skip mode)', () => {
  it('keeps the current slide when it is visible', () => {
    expect(resolveVisibleIndex(3, (i) => i === 0, 5)).toBe(3);
  });
  it('advances to the next visible slide when current is hidden', () => {
    expect(resolveVisibleIndex(0, (i) => i === 0 || i === 1, 5)).toBe(2);
  });
  it('falls back to the previous visible slide when none follow', () => {
    // visible: 0,1 ; hidden: 2,3,4 ; current 4 → nearest visible before is 1
    expect(resolveVisibleIndex(4, (i) => i >= 2, 5)).toBe(1);
  });
  it('returns current when every slide is hidden (skip degenerates)', () => {
    expect(resolveVisibleIndex(2, () => true, 5)).toBe(2);
  });
  it('returns current on an empty deck', () => {
    expect(resolveVisibleIndex(0, () => true, 0)).toBe(0);
  });
});

import { describe, it, expect } from 'vitest';
import { nextActive, prevActive, clampActive } from './find-cursor.js';

/**
 * IX2 active-match cursor. `findNext` / `findPrev` cycle a viewer's "active"
 * match with wrap-around (like a browser's find bar). The arithmetic is pure —
 * given the current active index and the total match count it returns the next /
 * previous index — so it lives in core and is shared by all three viewers.
 */
describe('nextActive', () => {
  it('advances to the next match', () => {
    expect(nextActive(0, 3)).toBe(1);
    expect(nextActive(1, 3)).toBe(2);
  });

  it('wraps from the last match back to the first', () => {
    expect(nextActive(2, 3)).toBe(0);
  });

  it('lands on the first match when there is no active match yet (-1)', () => {
    expect(nextActive(-1, 3)).toBe(0);
  });

  it('returns -1 when there are no matches', () => {
    expect(nextActive(-1, 0)).toBe(-1);
    expect(nextActive(0, 0)).toBe(-1);
  });

  it('stays on the only match in a one-match set', () => {
    expect(nextActive(0, 1)).toBe(0);
  });
});

describe('prevActive', () => {
  it('steps to the previous match', () => {
    expect(prevActive(2, 3)).toBe(1);
    expect(prevActive(1, 3)).toBe(0);
  });

  it('wraps from the first match to the last', () => {
    expect(prevActive(0, 3)).toBe(2);
  });

  it('lands on the last match when there is no active match yet (-1)', () => {
    expect(prevActive(-1, 3)).toBe(2);
  });

  it('returns -1 when there are no matches', () => {
    expect(prevActive(-1, 0)).toBe(-1);
    expect(prevActive(0, 0)).toBe(-1);
  });

  it('stays on the only match in a one-match set', () => {
    expect(prevActive(0, 1)).toBe(0);
  });
});

describe('clampActive', () => {
  it('keeps an in-range index unchanged', () => {
    expect(clampActive(1, 3)).toBe(1);
  });

  it('returns -1 for an out-of-range index (matches shrank)', () => {
    expect(clampActive(5, 3)).toBe(-1);
    expect(clampActive(3, 3)).toBe(-1);
  });

  it('returns -1 when there are no matches', () => {
    expect(clampActive(0, 0)).toBe(-1);
  });

  it('normalizes any negative index to -1 (no active)', () => {
    expect(clampActive(-2, 3)).toBe(-1);
  });
});

import { describe, it, expect } from 'vitest';
import { nextVisibleIndex, resolveVisibleIndex, countVisible } from './visible-index.js';

describe('nextVisibleIndex (skip-mode sequential nav)', () => {
  const isHidden = (i: number) => i === 1 || i === 2; // visible: 0,3,4
  const count = 5;
  it('skips hidden items going forward', () => {
    expect(nextVisibleIndex(0, 1, isHidden, count)).toBe(3);
  });
  it('skips hidden items going backward', () => {
    expect(nextVisibleIndex(3, -1, isHidden, count)).toBe(0);
  });
  it('stays put when no visible item exists in that direction', () => {
    expect(nextVisibleIndex(4, 1, isHidden, count)).toBe(4);
    expect(nextVisibleIndex(0, -1, isHidden, count)).toBe(0);
  });
});

describe('resolveVisibleIndex (initial load / entering skip mode)', () => {
  it('keeps the current item when it is visible', () => {
    expect(resolveVisibleIndex(3, (i) => i === 0, 5)).toBe(3);
  });
  it('advances to the next visible item when current is hidden', () => {
    expect(resolveVisibleIndex(0, (i) => i === 0 || i === 1, 5)).toBe(2);
  });
  it('falls back to the previous visible item when none follow', () => {
    expect(resolveVisibleIndex(4, (i) => i >= 2, 5)).toBe(1);
  });
  it('returns current when every item is hidden (skip degenerates)', () => {
    expect(resolveVisibleIndex(2, () => true, 5)).toBe(2);
  });
  it('returns current on an empty list', () => {
    expect(resolveVisibleIndex(0, () => true, 0)).toBe(0);
  });
});

describe('countVisible', () => {
  it('counts the non-hidden items', () => {
    expect(countVisible((i) => i === 1 || i === 2, 5)).toBe(3);
    expect(countVisible(() => false, 4)).toBe(4);
    expect(countVisible(() => true, 4)).toBe(0);
    expect(countVisible(() => false, 0)).toBe(0);
  });
});

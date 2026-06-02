import { describe, it, expect } from 'vitest';
import { niceStep, niceAxisMax, niceAxisMin } from './axis-scale.js';

describe('niceStep', () => {
  it('picks 1/2/5 × 10ⁿ for ~5 gridlines', () => {
    expect(niceStep(100)).toBe(20);  // raw 20 → 2×10
    expect(niceStep(50)).toBe(10);   // raw 10 → 1×10
    expect(niceStep(7)).toBe(1);     // raw 1.4 → 1×1
    expect(niceStep(40)).toBe(10);   // raw 8 → 1×10 (8 ≥ 7.5 → 10)
  });
  it('zero range falls back to 1', () => {
    expect(niceStep(0)).toBe(1);
  });
});

describe('niceAxisMax (Excel headroom: first major unit above Ymax + range/20)', () => {
  it('rounds up past the ~5% headroom to the next major unit', () => {
    expect(niceAxisMax(41, 10)).toBe(50);        // 41 + 2.05 = 43.05 → 50
    expect(niceAxisMax(9715, 2000)).toBe(12000); // 9715 + 485.75 = 10200.75 → 12000
  });
  it('adds headroom even when data sits on a gridline (not flush against the top)', () => {
    expect(niceAxisMax(40, 10)).toBe(50);   // 40 + 2 = 42 → 50
    expect(niceAxisMax(100, 20)).toBe(120); // 100 + 5 = 105 → 120
  });
  it('uses dataMin for the range', () => {
    // range 100-(-100)=200, headroom 10 → 110 → step 50 → 150
    expect(niceAxisMax(100, 50, -100)).toBe(150);
  });
  it('non-positive max returns one step', () => {
    expect(niceAxisMax(0, 10)).toBe(10);
    expect(niceAxisMax(-5, 10)).toBe(10);
  });
});

describe('niceAxisMin', () => {
  it('non-negative data anchors at 0', () => {
    expect(niceAxisMin(15, 10)).toBe(0);
    expect(niceAxisMin(0, 10)).toBe(0);
  });
  it('negative data floors to a major-unit multiple', () => {
    expect(niceAxisMin(-15, 10)).toBe(-20);
  });
  it('data exactly on a gridline drops one extra step', () => {
    expect(niceAxisMin(-20, 10)).toBe(-30);
  });
});

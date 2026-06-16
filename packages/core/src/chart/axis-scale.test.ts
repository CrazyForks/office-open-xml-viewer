import { describe, it, expect } from 'vitest';
import { niceStep, niceAxisMax, niceAxisMin, valueAxisScale } from './axis-scale.js';

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

describe('valueAxisScale (one niceStep drives min, max and gridline step)', () => {
  it('positive data anchored at 0 (bar/area/radar style)', () => {
    // step = niceStep(41-0) = niceStep(41) = 10; min = 0; max = niceAxisMax(41,10,0) = 50
    expect(valueAxisScale(0, 41)).toEqual({ min: 0, max: 50, step: 10 });
  });
  it('negative data floors the min and widens the max with the niced min', () => {
    // step = niceStep(100-(-15)) = niceStep(115) = 20;
    // min = niceAxisMin(-15,20) = -20; max = niceAxisMax(100,20,-20) = ceil((100+6)/20)*20 = 120
    expect(valueAxisScale(-15, 100)).toEqual({ min: -20, max: 120, step: 20 });
  });
  it('explicit min/max override the computed bounds (step still from data range)', () => {
    // step = niceStep(41-0) = 10; explicit min -5, max 60 win
    expect(valueAxisScale(0, 41, -5, 60)).toEqual({ min: -5, max: 60, step: 10 });
  });
  it('a null explicit bound falls back to the auto value', () => {
    expect(valueAxisScale(0, 41, null, 60)).toEqual({ min: 0, max: 60, step: 10 });
    expect(valueAxisScale(0, 41, -5, null)).toEqual({ min: -5, max: 50, step: 10 });
  });
  it('data 3.5 → max 4 with step 0.5 (fine-grained positive range)', () => {
    // step = niceStep(3.5) = 0.5; min = 0; max = niceAxisMax(3.5,0.5,0):
    //   3.5 + 3.5/20 = 3.675 → ceil(3.675/0.5)*0.5 = 4
    expect(valueAxisScale(0, 3.5)).toEqual({ min: 0, max: 4, step: 0.5 });
  });
  it('data 0.1129 → max 0.12 with step 0.02 (sub-unit range)', () => {
    // step = niceStep(0.1129) = 0.02; min = 0;
    //   0.1129 + 0.1129/20 = 0.118545 → ceil(0.118545/0.02)*0.02 = 0.12
    const { min, max, step } = valueAxisScale(0, 0.1129);
    expect(min).toBe(0);
    expect(step).toBeCloseTo(0.02, 12);
    expect(max).toBeCloseTo(0.12, 12);
  });
});

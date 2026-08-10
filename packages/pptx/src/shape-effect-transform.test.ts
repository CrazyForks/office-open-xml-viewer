import { describe, expect, it } from 'vitest';
import { transformedEffectBounds } from './renderer';

describe('shape effect bounds', () => {
  it('includes the active group transform translation and scale', () => {
    expect(transformedEffectBounds(
      { a: 1.5, b: 0, c: 0, d: 1.5, e: 72, f: 45 },
      464,
      120,
      146,
      200,
    )).toEqual({ x: 768, y: 225, w: 219, h: 300 });
  });

  it('returns the axis-aligned device bounds after rotation', () => {
    expect(transformedEffectBounds(
      { a: 0, b: 2, c: -2, d: 0, e: 500, f: 30 },
      10,
      20,
      40,
      30,
    )).toEqual({ x: 400, y: 50, w: 60, h: 80 });
  });
});

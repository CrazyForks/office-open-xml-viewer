import { describe, expect, it } from 'vitest';
import {
  axisAlignedRectsOverlap,
  resolveAxisAlignedOverlap,
} from './axis-aligned-overlap.js';

describe('axis-aligned overlap placement', () => {
  it('treats touching edges as non-overlapping', () => {
    expect(axisAlignedRectsOverlap(
      { left: 0, right: 10, top: 0, bottom: 10 },
      { left: 10, right: 20, top: 0, bottom: 10 },
      0,
    )).toBe(false);
  });

  it('clears each intersecting blocker by moving right when the page permits it', () => {
    expect(resolveAxisAlignedOverlap(
      { left: 0, right: 10, top: 0, bottom: 10 },
      [
        { left: 0, right: 10, top: 0, bottom: 10 },
        { left: 10, right: 20, top: 0, bottom: 10 },
      ],
      { overlapEpsilon: 0, rightBoundary: 30, rightBoundarySlack: 0 },
    )).toEqual({ left: 20, top: 0 });
  });

  it('moves below the blockers when the right edge cannot fit on the page', () => {
    expect(resolveAxisAlignedOverlap(
      { left: 0, right: 10, top: 0, bottom: 10 },
      [
        { left: 0, right: 10, top: 0, bottom: 10 },
        { left: 0, right: 10, top: 10, bottom: 20 },
      ],
      { overlapEpsilon: 0, rightBoundary: 10, rightBoundarySlack: 0 },
    )).toEqual({ left: 0, top: 20 });
  });
});

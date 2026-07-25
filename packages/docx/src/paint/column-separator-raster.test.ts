import { describe, expect, it } from 'vitest';
import { rasterizeColumnSeparator } from './column-separator-raster.js';

describe('rasterizeColumnSeparator', () => {
  it.each([
    {
      scale: 1, dpr: 1,
      input: { start: { xPt: 50.2, yPt: 10.2 }, end: { xPt: 50.2, yPt: 190.2 } },
      expected: {
        segment: { start: { xPt: 50.5, yPt: 10 }, end: { xPt: 50.5, yPt: 190 } },
        widthPt: 1,
      },
    },
    {
      scale: 1.5, dpr: 2,
      input: { start: { xPt: 50.2, yPt: 10.2 }, end: { xPt: 50.2, yPt: 190.2 } },
      expected: {
        segment: {
          start: { xPt: 151 / 3, yPt: 31 / 3 },
          end: { xPt: 151 / 3, yPt: 571 / 3 },
        },
        widthPt: 2 / 3,
      },
    },
  ])('snaps vertical rules at scale=$scale dpr=$dpr', ({ scale, dpr, input, expected }) => {
    expect(rasterizeColumnSeparator(input, scale, dpr)).toEqual(expected);
  });

  it.each([
    {
      scale: 2, dpr: 1,
      input: { start: { xPt: 170.2, yPt: 70.2 }, end: { xPt: 50.2, yPt: 70.2 } },
      expected: {
        segment: { start: { xPt: 170, yPt: 70.25 }, end: { xPt: 50, yPt: 70.25 } },
        widthPt: 0.5,
      },
    },
    {
      scale: 1, dpr: 2,
      input: { start: { xPt: 170.2, yPt: 70.2 }, end: { xPt: 50.2, yPt: 70.2 } },
      expected: {
        segment: { start: { xPt: 170, yPt: 70 }, end: { xPt: 50, yPt: 70 } },
        widthPt: 1,
      },
    },
  ])('snaps horizontal rules at scale=$scale dpr=$dpr', ({ scale, dpr, input, expected }) => {
    expect(rasterizeColumnSeparator(input, scale, dpr)).toEqual(expected);
  });
});

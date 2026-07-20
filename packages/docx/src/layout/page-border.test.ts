import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type { PageBorders } from '../types.js';
import { materializePageBorderLayout } from './page-border.js';

const horizontalSection: SectionLayoutContext = {
  geometry: {
    pageWidth: 200,
    pageHeight: 200,
    marginTop: 20,
    marginRight: 20,
    marginBottom: 20,
    marginLeft: 40,
    headerDistance: 10,
    footerDistance: 10,
  },
  columns: [{ xPt: 40, wPt: 140 }],
  columnSeparator: false,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'lrTb',
  verticalAlignment: 'top',
};

const edge = (over: Partial<NonNullable<PageBorders['top']>> = {}) => ({
  style: 'single',
  color: 'ff0000',
  width: 1,
  space: 4,
  ...over,
});

describe('retained page-border layout', () => {
  it.each([
    ['firstPage', true, true],
    ['firstPage', false, false],
    ['notFirstPage', true, false],
    ['notFirstPage', false, true],
    ['allPages', true, true],
    ['allPages', false, true],
  ] as const)(
    'resolves display=%s for firstSectionOwnedPage=%s',
    (display, firstSectionOwnedPage, visible) => {
      const pageBorders: PageBorders = {
        offsetFrom: 'page',
        display,
        zOrder: 'front',
        top: edge(),
      };

      const result = materializePageBorderLayout(
        pageBorders,
        horizontalSection,
        { widthPt: 200, heightPt: 200 },
        firstSectionOwnedPage,
      );

      expect(result !== null).toBe(visible);
    },
  );

  it('retains text-margin geometry, normalized treatment, and front/back ownership', () => {
    const pageBorders: PageBorders = {
      offsetFrom: 'text',
      display: 'allPages',
      zOrder: 'back',
      top: edge({ style: 'dotDash' }),
      bottom: edge(),
      left: edge(),
      right: edge(),
    };

    const result = materializePageBorderLayout(
      pageBorders,
      horizontalSection,
      { widthPt: 200, heightPt: 200 },
      true,
    );

    expect(result).toMatchObject({
      zOrder: 'back',
      logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      segments: [
        {
          edge: 'top',
          from: { xPt: 44, yPt: 24 },
          to: { xPt: 176, yPt: 24 },
          color: '#ff0000',
          widthPt: 1,
          authoredStyle: 'dotDash',
          style: 'dashed',
          dashPatternPt: [1, 2, 3, 2],
        },
        {
          edge: 'bottom',
          from: { xPt: 44, yPt: 176 },
          to: { xPt: 176, yPt: 176 },
        },
        {
          edge: 'left',
          from: { xPt: 44, yPt: 24 },
          to: { xPt: 44, yPt: 176 },
        },
        {
          edge: 'right',
          from: { xPt: 176, yPt: 24 },
          to: { xPt: 176, yPt: 176 },
        },
      ],
    });
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result?.logicalToPhysical)).toBe(true);
    expect(Object.isFrozen(result?.segments)).toBe(true);
  });

  it.each([
    ['ABCDEF', '#ABCDEF'],
    ['abcdef', '#abcdef'],
    [undefined, '#000000'],
    ['F00', '#000000'],
    ['red', '#000000'],
    ['#ff0000', '#000000'],
    [' ff0000 ', '#000000'],
  ] as const)('normalizes parser-owned color %s to %s', (color, expected) => {
    const result = materializePageBorderLayout(
      {
        offsetFrom: 'page',
        display: 'allPages',
        zOrder: 'front',
        top: edge({ color }),
      },
      horizontalSection,
      { widthPt: 200, heightPt: 200 },
      true,
    );

    expect(result?.segments[0]?.color).toBe(expected);
  });

  it.each([
    [Number.NaN, 4, 0.5, 4],
    [Number.POSITIVE_INFINITY, 4, 0.5, 4],
    [Number.NEGATIVE_INFINITY, 4, 0.5, 4],
    [1, Number.NaN, 1, 0],
    [1, Number.POSITIVE_INFINITY, 1, 0],
    [1, Number.NEGATIVE_INFINITY, 1, 0],
  ])(
    'normalizes non-finite width=%s and space=%s to width=%s and space=%s',
    (width, space, expectedWidth, expectedSpace) => {
      const result = materializePageBorderLayout(
        {
          offsetFrom: 'page',
          display: 'allPages',
          zOrder: 'front',
          top: edge({ width, space }),
        },
        horizontalSection,
        { widthPt: 200, heightPt: 200 },
        true,
      );

      expect(result?.segments[0]).toMatchObject({
        from: { xPt: 0, yPt: expectedSpace },
        to: { xPt: 200, yPt: expectedSpace },
        widthPt: expectedWidth,
      });
      expect(result?.segments[0]?.dashPatternPt?.every(Number.isFinite)).toBe(true);
    },
  );

  it('retains the section logical-to-physical page transform for vertical paint', () => {
    const verticalSection: SectionLayoutContext = {
      ...horizontalSection,
      geometry: {
        ...horizontalSection.geometry,
        pageWidth: 200,
        pageHeight: 100,
      },
      textDirection: 'tbRl',
    };

    const result = materializePageBorderLayout(
      {
        offsetFrom: 'page',
        display: 'allPages',
        zOrder: 'front',
        top: edge({ space: 12 }),
      },
      verticalSection,
      { widthPt: 100, heightPt: 200 },
      true,
    );

    expect(result?.logicalToPhysical).toEqual({
      a: 0,
      b: 1,
      c: -1,
      d: 0,
      e: 100,
      f: 0,
    });
    expect(result?.segments[0]).toMatchObject({
      from: { xPt: 0, yPt: 12 },
      to: { xPt: 200, yPt: 12 },
    });
  });
});

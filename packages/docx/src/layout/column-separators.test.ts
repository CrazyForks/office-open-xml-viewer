import { describe, expect, it } from 'vitest';
import type { PageSectionRegion } from './types.js';
import type { SectionLayoutContext } from '../layout-context.js';
import { columnSeparatorSegments } from './column-separators.js';

const section: SectionLayoutContext = {
  geometry: {
    pageWidth: 200, pageHeight: 200,
    marginTop: 30, marginRight: 20, marginBottom: 50, marginLeft: 20,
    headerDistance: 10, footerDistance: 10,
  },
  columns: [
    { xPt: 20, wPt: 40 },
    { xPt: 80, wPt: 30 },
    { xPt: 130, wPt: 50 },
  ],
  columnSeparator: true,
  grid: { kind: 'none', linePitchPt: null, charSpacePt: null },
  textDirection: 'lrTb',
  verticalAlignment: 'top',
};

const region = (
  overrides: Partial<PageSectionRegion> = {},
  sectionOverrides: Partial<SectionLayoutContext> = {},
): PageSectionRegion => ({
  id: 'region:0',
  sectionOccurrenceId: 'section:0',
  coordinateSpace: {
    writingMode: 'horizontal-tb',
    logicalToPhysical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    physicalToLogical: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
  },
  blockStartPt: 30,
  blockEndPt: 150,
  flowDomainIds: ['column:0', 'column:1', 'column:2'],
  section: { ...section, ...sectionOverrides },
  ...overrides,
});

describe('columnSeparatorSegments', () => {
  it('uses each actual unequal-column gap midpoint and the canonical region band', () => {
    expect(columnSeparatorSegments([region()])).toEqual([
      { start: { xPt: 70, yPt: 30 }, end: { xPt: 70, yPt: 150 } },
      { start: { xPt: 120, yPt: 30 }, end: { xPt: 120, yPt: 150 } },
    ]);
  });

  it('transforms logical vertical rules into physical horizontal segments', () => {
    expect(columnSeparatorSegments([region({
      coordinateSpace: {
        writingMode: 'vertical-rl',
        logicalToPhysical: { a: 0, b: 1, c: -1, d: 0, e: 200, f: 0 },
        physicalToLogical: { a: 0, b: -1, c: 1, d: 0, e: 0, f: 200 },
      },
    })])).toEqual([
      { start: { xPt: 170, yPt: 70 }, end: { xPt: 50, yPt: 70 } },
      { start: { xPt: 170, yPt: 120 }, end: { xPt: 50, yPt: 120 } },
    ]);
  });

  it('returns none for disabled, single-column, and zero-height regions', () => {
    expect(columnSeparatorSegments([
      region({}, { columnSeparator: false }),
      region({}, { columns: [{ xPt: 20, wPt: 160 }] }),
      region({ blockEndPt: 30 }),
    ])).toEqual([]);
  });

  it('keeps separators scoped to enabled regions on a mixed continuous page', () => {
    expect(columnSeparatorSegments([
      region({ blockStartPt: 20, blockEndPt: 70 }, { columnSeparator: false }),
      region({ id: 'region:1', sectionOccurrenceId: 'section:1', blockStartPt: 70, blockEndPt: 140 }),
    ])).toEqual([
      { start: { xPt: 70, yPt: 70 }, end: { xPt: 70, yPt: 140 } },
      { start: { xPt: 120, yPt: 70 }, end: { xPt: 120, yPt: 140 } },
    ]);
  });
});

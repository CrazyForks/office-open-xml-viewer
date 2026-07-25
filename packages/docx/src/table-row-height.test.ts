import { describe, it, expect } from 'vitest';
import { resolveSingleRowHeight, resolveTableRowHeights } from './table-geometry.js';
import { layoutBodyTableRowAdvances } from './test-support/document-layout.test-support.js';
import type {
  DocTable,
  DocTableRow,
  DocTableCell,
  SectionProps,
} from './types.js';

// ECMA-376 §17.4.80 (trHeight) + §17.18.37 (ST_HeightRule): the @hRule attribute
// decides how w:trHeight/@val constrains the row height.
//   exact   — height is exactly @val (overflow clipped).
//   atLeast — @val is a lower bound; content can expand the row.
//   auto    — per the §17.4.80 literal, @val is IGNORED ("no predetermined
//             minimum or maximum size", advisory layout cache only). Word's
//             output PDFs, however, treat @val as a LOWER BOUND (same as
//             atLeast) when hRule is omitted and @val is present — e.g.
//             sample-11.docx's December 2007 calendar emits trHeight w:val=576
//             (no hRule, spec default = auto) on its date rows and Word renders
//             each such row at exactly 576 / 20 = 28.8 pt, matching @val as a
//             floor (the larger per-week cadence is that
//             28.8 pt date row plus an unmarked auto spacer row, not one @val
//             row). XML inspection confirms no other height signal exists. We
//             deliberately deviate from the §17.4.80 literal to match Word's
//             behavior; with @val absent the canonical retained model uses the
//             specification-compatible zero floor and lets content own height.
//
// These tests pin the floor logic with empty cells: with no content, the cell's
// content height is just its (here zero) margins, so the retained row advance is
// governed entirely by the trHeight rule.

const SECTION: SectionProps = {
  pageWidth: 1000,
  pageHeight: 2000,
  marginTop: 0,
  marginRight: 0,
  marginBottom: 0,
  marginLeft: 0,
  headerDistance: 0,
  footerDistance: 0,
  titlePage: false,
  evenAndOddHeaders: false,
};

const MEASURE_CONTEXT = {
  font: '10px serif',
  letterSpacing: '0px',
  measureText: () => ({
    width: 0,
    fontBoundingBoxAscent: 8,
    fontBoundingBoxDescent: 2,
    actualBoundingBoxAscent: 8,
    actualBoundingBoxDescent: 2,
  }) as TextMetrics,
  save() {},
  restore() {},
} as unknown as CanvasRenderingContext2D;

function emptyCell(): DocTableCell {
  return {
    content: [],
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: null,
  } as unknown as DocTableCell;
}

function rowWith(rowHeight: number | null, rule: string): DocTableRow {
  return {
    cells: [emptyCell()],
    rowHeight,
    rowHeightRule: rule,
    isHeader: false,
  } as unknown as DocTableRow;
}

function table(): DocTable {
  return {
    colWidths: [100],
    rows: [],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

describe('canonical retained table row height — ST_HeightRule (§17.4.80 / §17.18.37)', () => {
  const t = table();
  const retainedAdvance = (row: DocTableRow) => {
    const advance = layoutBodyTableRowAdvances(
      { ...t, rows: [row] },
      SECTION,
      MEASURE_CONTEXT,
    )[0];
    if (advance === undefined) throw new Error('Canonical table omitted the row');
    return advance;
  };

  it('exact — height is exactly @val regardless of (here empty) content', () => {
    expect(retainedAdvance(rowWith(600, 'exact'))).toBe(600);
  });

  it('atLeast — @val is a lower bound; empty content cannot shrink below it', () => {
    expect(retainedAdvance(rowWith(600, 'atLeast'))).toBe(600);
  });

  // `word-authored-auto-row-height-floor`: an authored auto height remains a
  // lower bound in the retained model.
  it('auto with @val — @val is honored as a lower bound', () => {
    expect(retainedAdvance(rowWith(600, 'auto'))).toBe(600);
  });

  it('auto with no @val — has no predetermined minimum', () => {
    expect(retainedAdvance(rowWith(null, 'auto'))).toBe(0);
  });

  it('atLeast with no @val — has no authored floor', () => {
    expect(retainedAdvance(rowWith(null, 'atLeast'))).toBe(0);
  });

  it('§17.4.15: measures a cell against columns after gridBefore', () => {
    const measuredWidths: number[] = [];
    const row = {
      ...rowWith(null, 'auto'),
      gridBefore: 1,
      gridAfter: 1,
    } as unknown as DocTableRow;

    resolveSingleRowHeight(row, [20, 40, 60], 1, (_cell, width) => {
      measuredWidths.push(width);
      return 10;
    });

    expect(measuredWidths).toEqual([40]);
  });

  it('includes half of each resolved horizontal border in adjacent non-exact row boxes', () => {
    const single = { style: 'single', width: 0.5, color: '#000000' };
    const rows = Array.from({ length: 4 }, () => ({
      ...rowWith(20.4, 'auto'),
      cells: [{
        ...emptyCell(),
        borders: { top: single, bottom: single, left: null, right: null, insideH: null, insideV: null },
      }],
    } as unknown as DocTableRow));
    const bordered = {
      ...table(),
      rows,
    } as DocTable;

    const heights = resolveTableRowHeights(bordered, [100], 1, () => 0);

    // Five 0.5pt boundary rules contribute half at the two outer edges and a
    // full rule across each of the three shared boundaries: 4 × 20.4 + 2.0.
    expect(heights).toEqual([20.9, 20.9, 20.9, 20.9]);
    expect(heights.reduce((sum, height) => sum + height, 0)).toBeCloseTo(83.6, 8);
  });
});

import { describe, expect, it } from 'vitest';
import {
  dashArray,
  docxBorderDashArray,
  xlsxBorderDashArray,
  pptxDashArray,
} from './dash.js';

describe('dashArray (generic on/off × unit helper)', () => {
  it('scales a relative pattern by unit', () => {
    expect(dashArray([1, 2], 2)).toEqual([2, 4]);
    expect(dashArray([3, 2], 4)).toEqual([12, 8]);
  });
  it('is the identity at unit = 1', () => {
    expect(dashArray([4, 3], 1)).toEqual([4, 3]);
  });
  it('returns [] for an empty (solid) pattern', () => {
    expect(dashArray([], 5)).toEqual([]);
  });
});

// Byte-for-byte equivalence with the former inline docx implementation
// (§17.18.2 ST_Border, lw-relative).
describe('docxBorderDashArray (§17.18.2 ST_Border)', () => {
  const lw = 2;
  it('maps the dash/dot family to lw-scaled patterns', () => {
    expect(docxBorderDashArray('dotted', lw)).toEqual([2, 4]);
    expect(docxBorderDashArray('dashed', lw)).toEqual([6, 4]);
    expect(docxBorderDashArray('dashSmallGap', lw)).toEqual([6, 2]);
    expect(docxBorderDashArray('dotDash', lw)).toEqual([2, 4, 6, 4]);
    expect(docxBorderDashArray('dotDotDash', lw)).toEqual([2, 4, 2, 4, 6, 4]);
    // dashDotStroked (thin/thick alternation) is approximated as dotDash.
    expect(docxBorderDashArray('dashDotStroked', lw)).toEqual([2, 4, 6, 4]);
  });
  it('scales with the border width', () => {
    expect(docxBorderDashArray('dashed', 1)).toEqual([3, 2]);
    expect(docxBorderDashArray('dashed', 4)).toEqual([12, 8]);
  });
  it('returns [] for solid / non-dash styles', () => {
    for (const s of ['single', 'thick', 'triple', 'double', 'wave', 'none', 'nil', 'inset']) {
      expect(docxBorderDashArray(s, lw)).toEqual([]);
    }
  });
});

// Byte-for-byte equivalence with the former inline xlsx implementation
// (§18.18.3 ST_BorderStyle, static px — the medium* variants share the cadence
// of their thin counterparts).
describe('xlsxBorderDashArray (§18.18.3 ST_BorderStyle, static px)', () => {
  it('maps the dash families to static-pixel patterns', () => {
    expect(xlsxBorderDashArray('hair')).toEqual([1, 1]);
    expect(xlsxBorderDashArray('dashed')).toEqual([4, 3]);
    expect(xlsxBorderDashArray('mediumDashed')).toEqual([4, 3]);
    expect(xlsxBorderDashArray('dotted')).toEqual([2, 2]);
    expect(xlsxBorderDashArray('dashDot')).toEqual([4, 2, 1, 2]);
    expect(xlsxBorderDashArray('mediumDashDot')).toEqual([4, 2, 1, 2]);
    expect(xlsxBorderDashArray('dashDotDot')).toEqual([4, 2, 1, 2, 1, 2]);
    expect(xlsxBorderDashArray('mediumDashDotDot')).toEqual([4, 2, 1, 2, 1, 2]);
    expect(xlsxBorderDashArray('slantDashDot')).toEqual([5, 3, 1, 3]);
  });
  it('returns [] for solid styles', () => {
    for (const s of ['thin', 'medium', 'thick', 'double', '']) {
      expect(xlsxBorderDashArray(s)).toEqual([]);
    }
  });
});

// Byte-for-byte equivalence with the former inline pptx implementation
// (§20.1.10.49 ST_PresetLineDashVal, lineW-relative — the *Heavy variants share
// the base cadence). The old call site did `dashFor(s).map(v => v*lineW)`.
describe('pptxDashArray (§20.1.10.49 ST_PresetLineDashVal)', () => {
  const lineW = 2;
  it('maps the preset dash names to lineW-scaled patterns', () => {
    expect(pptxDashArray('dotted', lineW)).toEqual([3, 6]);
    expect(pptxDashArray('dottedHeavy', lineW)).toEqual([3, 6]);
    expect(pptxDashArray('dash', lineW)).toEqual([12, 6]);
    expect(pptxDashArray('dashHeavy', lineW)).toEqual([12, 6]);
    expect(pptxDashArray('dashLong', lineW)).toEqual([20, 8]);
    expect(pptxDashArray('dashLongHeavy', lineW)).toEqual([20, 8]);
    expect(pptxDashArray('dotDash', lineW)).toEqual([12, 6, 3, 6]);
    expect(pptxDashArray('dotDashHeavy', lineW)).toEqual([12, 6, 3, 6]);
    expect(pptxDashArray('dotDotDash', lineW)).toEqual([12, 6, 3, 6, 3, 6]);
    expect(pptxDashArray('dotDotDashHeavy', lineW)).toEqual([12, 6, 3, 6, 3, 6]);
  });
  it('returns [] for solid styles (sng / unknown)', () => {
    expect(pptxDashArray('sng', lineW)).toEqual([]);
    expect(pptxDashArray('solid', lineW)).toEqual([]);
  });
});

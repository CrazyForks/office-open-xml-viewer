import { describe, it, expect } from 'vitest';
import { cellInSqref, findListValidationAt } from './data-validation.js';
import type { DataValidation } from './types.js';

describe('cellInSqref — ECMA-376 §18.3.1.33 @sqref matching', () => {
  it('matches a single-cell sqref', () => {
    expect(cellInSqref('B2', 2, 2)).toBe(true);
    expect(cellInSqref('B2', 2, 3)).toBe(false);
  });

  it('matches a rectangular range', () => {
    // C3:E6 → rows 3-6, cols 3-5
    expect(cellInSqref('C3:E6', 3, 3)).toBe(true);
    expect(cellInSqref('C3:E6', 6, 5)).toBe(true);
    expect(cellInSqref('C3:E6', 4, 4)).toBe(true);
    expect(cellInSqref('C3:E6', 2, 3)).toBe(false); // row above
    expect(cellInSqref('C3:E6', 3, 6)).toBe(false); // col right
  });

  it('handles a reversed range (Excel writes from active corner)', () => {
    expect(cellInSqref('E6:C3', 4, 4)).toBe(true);
  });

  it('matches across a space-separated multi-range sqref', () => {
    const sq = 'A1 C3:D4 F10';
    expect(cellInSqref(sq, 1, 1)).toBe(true);
    expect(cellInSqref(sq, 3, 4)).toBe(true);
    expect(cellInSqref(sq, 10, 6)).toBe(true);
    expect(cellInSqref(sq, 5, 5)).toBe(false);
  });

  it('returns false for malformed sqref tokens without throwing', () => {
    expect(cellInSqref('', 1, 1)).toBe(false);
    expect(cellInSqref('not-a-ref', 1, 1)).toBe(false);
  });
});

describe('findListValidationAt — list-type dropdown gating', () => {
  const listDv: DataValidation = {
    sqref: 'B2:B10',
    validationType: 'list',
    formula1: '"Yes,No,Maybe"',
  };
  const wholeDv: DataValidation = {
    sqref: 'C2:C10',
    validationType: 'whole',
    operator: 'between',
    formula1: '1',
    formula2: '100',
  };

  it('returns the list rule when the cell is inside its sqref', () => {
    expect(findListValidationAt([wholeDv, listDv], 5, 2)).toBe(listDv);
  });

  it('ignores non-list validation types (no dropdown for whole/decimal/etc.)', () => {
    expect(findListValidationAt([wholeDv], 5, 3)).toBeNull();
  });

  it('returns null when the cell is outside every list rule', () => {
    expect(findListValidationAt([listDv], 11, 2)).toBeNull();
    expect(findListValidationAt([listDv], 5, 3)).toBeNull();
  });

  it('returns null for an empty / undefined rule set', () => {
    expect(findListValidationAt([], 1, 1)).toBeNull();
    expect(findListValidationAt(undefined, 1, 1)).toBeNull();
  });
});

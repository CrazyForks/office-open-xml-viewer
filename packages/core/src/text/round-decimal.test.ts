import { describe, it, expect } from 'vitest';
import { roundDecimalHalfUp } from './round-decimal';

describe('roundDecimalHalfUp — Office-style decimal (round-half-up) formatting', () => {
  // Office (Excel/Word/PowerPoint) rounds on the DECIMAL representation, so a
  // `.xx5` boundary rounds UP. JavaScript `toFixed` rounds the binary double,
  // which for these values sits just below the .5 boundary → rounds down.
  describe('.xx5 boundaries that toFixed gets wrong', () => {
    const cases: Array<[number, number, string, string]> = [
      // [value, digits, expected (Excel), toFixed (wrong)]. Only values whose
      // nearest double sits just BELOW the .xx5 boundary (so toFixed rounds
      // down) — not e.g. 0.125, which is exactly representable and toFixed
      // already rounds up.
      [2.675, 2, '2.68', '2.67'],
      [1.005, 2, '1.01', '1.00'],
      [9.995, 2, '10.00', '9.99'],
      [1.255, 2, '1.26', '1.25'],
      [8.575, 2, '8.58', '8.57'],
    ];
    for (const [value, digits, expected, toFixedResult] of cases) {
      it(`${value} @ ${digits} → "${expected}" (toFixed gives "${toFixedResult}")`, () => {
        expect(roundDecimalHalfUp(value, digits)).toBe(expected);
        // Guard: confirm this really is a case toFixed mishandles, so the test
        // is a genuine regression pin, not a tautology.
        expect(value.toFixed(digits)).toBe(toFixedResult);
      });
    }
  });

  describe('carry propagation', () => {
    it('9.995 @ 2 → "10.00" (carry ripples through all places)', () => {
      expect(roundDecimalHalfUp(9.995, 2)).toBe('10.00');
    });
    it('0.995 @ 2 → "1.00"', () => {
      expect(roundDecimalHalfUp(0.995, 2)).toBe('1.00');
    });
    it('99.95 @ 1 → "100.0"', () => {
      expect(roundDecimalHalfUp(99.95, 1)).toBe('100.0');
    });
    it('9.5 @ 0 → "10"', () => {
      expect(roundDecimalHalfUp(9.5, 0)).toBe('10');
    });
  });

  describe('negatives round half AWAY from zero (Excel display magnitude)', () => {
    it('-2.675 @ 2 → "-2.68"', () => {
      expect(roundDecimalHalfUp(-2.675, 2)).toBe('-2.68');
    });
    it('-1.005 @ 2 → "-1.01"', () => {
      expect(roundDecimalHalfUp(-1.005, 2)).toBe('-1.01');
    });
    it('-9.995 @ 2 → "-10.00"', () => {
      expect(roundDecimalHalfUp(-9.995, 2)).toBe('-10.00');
    });
    it('-0.5 @ 0 → "-1"', () => {
      expect(roundDecimalHalfUp(-0.5, 0)).toBe('-1');
    });
  });

  describe('digits / padding', () => {
    it('pads to the requested precision (0.5 @ 2 → "0.50")', () => {
      expect(roundDecimalHalfUp(0.5, 2)).toBe('0.50');
    });
    it('digits=0 emits an integer with no dot (3.14 → "3")', () => {
      expect(roundDecimalHalfUp(3.14, 0)).toBe('3');
    });
    it('exact integers pad correctly (5 @ 2 → "5.00")', () => {
      expect(roundDecimalHalfUp(5, 2)).toBe('5.00');
    });
    it('value already shorter than digits pads (1.2 @ 3 → "1.200")', () => {
      expect(roundDecimalHalfUp(1.2, 3)).toBe('1.200');
    });
  });

  describe('non-boundary values match toFixed (no drift on normal data)', () => {
    const cases: Array<[number, number, string]> = [
      [3.14159, 2, '3.14'],
      [3.14659, 2, '3.15'],
      [1234.5678, 2, '1234.57'],
      [0.1, 2, '0.10'],
      [0, 2, '0.00'],
      [-3.14159, 2, '-3.14'],
      [1000, 0, '1000'],
      [2.4, 0, '2'],
      [2.6, 0, '3'],
    ];
    for (const [value, digits, expected] of cases) {
      it(`${value} @ ${digits} → "${expected}"`, () => {
        expect(roundDecimalHalfUp(value, digits)).toBe(expected);
        // These are NOT .xx5 boundaries, so Office and toFixed agree.
        expect(value.toFixed(digits)).toBe(expected);
      });
    }
  });

  describe('edge inputs', () => {
    it('negative zero normalizes to "0.00" (no "-0.00")', () => {
      expect(roundDecimalHalfUp(-0, 2)).toBe('0.00');
      // A tiny negative that rounds to zero must not carry a minus sign.
      expect(roundDecimalHalfUp(-0.0001, 2)).toBe('0.00');
    });
    it('non-finite falls back to the raw string', () => {
      expect(roundDecimalHalfUp(NaN, 2)).toBe('NaN');
      expect(roundDecimalHalfUp(Infinity, 2)).toBe('Infinity');
      expect(roundDecimalHalfUp(-Infinity, 2)).toBe('-Infinity');
    });
    it('large magnitudes keep integer precision (12345678.905 @ 2 → "12345678.91")', () => {
      expect(roundDecimalHalfUp(12345678.905, 2)).toBe('12345678.91');
    });
  });
});

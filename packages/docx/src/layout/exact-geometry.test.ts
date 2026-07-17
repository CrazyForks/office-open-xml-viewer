import { describe, expect, it } from 'vitest';
import {
  compareExactRational,
  decodeBinary64,
  exactRationalToNumberDown,
  exactRationalToNumber,
  exactRationalToNumberUp,
  normalizeExactRational,
} from './exact-geometry.js';

describe('exact binary64 geometry primitives', () => {
  it('decodes the complete finite binary64 range as canonical dyadics', () => {
    expect(decodeBinary64(1)).toEqual({ coefficient: 1n, exponent: 0 });
    expect(decodeBinary64(-0.5)).toEqual({ coefficient: -1n, exponent: -1 });
    expect(decodeBinary64(Number.MIN_VALUE))
      .toEqual({ coefficient: 1n, exponent: -1074 });
    expect(decodeBinary64(Number.MAX_VALUE))
      .toEqual({ coefficient: 9007199254740991n, exponent: 971 });
    expect(decodeBinary64(-0)).toEqual({ coefficient: 0n, exponent: 0 });
  });

  it('normalizes signs, common factors, and zero', () => {
    expect(normalizeExactRational(-2n, -4n))
      .toEqual({ numerator: 1n, denominator: 2n });
    expect(normalizeExactRational(0n, -9n))
      .toEqual({ numerator: 0n, denominator: 1n });
  });

  it('orders rationals without converting either operand to Number', () => {
    const left = normalizeExactRational((1n << 400n) + 1n, 3n);
    const right = normalizeExactRational(1n << 400n, 3n);

    expect(compareExactRational(left, right)).toBeGreaterThan(0);
    expect(compareExactRational(right, left)).toBeLessThan(0);
    expect(compareExactRational(left, left)).toBe(0);
  });

  it('rounds exact halfway values with binary64 ties-to-even', () => {
    expect(exactRationalToNumber(
      normalizeExactRational((1n << 53n) + 1n, 1n << 53n),
    )).toBe(1);
    expect(exactRationalToNumber(
      normalizeExactRational((1n << 53n) + 3n, 1n << 53n),
    )).toBe(1 + 2 ** -51);
    expect(exactRationalToNumber(
      normalizeExactRational(1n, 1n << 1075n),
    )).toBe(0);
    expect(exactRationalToNumber(
      normalizeExactRational(3n, 1n << 1075n),
    )).toBe(Number.MIN_VALUE * 2);
  });

  it('rounds threshold roots upward to the first usable binary64 value', () => {
    expect(exactRationalToNumberUp(
      normalizeExactRational(30n, 13n),
    )).toBe(2.307692307692308);
    expect(exactRationalToNumberUp(
      normalizeExactRational(6n, 2n),
    )).toBe(3);
    expect(exactRationalToNumberUp(
      normalizeExactRational(-30n, 13n),
    )).toBe(-2.3076923076923075);
    expect(exactRationalToNumberUp(
      normalizeExactRational(-(1n << 2000n), 1n),
    )).toBe(-Number.MAX_VALUE);
    expect(exactRationalToNumberDown(
      normalizeExactRational(30n, 13n),
    )).toBe(2.3076923076923075);
  });
});

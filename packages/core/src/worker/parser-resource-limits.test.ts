import { describe, expect, it } from 'vitest';
import { parserResourceLimitsForWasm } from './parser-resource-limits';

describe('parserResourceLimitsForWasm', () => {
  it('maps positive safe byte counts to optional u64 values', () => {
    expect(
      parserResourceLimitsForWasm({
        maxParsedPartInflatedBytes: 1024,
        maxOperationInflatedBytes: 4096,
      }),
    ).toEqual([1024n, 4096n]);
  });

  it('keeps omitted limits disabled', () => {
    expect(parserResourceLimitsForWasm(undefined)).toEqual([undefined, undefined]);
    expect(parserResourceLimitsForWasm({})).toEqual([undefined, undefined]);
  });

  it.each([0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Number.MAX_SAFE_INTEGER + 1])(
    'rejects an invalid supplied byte limit (%s)',
    (value) => {
      expect(() =>
        parserResourceLimitsForWasm({ maxParsedPartInflatedBytes: value }),
      ).toThrow(/positive safe integer/);
      expect(() =>
        parserResourceLimitsForWasm({ maxOperationInflatedBytes: value }),
      ).toThrow(/positive safe integer/);
    },
  );

  it('rejects a non-number supplied at runtime', () => {
    expect(() =>
      parserResourceLimitsForWasm({
        maxParsedPartInflatedBytes: '1024' as unknown as number,
      }),
    ).toThrow(/positive safe integer/);
  });
});

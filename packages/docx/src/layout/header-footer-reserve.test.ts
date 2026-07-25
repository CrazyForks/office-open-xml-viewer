import { describe, expect, it } from 'vitest';
import {
  convergeHeaderFooterReserves,
  headerFooterOverflowReservePt,
} from './header-footer-reserve.js';

describe('header/footer body reserve', () => {
  it('charges only overflow beyond the signed-margin allowance', () => {
    expect(headerFooterOverflowReservePt(30, 72, 36)).toBe(0);
    expect(headerFooterOverflowReservePt(48, 72, 36)).toBe(12);
    expect(headerFooterOverflowReservePt(120, -72, 36)).toBe(0);
  });

  it('returns the repaginated candidate when the exact next-pass inputs are stable', () => {
    type Candidate = Readonly<{
      geometryVersion: string;
      fieldContexts: readonly Readonly<{
        pageIndex: number;
        displayPageNumber: number;
        pageNumberFormat: string;
      }>[];
    }>;
    const fieldContexts = Object.freeze([
      Object.freeze({ pageIndex: 0, displayPageNumber: 1, pageNumberFormat: 'decimal' }),
    ]);
    const seed: Candidate = Object.freeze({ geometryVersion: 'seed', fieldContexts });
    let repaginations = 0;

    const result = convergeHeaderFooterReserves({
      seed,
      measure: () => [Object.freeze({ top: 12, bottom: 0 })],
      repaginate: (_reserves, current) => {
        repaginations += 1;
        return Object.freeze({ ...current, geometryVersion: 'repaginated' });
      },
      identity: (candidate) => candidate.fieldContexts,
    });

    expect(result.result.geometryVersion).toBe('repaginated');
    expect(repaginations).toBe(1);
  });
});

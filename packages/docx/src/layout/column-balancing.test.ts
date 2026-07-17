import { describe, expect, it } from 'vitest';
import { minimumColumnBalanceTarget } from './column-balancing.js';

describe('minimumColumnBalanceTarget', () => {
  it('selects the exact lowest constraint boundary without an epsilon', () => {
    const constraints = [80, 60, 45];

    const target = minimumColumnBalanceTarget(100, (candidate) => {
      const required = constraints.find((constraint) => candidate >= constraint);
      return required === undefined
        ? { fits: false, requiredTargetPt: 0, thresholdsPt: constraints }
        : { fits: true, requiredTargetPt: required, thresholdsPt: constraints };
    });

    expect(target).toBe(45);
  });

  it('retains zero for an empty balanced section', () => {
    expect(minimumColumnBalanceTarget(100, () => ({
      fits: true,
      requiredTargetPt: 0,
      thresholdsPt: [0],
    }))).toBe(0);
  });

  it('finds a lower legal partition after an immediately smaller cap fails', () => {
    expect(minimumColumnBalanceTarget(80, (candidate) => ({
      fits: candidate === 80 || candidate === 60 || candidate === 40,
      requiredTargetPt: candidate,
      thresholdsPt: [20, 40, 60, 80],
    }))).toBe(40);
  });

  it('rejects a probe that exceeds the candidate target', () => {
    expect(() => minimumColumnBalanceTarget(10, () => ({
      fits: true,
      requiredTargetPt: 11,
      thresholdsPt: [],
    }))).toThrow('invalid occupied boundary');
  });
});

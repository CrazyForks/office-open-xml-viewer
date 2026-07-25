import { describe, expect, it } from 'vitest';
import {
  ExactConvergenceError,
  convergeExactState,
  convergeLayout,
  type LayoutIteration,
} from './convergence.js';

const iteration = (fingerprint: string, pageCount: number): LayoutIteration => ({
  fingerprint,
  pageCount,
});

describe('convergeLayout', () => {
  it('returns the newest iteration when its relevant geometry fingerprint stabilizes', () => {
    const seed = iteration('a', 1);
    const calls: string[] = [];
    const result = convergeLayout(seed, (current) => {
      calls.push(current.fingerprint);
      return current.fingerprint === 'a' ? iteration('b', 2) : iteration('b', 3);
    }, 5);

    expect(calls).toEqual(['a', 'b']);
    expect(result).toEqual(iteration('b', 3));
    expect(result).not.toBe(seed);
  });

  it('throws NON_CONVERGENCE for a repeated cycle', () => {
    expect(() => convergeLayout(
      iteration('a', 1),
      (current) => current.fingerprint === 'a' ? iteration('b', 2) : iteration('a', 3),
      5,
    )).toThrow(/NON_CONVERGENCE.*cycle/i);
  });

  it('throws NON_CONVERGENCE at the hard limit', () => {
    expect(() => convergeLayout(
      iteration('0', 0),
      (current) => iteration(String(Number(current.fingerprint) + 1), current.pageCount + 1),
      2,
    )).toThrow(/NON_CONVERGENCE.*limit/i);
  });
});

describe('convergeExactState', () => {
  it('returns the confirming pass value for an adjacent fixed point', () => {
    const states = ['A', 'B', 'B'] as const;

    const result = convergeExactState<{ state: string; pass: number }>({
      step: (_previous, pass) => ({ state: states[pass - 1]!, pass }),
      stateOf: (value) => value.state,
      limit: states.length,
    });

    expect(result).toEqual({
      value: { state: 'B', pass: 3 },
      passes: 3,
    });
  });

  it('allows the first pass to confirm an explicitly observed seed state', () => {
    const result = convergeExactState<{ state: string; pass: number }>({
      seedState: 'A',
      step: (_previous, pass) => ({ state: 'A', pass }),
      stateOf: (value) => value.state,
      limit: 1,
    });

    expect(result).toEqual({
      value: { state: 'A', pass: 1 },
      passes: 1,
    });
  });

  it('diagnoses a non-adjacent exact-state cycle with its ordered states', () => {
    const states = ['A', 'B', 'A'] as const;

    try {
      convergeExactState<{ state: string; pass: number }>({
        step: (_previous, pass) => ({ state: states[pass - 1]!, pass }),
        stateOf: (value) => value.state,
        limit: states.length,
      });
      throw new Error('Expected exact convergence to reject the cycle');
    } catch (error) {
      expect(error).toBeInstanceOf(ExactConvergenceError);
      expect(error).toMatchObject({
        code: 'NON_CONVERGENCE',
        reason: 'cycle',
        states: ['A', 'B', 'A'],
        passes: 3,
      });
    }
  });

  it('hard-fails an all-distinct deterministic orbit after exactly the pass budget', () => {
    let calls = 0;

    try {
      convergeExactState<{ state: string; pass: number }>({
        step: (_previous, pass) => {
          calls += 1;
          return { state: String(pass), pass };
        },
        stateOf: (value) => value.state,
        limit: 4,
      });
      throw new Error('Expected exact convergence to exhaust the pass budget');
    } catch (error) {
      expect(error).toBeInstanceOf(ExactConvergenceError);
      expect(error).toMatchObject({
        code: 'NON_CONVERGENCE',
        reason: 'limit',
        states: ['1', '2', '3', '4'],
        passes: 4,
      });
      expect(calls).toBe(4);
    }
  });

  it('rejects a budget that cannot observe and confirm a state', () => {
    expect(() => convergeExactState<{ state: string }>({
      step: () => ({ state: 'A' }),
      stateOf: (value) => value.state,
      limit: 1,
    })).toThrow(RangeError);
    expect(() => convergeExactState<{ state: string }>({
      seedState: 'A',
      step: () => ({ state: 'A' }),
      stateOf: (value) => value.state,
      limit: 0,
    })).toThrow(RangeError);
  });
});

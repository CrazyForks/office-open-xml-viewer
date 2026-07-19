import { describe, expect, it } from 'vitest';
import { createExactStateTracker, observeExactState } from './repeated-state.js';

describe('exact repeated-state tracking', () => {
  it('distinguishes an A -> B -> A cycle from an adjacent fixed point', () => {
    const tracker = createExactStateTracker();

    expect(observeExactState(tracker, 'A')).toMatchObject({ kind: 'advance' });
    expect(observeExactState(tracker, 'B')).toMatchObject({ kind: 'advance' });
    expect(observeExactState(tracker, 'A')).toEqual({
      kind: 'cycle',
      states: ['A', 'B', 'A'],
    });
  });

  it('recognizes an adjacent repeated state as convergence', () => {
    const tracker = createExactStateTracker();

    expect(observeExactState(tracker, 'A')).toMatchObject({ kind: 'advance' });
    expect(observeExactState(tracker, 'A')).toEqual({
      kind: 'fixed',
      states: ['A', 'A'],
    });
  });
});

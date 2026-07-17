import { describe, expect, it } from 'vitest';
import {
  cloneSegmentsForLinePass,
  convergeLineWrap,
  LineWrapNonConvergenceError,
} from './line-wrap-convergence.js';

interface TestSegment {
  text: string;
  measuredWidth: number;
}

interface TestLine {
  consumedEnd: Readonly<{ segIndex: number; charOffset: number }>;
  topY: number;
  xOffset: number;
  availWidth: number;
  segments: readonly Readonly<{
    src?: Readonly<{ segIndex: number; charOffset: number }>;
    text?: string;
  }>[];
  probeHeight: number;
}

function line(state: 'A' | 'B'): TestLine {
  return {
    consumedEnd: { segIndex: 0, charOffset: state === 'A' ? 1 : 2 },
    topY: state === 'A' ? 0 : 10,
    xOffset: 0,
    availWidth: state === 'A' ? 100 : 90,
    segments: [{
      src: { segIndex: 0, charOffset: 0 },
      text: state,
    }],
    probeHeight: state === 'A' ? 10 : 20,
  };
}

describe('line wrap convergence', () => {
  it('remeasures until the complete line state is an adjacent fixed point', () => {
    const probes: (readonly number[] | null)[] = [];

    const result = convergeLineWrap<TestLine>(
      (currentProbes) => {
        probes.push(currentProbes);
        return [line('A')];
      },
      (currentLine) => currentLine.probeHeight,
    );

    expect(result).toEqual([line('A')]);
    expect(probes).toEqual([null, [10]]);
  });

  it('diagnoses an exact non-adjacent line-state cycle without a pass budget', () => {
    let pass = 0;

    expect(() => convergeLineWrap<TestLine>(
      () => [line(pass++ % 2 === 0 ? 'A' : 'B')],
      (currentLine) => currentLine.probeHeight,
    )).toThrow(LineWrapNonConvergenceError);
  });

  it('clones mutable pass-local segment records without cloning their authorities', () => {
    const authority = Object.freeze({ id: 'shape-authority' });
    const segments = [{ text: 'A', measuredWidth: 0, authority }];

    const cloned = cloneSegmentsForLinePass(segments);
    cloned[0]!.measuredWidth = 12;

    expect(cloned).not.toBe(segments);
    expect(cloned[0]).not.toBe(segments[0]);
    expect(cloned[0]!.authority).toBe(authority);
    expect(segments[0]!.measuredWidth).toBe(0);
  });
});

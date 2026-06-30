import { describe, it, expect } from 'vitest';
import { lineEndRetract, retractLineEndpoint } from './arrow';
import type { ArrowEnd, Stroke } from '../types/common';

const stroke: Stroke = { color: '#000000', width: 10 };
const end = (type: string, w = 'med', len = 'med'): ArrowEnd => ({ type, w, len });

describe('lineEndRetract — how far to pull the leader line back from the tip', () => {
  // A filled decoration (triangle/stealth/diamond/oval) covers the segment from
  // the tip back to `-len`; the line must stop there so its cap hides inside the
  // decoration. An open `arrow` (V) and `none` leave the line running to the tip.
  // lw = max(0.5, width·scale) = 10; med len multiplier = 6 → len = 60.
  it('triangle / stealth / diamond / oval retract by the decoration length', () => {
    expect(lineEndRetract(end('triangle'), stroke, 1)).toBeCloseTo(60, 5);
    expect(lineEndRetract(end('stealth'), stroke, 1)).toBeCloseTo(60, 5);
    expect(lineEndRetract(end('diamond'), stroke, 1)).toBeCloseTo(60, 5);
    expect(lineEndRetract(end('oval'), stroke, 1)).toBeCloseTo(60, 5);
  });

  it('open arrow and none do not retract (line reaches the tip)', () => {
    expect(lineEndRetract(end('arrow'), stroke, 1)).toBe(0);
    expect(lineEndRetract(end('none'), stroke, 1)).toBe(0);
  });

  it('scales with len multiplier (lg = 8·lw) and stroke width·scale', () => {
    expect(lineEndRetract(end('triangle', 'med', 'lg'), stroke, 1)).toBeCloseTo(80, 5);
    expect(lineEndRetract(end('triangle', 'med', 'sm'), stroke, 1)).toBeCloseTo(40, 5);
    // width 5, scale 2 → lw = 10 → same as above
    expect(lineEndRetract(end('triangle'), { color: '#000', width: 5 }, 2)).toBeCloseTo(60, 5);
  });
});

describe('retractLineEndpoint — pull a point toward its neighbour by `amount`', () => {
  it('returns the point unchanged when amount is 0', () => {
    expect(retractLineEndpoint({ x: 0, y: 0 }, { x: 10, y: 0 }, 0)).toEqual({ x: 0, y: 0 });
  });

  it('moves the point toward the neighbour by `amount`', () => {
    const p = retractLineEndpoint({ x: 0, y: 0 }, { x: 10, y: 0 }, 3);
    expect(p.x).toBeCloseTo(3, 5);
    expect(p.y).toBeCloseTo(0, 5);
  });

  it('does not overshoot past the neighbour', () => {
    const p = retractLineEndpoint({ x: 0, y: 0 }, { x: 10, y: 0 }, 25);
    expect(p.x).toBeCloseTo(10, 5);
    expect(p.y).toBeCloseTo(0, 5);
  });

  it('works on a diagonal (length 5 → 3-4-5 triangle)', () => {
    const p = retractLineEndpoint({ x: 0, y: 0 }, { x: 3, y: 4 }, 2.5);
    expect(p.x).toBeCloseTo(1.5, 5);
    expect(p.y).toBeCloseTo(2.0, 5);
  });
});

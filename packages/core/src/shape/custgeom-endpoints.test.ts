import { describe, it, expect } from 'vitest';
import type { PathCmd } from '../types/common';
import { getCustGeomEndpoints, type CustGeomEndpoint } from './custgeom-endpoints';

/** atan2 of a returned endpoint's normalised tangent, rounded for comparison. */
function angleOf(e: CustGeomEndpoint | null): number {
  if (!e) throw new Error('expected endpoint, got null');
  return Math.atan2(e.dy, e.dx);
}

function near(a: number, b: number, eps = 1e-9): boolean {
  return Math.abs(a - b) <= eps;
}

/** Compare two angles as directions, treating +π and -π (and 2π wraps) as equal. */
function angleNear(a: number, b: number, eps = 1e-9): boolean {
  let d = a - b;
  while (d > Math.PI) d -= 2 * Math.PI;
  while (d < -Math.PI) d += 2 * Math.PI;
  return Math.abs(d) <= eps;
}

describe('getCustGeomEndpoints', () => {
  it('returns null/null for empty input', () => {
    expect(getCustGeomEndpoints([])).toEqual({ start: null, end: null });
    expect(getCustGeomEndpoints([[]])).toEqual({ start: null, end: null });
  });

  describe('lineTo terminal', () => {
    // moveTo(0,0) -> lineTo(1,1): an open diagonal.
    const path: PathCmd[][] = [[
      { cmd: 'moveTo', x: 0, y: 0 },
      { cmd: 'lineTo', x: 1, y: 1 },
    ]];

    it('start sits at the first moveTo point', () => {
      const { start } = getCustGeomEndpoints(path);
      expect(start?.x).toBe(0);
      expect(start?.y).toBe(0);
    });

    it('start tangent points OUTWARD (away from the next point)', () => {
      // First drawn segment direction is (1,1); outward = reverse = (-1,-1).
      const { start } = getCustGeomEndpoints(path);
      expect(angleNear(angleOf(start), Math.atan2(-1, -1))).toBe(true);
    });

    it('end sits at the last lineTo point', () => {
      const { end } = getCustGeomEndpoints(path);
      expect(end?.x).toBe(1);
      expect(end?.y).toBe(1);
    });

    it('end tangent points along the incoming segment (outward)', () => {
      // Travel direction into the end is (1,1); outward at the tail is the same.
      const { end } = getCustGeomEndpoints(path);
      expect(angleNear(angleOf(end), Math.atan2(1, 1))).toBe(true);
    });
  });

  describe('cubicBezTo terminal', () => {
    it('end tangent uses control point (x2,y2) -> end vector', () => {
      // moveTo(0,0) -> C (0.2,0) (0.8,0.2) (1,1). End tangent = (1-0.8, 1-0.2)=(0.2,0.8).
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'cubicBezTo', x1: 0.2, y1: 0, x2: 0.8, y2: 0.2, x: 1, y: 1 },
      ]];
      const { end } = getCustGeomEndpoints(path);
      expect(end?.x).toBe(1);
      expect(end?.y).toBe(1);
      expect(angleNear(angleOf(end), Math.atan2(0.8, 0.2))).toBe(true);
    });

    it('start tangent uses moveTo -> control point (x1,y1) vector, reversed', () => {
      // First drawn cmd is the cubic; start tangent forward = (x1,y1)-(0,0)=(0.2,0).
      // Outward (head) = reversed = (-0.2, 0).
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'cubicBezTo', x1: 0.2, y1: 0, x2: 0.8, y2: 0.2, x: 1, y: 1 },
      ]];
      const { start } = getCustGeomEndpoints(path);
      expect(angleNear(angleOf(start), Math.atan2(0, -0.2))).toBe(true);
    });

    it('degenerate (x2==x && y2==y): falls back to (x1,y1) -> end', () => {
      // C (0.1,0.3) (1,1) (1,1): control2 == end, so use (x1,y1)->end = (0.9,0.7).
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'cubicBezTo', x1: 0.1, y1: 0.3, x2: 1, y2: 1, x: 1, y: 1 },
      ]];
      const { end } = getCustGeomEndpoints(path);
      expect(angleNear(angleOf(end), Math.atan2(1 - 0.3, 1 - 0.1))).toBe(true);
    });

    it('fully degenerate (both controls == end): falls back to start -> end', () => {
      // moveTo(0.2,0.2) -> C (1,1) (1,1) (1,1): use prevPoint(0.2,0.2)->end(1,1).
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0.2, y: 0.2 },
        { cmd: 'cubicBezTo', x1: 1, y1: 1, x2: 1, y2: 1, x: 1, y: 1 },
      ]];
      const { end } = getCustGeomEndpoints(path);
      expect(angleNear(angleOf(end), Math.atan2(0.8, 0.8))).toBe(true);
    });
  });

  describe('arcTo terminal', () => {
    // moveTo at the 3 o'clock point of a unit-ish ellipse, sweep +90deg (CW in
    // screen coords where +y is down). Pen starts at stAng=0.
    // center = pen - (wr*cos0, hr*sin0) = (0.5,0.5) - (0.4,0) = (0.1,0.5).
    // end at stAng+sw = 90deg: (0.1 + 0.4*cos90, 0.5 + 0.3*sin90) = (0.1, 0.8).
    it('positive sweep: end point and outward tangent are analytic', () => {
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0.5, y: 0.5 },
        { cmd: 'arcTo', wr: 0.4, hr: 0.3, stAng: 0, swAng: 90 },
      ]];
      const { end } = getCustGeomEndpoints(path);
      expect(near(end!.x, 0.1, 1e-9)).toBe(true);
      expect(near(end!.y, 0.8, 1e-9)).toBe(true);
      // Tangent of (cx+wr cos t, cy+hr sin t) is (-wr sin t, hr cos t); for
      // sw>0 travel is increasing t, so outward tangent at t=90deg:
      // (-0.4*sin90, 0.3*cos90) = (-0.4, 0).
      expect(angleNear(angleOf(end), Math.atan2(0, -0.4))).toBe(true);
    });

    it('negative sweep flips the direction of travel', () => {
      // Same geometry, sweep -90deg. end at t=-90deg:
      // (0.1 + 0.4*cos(-90), 0.5 + 0.3*sin(-90)) = (0.1, 0.2).
      // For sw<0 travel is decreasing t, outward tangent = -(-wr sin t, hr cos t)
      // at t=-90deg: -(-0.4*sin(-90), 0.3*cos(-90)) = -((0.4),(0)) = (-0.4, 0).
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0.5, y: 0.5 },
        { cmd: 'arcTo', wr: 0.4, hr: 0.3, stAng: 0, swAng: -90 },
      ]];
      const { end } = getCustGeomEndpoints(path);
      expect(near(end!.x, 0.1, 1e-9)).toBe(true);
      expect(near(end!.y, 0.2, 1e-9)).toBe(true);
      expect(angleNear(angleOf(end), Math.atan2(0, -0.4))).toBe(true);
    });
  });

  describe('multiple subpaths', () => {
    it('start comes from the first subpath, end from the last', () => {
      const path: PathCmd[][] = [
        [ { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 0.3, y: 0 } ],
        [ { cmd: 'moveTo', x: 0.6, y: 0.6 }, { cmd: 'lineTo', x: 1, y: 1 } ],
      ];
      const { start, end } = getCustGeomEndpoints(path);
      expect(start?.x).toBe(0);
      expect(start?.y).toBe(0);
      expect(end?.x).toBe(1);
      expect(end?.y).toBe(1);
    });
  });

  describe('closed paths produce no endpoints', () => {
    it('explicit close suppresses both arrowheads', () => {
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0, y: 0 },
        { cmd: 'lineTo', x: 1, y: 0 },
        { cmd: 'lineTo', x: 1, y: 1 },
        { cmd: 'close' },
      ]];
      expect(getCustGeomEndpoints(path)).toEqual({ start: null, end: null });
    });

    it('implicit closure (last point == first point) suppresses arrowheads', () => {
      const path: PathCmd[][] = [[
        { cmd: 'moveTo', x: 0.25, y: 0.25 },
        { cmd: 'lineTo', x: 0.75, y: 0.25 },
        { cmd: 'lineTo', x: 0.75, y: 0.75 },
        { cmd: 'lineTo', x: 0.25, y: 0.25 },
      ]];
      const { start, end } = getCustGeomEndpoints(path);
      expect(start).toBeNull();
      expect(end).toBeNull();
    });

    it('a closed FIRST subpath suppresses only the start; an open LAST keeps the end', () => {
      const path: PathCmd[][] = [
        [ // closed triangle
          { cmd: 'moveTo', x: 0, y: 0 },
          { cmd: 'lineTo', x: 0.2, y: 0 },
          { cmd: 'lineTo', x: 0.1, y: 0.2 },
          { cmd: 'close' },
        ],
        [ // open line
          { cmd: 'moveTo', x: 0.5, y: 0.5 },
          { cmd: 'lineTo', x: 1, y: 1 },
        ],
      ];
      const { start, end } = getCustGeomEndpoints(path);
      expect(start).toBeNull();
      expect(end?.x).toBe(1);
      expect(end?.y).toBe(1);
    });
  });
});

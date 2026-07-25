import { describe, expect, it } from 'vitest';
import {
  resolveFloatPlacement,
  type FloatPlacementParticipant,
} from './layout/floats.js';

const table = (
  occurrenceId: string,
  xPt: number,
  widthPt: number,
  paragraphId: number,
  tableOverlap: 'never' | 'overlap',
): FloatPlacementParticipant => ({
  kind: 'table',
  tableOverlap,
  occurrenceId,
  paragraphId,
  bounds: { xPt, yPt: 0, widthPt, heightPt: 1 },
  exclusionBounds: { xPt, yPt: 0, widthPt, heightPt: 1 },
});

describe('typed float placement convergence', () => {
  it('clears every blocker in a finite registry larger than sixteen entries', () => {
    const blockers = Array.from({ length: 17 }, (_, index) =>
      table(`blocker:${index}`, index * 1.1, 1.1, index, 'overlap'));

    const resolved = resolveFloatPlacement({
      moving: table('moving', 0, 1, 100, 'never'),
      blockers,
      avoidance: { kind: 'none' },
      rightBoundaryPt: 100,
      overlapEpsilonPt: 0.01,
      rightBoundarySlackPt: 0.5,
    });

    expect(resolved.bounds.xPt).toBeCloseTo(18.7);
  });

  it('snapshots accessor-backed participant geometry before deterministic placement', () => {
    let left = 0;
    let leftReads = 0;
    let widthReads = 0;
    const blocker = table('blocker', 0, 1.1, 0, 'overlap');
    const chasingBlocker: FloatPlacementParticipant = {
      ...blocker,
      bounds: {
        get xPt() {
          leftReads += 1;
          return left;
        },
        yPt: 0,
        get widthPt() {
          widthReads += 1;
          const width = 1.1;
          left += width;
          return width;
        },
        heightPt: 1,
      },
    };

    expect(resolveFloatPlacement({
      moving: table('moving', 0, 1, 100, 'never'),
      blockers: [chasingBlocker],
      avoidance: { kind: 'none' },
      rightBoundaryPt: 100,
      overlapEpsilonPt: 0.01,
      rightBoundarySlackPt: 0.5,
    }).bounds).toMatchObject({ xPt: 1.1, yPt: 0 });
    expect({ leftReads, widthReads }).toEqual({ leftReads: 1, widthReads: 1 });
  });
});

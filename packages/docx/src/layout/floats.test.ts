import { describe, expect, it } from 'vitest';
import {
  resolveFloatPlacement,
  type FloatPlacementParticipant,
} from './floats.js';

const participant = (
  occurrenceId: string,
  kind: FloatPlacementParticipant['kind'],
  xPt: number,
  paragraphId: number,
  exclusionPaddingPt = 0,
): FloatPlacementParticipant => Object.freeze({
  occurrenceId,
  kind,
  paragraphId,
  bounds: Object.freeze({ xPt, yPt: 0, widthPt: 10, heightPt: 10 }),
  exclusionBounds: Object.freeze({
    xPt: xPt - exclusionPaddingPt,
    yPt: -exclusionPaddingPt,
    widthPt: 10 + exclusionPaddingPt * 2,
    heightPt: 10 + exclusionPaddingPt * 2,
  }),
});

describe('float displacement policy', () => {
  it('applies DrawingML allowOverlap=false only against DrawingML object bounds', () => {
    const moving = participant('moving', 'drawingml', 0, 3, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [
        participant('table', 'table', 10, 0),
        participant('frame', 'frame', 20, 1),
        participant('drawing', 'drawingml', 0, 2),
      ],
      avoidance: { kind: 'drawingml-normative' },
      rightBoundaryPt: 100,
    });

    expect(result.bounds.xPt).toBe(10);
    expect(result.exclusionBounds.xPt).toBe(8);
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('applies tblOverlap=never to raw floating-table extents, not text padding', () => {
    const moving = participant('moving', 'table', 0, 2, 3);
    const blocker = participant('blocker', 'table', 0, 1, 4);
    const result = resolveFloatPlacement({
      moving,
      blockers: [blocker],
      avoidance: { kind: 'floating-table-never' },
      rightBoundaryPt: 100,
    });

    expect(result.bounds.xPt).toBe(10);
    expect(result.exclusionBounds.xPt).toBe(7);
  });

  it('uses the supplied page or cell right boundary before moving down', () => {
    const moving = participant('moving', 'drawingml', 0, 2);
    const blocker = participant('blocker', 'drawingml', 0, 1);
    const result = resolveFloatPlacement({
      moving,
      blockers: [blocker],
      avoidance: { kind: 'drawingml-normative' },
      rightBoundaryPt: 10,
    });

    expect(result.bounds).toMatchObject({ xPt: 0, yPt: 10 });
  });

  it('is independent of blocker input order', () => {
    const moving = participant('moving', 'drawingml', 0, 3);
    const left = participant('left', 'drawingml', 0, 0);
    const right = participant('right', 'drawingml', 10, 1);
    const place = (blockers: readonly FloatPlacementParticipant[]) =>
      resolveFloatPlacement({
        moving,
        blockers,
        avoidance: { kind: 'drawingml-normative' },
        rightBoundaryPt: 30,
      });

    expect(place([left, right])).toEqual(place([right, left]));
    expect(place([left, right]).bounds.xPt).toBe(20);
  });

  it('keeps observed different-paragraph displacement on exclusion bounds', () => {
    const moving = participant('moving', 'table', 0, 2, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [
        participant('same-paragraph', 'drawingml', 0, 2, 8),
        participant('other-paragraph', 'frame', 0, 1, 4),
      ],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.exclusionBounds.xPt).toBe(14);
    expect(result.appliedCompatibilityRuleIds).toEqual([
      'word-float-different-paragraph-displacement',
    ]);
  });

  it('reports a Word compatibility rule only when it changes placement', () => {
    const moving = participant('moving', 'table', 0, 2, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [participant('other-paragraph', 'frame', 50, 1, 4)],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.displacement).toEqual({ xPt: 0, yPt: 0 });
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('preserves authored overlap when no avoidance policy applies', () => {
    const moving = participant('moving', 'drawingml', 0, 1, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [participant('blocker', 'drawingml', 0, 0, 2)],
      avoidance: { kind: 'none' },
      rightBoundaryPt: 100,
    });

    expect(result.bounds).toBe(moving.bounds);
    expect(result.exclusionBounds).toBe(moving.exclusionBounds);
    expect(result.displacement).toEqual({ xPt: 0, yPt: 0 });
  });
});

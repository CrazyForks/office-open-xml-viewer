import { describe, expect, it } from 'vitest';
import {
  applyDrawingMLCollisionRegistryDelta,
  createDrawingMLCollisionRegistry,
  drawingMLCollisionRegistryDelta,
  validateDrawingMLCollisionRegistryDelta,
} from './drawingml-collision-registry.js';

const entry = (occurrenceId: string) => Object.freeze({
  occurrenceId,
  bounds: Object.freeze({ xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 }),
  horizontalOwnership: 'page' as const,
  verticalOwnership: 'page' as const,
});

describe('DrawingML collision registry', () => {
  it('retains accepted object bounds in commit order', () => {
    const initial = createDrawingMLCollisionRegistry(
      'body:page:0',
      'logical-page-points',
    );
    const first = applyDrawingMLCollisionRegistryDelta(
      initial,
      drawingMLCollisionRegistryDelta(initial, [entry('first')]),
    );
    const second = applyDrawingMLCollisionRegistryDelta(
      first,
      drawingMLCollisionRegistryDelta(first, [entry('second')]),
    );

    expect(second.entries.map((candidate) => candidate.occurrenceId))
      .toEqual(['first', 'second']);
  });

  it('rejects stale, cross-flow, and duplicate occurrence deltas', () => {
    const initial = createDrawingMLCollisionRegistry(
      'body:page:0',
      'logical-page-points',
    );
    const committed = applyDrawingMLCollisionRegistryDelta(
      initial,
      drawingMLCollisionRegistryDelta(initial, [entry('first')]),
    );

    expect(() => validateDrawingMLCollisionRegistryDelta(
      committed,
      drawingMLCollisionRegistryDelta(initial, [entry('stale')]),
    )).toThrow('stale');
    expect(() => validateDrawingMLCollisionRegistryDelta(committed, {
      ...drawingMLCollisionRegistryDelta(committed, [entry('other-flow')]),
      flowDomainId: 'body:page:1',
    })).toThrow('flow domain');
    expect(() => validateDrawingMLCollisionRegistryDelta(
      committed,
      drawingMLCollisionRegistryDelta(committed, [entry('first')]),
    )).toThrow('committed twice');
    expect(() => validateDrawingMLCollisionRegistryDelta(
      committed,
      drawingMLCollisionRegistryDelta(committed, [entry('same'), entry('same')]),
    )).toThrow('committed twice');
  });

  it('snapshots and freezes caller-owned entry geometry at the delta boundary', () => {
    const initial = createDrawingMLCollisionRegistry(
      'body:page:0',
      'logical-page-points',
    );
    const mutable = {
      occurrenceId: 'owned',
      bounds: { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 },
      horizontalOwnership: 'host' as const,
      verticalOwnership: 'page' as const,
    };
    const delta = drawingMLCollisionRegistryDelta(initial, [mutable]);

    mutable.occurrenceId = 'changed';
    mutable.bounds.xPt = 99;

    expect(delta.entries[0]).toMatchObject({
      occurrenceId: 'owned',
      bounds: { xPt: 1, yPt: 2, widthPt: 3, heightPt: 4 },
      horizontalOwnership: 'host',
      verticalOwnership: 'page',
    });
    expect(Object.isFrozen(delta.entries[0])).toBe(true);
    expect(Object.isFrozen(delta.entries[0]?.bounds)).toBe(true);
  });
});

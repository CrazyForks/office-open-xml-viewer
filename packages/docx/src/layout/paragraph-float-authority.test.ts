import { describe, expect, it } from 'vitest';
import {
  paragraphAnchorCollisions,
  paragraphWrapExclusions,
} from './paragraph-float-authority.js';
import type { FloatRect } from './float-wrap.js';

const float = (
  overrides: Partial<Extract<FloatRect, { kind: 'shape' }>> = {},
): FloatRect => ({
  kind: 'shape',
  mode: 'square',
  imageKey: '',
  imageX: 12,
  imageY: 24,
  imageW: 30,
  imageH: 40,
  xLeft: 10,
  xRight: 44,
  yTop: 20,
  yBottom: 68,
  side: 'left',
  distLeft: 2,
  distRight: 2,
  distTop: 4,
  distBottom: 4,
  paraId: 0,
  ...overrides,
});

describe('paragraph float authority projection', () => {
  it('keeps wrapNone DrawingML objects in collision authority but not text exclusions', () => {
    const input = float({ anchorOccurrenceId: 'shape:0' });

    expect(paragraphWrapExclusions([input], 'body:0')).toEqual([]);
    expect(paragraphAnchorCollisions([input])).toEqual([{
      occurrenceId: 'shape:0',
      bounds: { xPt: 12, yPt: 24, widthPt: 30, heightPt: 40 },
      horizontalOwnership: 'page',
      verticalOwnership: 'page',
    }]);
  });

  it('projects retained wrap semantics without renderer state', () => {
    const input = float({
      anchorOccurrenceId: 'shape:0',
      authoredWrap: 'tight',
      wrapPolygon: [
        { xPt: 11, yPt: 21 },
        { xPt: 43, yPt: 21 },
        { xPt: 43, yPt: 67 },
      ],
    });

    expect(paragraphWrapExclusions([input], 'body:0')).toEqual([{
      id: 'body:0:float:0',
      wrap: 'tight',
      wrapSide: 'left',
      bounds: { xPt: 10, yPt: 20, widthPt: 34, heightPt: 48 },
      polygon: input.wrapPolygon,
      anchorOccurrenceId: 'shape:0',
      verticalOwnership: 'page',
    }]);
  });
});

import { describe, expect, it } from 'vitest';
import { splitScene3dShapeTransform } from './renderer.js';

describe('DrawingML scene3d transform order', () => {
  it('warps authored flips with the local face and keeps 2-D rotation outside', () => {
    expect(splitScene3dShapeTransform(25, true, false)).toEqual({
      outerRotation: 25,
      localFlipH: true,
      localFlipV: false,
    });
  });
});

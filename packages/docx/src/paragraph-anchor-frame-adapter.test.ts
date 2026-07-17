import { describe, expect, it } from 'vitest';
import { paragraphAnchorReferenceFrames } from './paragraph-anchor-frame-adapter.js';

describe('paragraph anchor frame renderer adapter', () => {
  it('converts mixed renderer units before layout receives reference frames', () => {
    expect(paragraphAnchorReferenceFrames({
      pageIndex: 1,
      scale: 2,
      pageWidth: 600,
      pageH: 1600,
      marginLeft: 40,
      marginRight: 50,
      marginTop: 60,
      marginBottom: 70,
      contentX: 100,
      contentW: 400,
    })).toEqual({
      page: { xPt: 0, yPt: 0, widthPt: 600, heightPt: 800 },
      margin: { xPt: 40, yPt: 60, widthPt: 510, heightPt: 670 },
      column: { xPt: 50, yPt: 60, widthPt: 200, heightPt: 670 },
      pageParity: 'even',
    });
  });

  it('preserves odd-page parity and clamps an inverted margin block extent', () => {
    expect(paragraphAnchorReferenceFrames({
      pageIndex: 2,
      scale: 4,
      pageWidth: 300,
      pageH: 200,
      marginLeft: 20,
      marginRight: 30,
      marginTop: 40,
      marginBottom: 20,
      contentX: 80,
      contentW: 600,
    })).toEqual({
      page: { xPt: 0, yPt: 0, widthPt: 300, heightPt: 50 },
      margin: { xPt: 20, yPt: 40, widthPt: 250, heightPt: 0 },
      column: { xPt: 20, yPt: 40, widthPt: 150, heightPt: 0 },
      pageParity: 'odd',
    });
  });
});

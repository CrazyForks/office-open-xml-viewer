import { describe, expect, it } from 'vitest';
import type { ImageRun } from '../types.js';
import {
  isPageLevelAnchorY,
  isPageLevelWrapFloat,
} from './anchor-classification.js';

describe('page-owned DrawingML anchor classification', () => {
  it('keeps paragraph, line, and character references paragraph-local', () => {
    for (const relativeFrom of ['paragraph', 'line', 'character']) {
      expect(isPageLevelAnchorY(relativeFrom, false)).toBe(false);
    }
    for (const relativeFrom of [
      'page', 'margin', 'topMargin', 'bottomMargin', 'insideMargin',
      'outsideMargin', 'column',
    ]) {
      expect(isPageLevelAnchorY(relativeFrom, false)).toBe(true);
    }
    expect(isPageLevelAnchorY(null, false)).toBe(true);
    expect(isPageLevelAnchorY(undefined, true)).toBe(false);
  });

  it('admits only wrapping anchors whose Y is page-owned', () => {
    const run = (overrides: Partial<ImageRun>): ImageRun => ({
      imagePath: 'word/media/image.png',
      mimeType: 'image/png',
      widthPt: 20,
      heightPt: 10,
      anchor: true,
      wrapMode: 'square',
      anchorYRelativeFrom: 'page',
      anchorYFromPara: false,
      ...overrides,
    });

    expect(isPageLevelWrapFloat(run({}))).toBe(true);
    expect(isPageLevelWrapFloat(run({ wrapMode: 'none' }))).toBe(false);
    expect(isPageLevelWrapFloat(run({
      anchorYRelativeFrom: 'paragraph',
      anchorYFromPara: true,
    }))).toBe(false);
  });
});

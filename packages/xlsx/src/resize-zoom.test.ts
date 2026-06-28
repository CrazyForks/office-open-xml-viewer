import { describe, it, expect } from 'vitest';
import { colWidthToPx, rowHeightToPx, pxToColWidth, pxToRowHeight } from './renderer.js';
import { selectionOverlayStyle } from './viewer.js';

/**
 * Drag-to-resize (issue #567) stores the user's dragged pixel size back into the
 * worksheet's `colWidths` / `rowHeights` model in its native units (Excel column
 * "characters" for columns, points for rows). `pxToColWidth` / `pxToRowHeight`
 * are the exact inverses of the forward converters the renderer uses, so a
 * column dragged to N px renders back at exactly N px with no drift.
 */
describe('px <-> model-unit round trip (drag-to-resize)', () => {
  for (const mdw of [7, 8, 10, 11]) {
    for (const px of [1, 5, 10, 32, 64, 100, 128, 255, 512]) {
      it(`column ${px}px @ mdw=${mdw} round-trips exactly`, () => {
        expect(colWidthToPx(pxToColWidth(px, mdw), mdw)).toBe(px);
      });
    }
  }

  for (const px of [1, 4, 10, 18, 20, 32, 64, 100, 255]) {
    it(`row ${px}px round-trips exactly`, () => {
      expect(rowHeightToPx(pxToRowHeight(px))).toBe(px);
    });
  }
});

/**
 * The viewer takes a single `selectionColor`; the rectangle border uses it as-is
 * and the fill is the same color made translucent (issue follow-up). The default
 * (`#1a73e8`) must keep the historical Google-blue look.
 */
describe('selectionOverlayStyle', () => {
  it('uses the color verbatim for the border', () => {
    expect(selectionOverlayStyle('red').border).toBe('2px solid red');
    expect(selectionOverlayStyle('#1a73e8').border).toBe('2px solid #1a73e8');
  });

  it('derives a translucent fill from the same color', () => {
    expect(selectionOverlayStyle('#1a73e8').background).toBe(
      'color-mix(in srgb, #1a73e8 8%, transparent)',
    );
    expect(selectionOverlayStyle('rgb(0,128,0)').background).toBe(
      'color-mix(in srgb, rgb(0,128,0) 8%, transparent)',
    );
  });
});

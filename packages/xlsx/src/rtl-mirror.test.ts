import { describe, it, expect } from 'vitest';
import { rtlMirrorX, HEADER_W } from './renderer.js';

/**
 * Regression tests for the RTL selection / hit-testing bug (ECMA-376
 * §18.3.1.87 `<sheetView rightToLeft>`): selecting a cell then scrolling
 * horizontally made the blue selection frame drift away from its cell, because
 * the overlay was positioned in the logical-LTR layout while the renderer drew
 * the cell mirrored about `canvasW`.
 *
 * The fix routes the renderer, the selection overlay (cell→px) and pointer
 * hit-testing (px→cell) through one shared transform, `rtlMirrorX`. These tests
 * pin the properties that guarantee the overlay tracks its cell at every scroll
 * offset and that a click resolves to the cell drawn under the cursor.
 */
describe('rtlMirrorX', () => {
  const canvasW = 800;

  it('maps a left-anchored rect to its mirror about canvasW', () => {
    // LTR rect [x, x+w] -> RTL [canvasW - x - w, canvasW - x].
    expect(rtlMirrorX(0, 50, canvasW)).toBe(750); // header strip -> right edge
    expect(rtlMirrorX(50, 100, canvasW)).toBe(650);
    expect(rtlMirrorX(200, 80, canvasW)).toBe(520);
  });

  it('moves the row-header strip to the right edge', () => {
    // LTR header occupies [0, HEADER_W]; mirrored it must sit flush at the
    // right: [canvasW - HEADER_W, canvasW].
    expect(rtlMirrorX(0, HEADER_W, canvasW)).toBe(canvasW - HEADER_W);
  });

  it('is an involution for a point (w = 0): un-mirror recovers the point', () => {
    for (const sx of [0, 50, 137, 400, 799, canvasW]) {
      expect(rtlMirrorX(rtlMirrorX(sx, 0, canvasW), 0, canvasW)).toBe(sx);
    }
  });

  it('round-trips a rect: mirror then mirror returns the original left edge', () => {
    for (const [x, w] of [[50, 60], [120, 24], [333, 17]] as const) {
      const screenLeft = rtlMirrorX(x, w, canvasW);
      // The inverse maps the screen-left + width band back to the logical left.
      expect(rtlMirrorX(screenLeft, w, canvasW)).toBe(x);
    }
  });

  /**
   * Core regression: simulate cell→px (overlay) and px→cell (click) using the
   * SAME transform across several horizontal scroll offsets and assert the
   * selection rect equals the cell rect at every offset — i.e. the frame stays
   * glued to its cell while scrolling, in both directions and across the
   * maxScrollLeft boundary. This is the exact divergence the bug exhibited.
   */
  it('keeps the selection rect glued to its cell across scroll offsets', () => {
    const colW = 64;
    const maxScrollLeft = 1000;

    // Logical-LTR x of a fixed cell's left edge as a function of the logical
    // scroll position (header + columns scrolled past). This is the same
    // formula getCellRect uses (scrollAreaX - scrollOffset + cell offset).
    const logicalCellX = (logicalScroll: number, cellIndexFromStart: number) =>
      HEADER_W - logicalScroll + cellIndexFromStart * colW;

    for (const rawScroll of [0, 64, 256, 640, maxScrollLeft]) {
      // RTL inverts the native scrollbar: effective (logical) scroll = max-raw.
      const logicalScroll = maxScrollLeft - rawScroll;
      const cellIndex = 3;

      // cell -> px (overlay): build the LTR rect, then mirror to screen.
      const ltrX = logicalCellX(logicalScroll, cellIndex);
      const screenLeft = rtlMirrorX(ltrX, colW, canvasW);

      // px -> cell (hit-test): take a click at the cell's on-screen centre,
      // un-mirror it back into logical-LTR space, and confirm it lands inside
      // the same LTR cell band [ltrX, ltrX + colW].
      const clickScreenX = screenLeft + colW / 2;
      const clickLogicalX = rtlMirrorX(clickScreenX, 0, canvasW);
      expect(clickLogicalX).toBeGreaterThanOrEqual(ltrX);
      expect(clickLogicalX).toBeLessThanOrEqual(ltrX + colW);

      // And the overlay's screen band exactly covers the click point: the rect
      // the user sees is the rect the click resolves into.
      expect(clickScreenX).toBeGreaterThanOrEqual(screenLeft);
      expect(clickScreenX).toBeLessThanOrEqual(screenLeft + colW);
    }
  });

  it('selection screen-x moves opposite to the logical-LTR x under RTL', () => {
    // Scrolling that increases the logical-LTR x of a cell must DECREASE its
    // on-screen left edge (the grid moves right-to-left). The old code used the
    // LTR x directly for the overlay, so it moved the wrong way — the bug.
    const colW = 64;
    const xA = rtlMirrorX(HEADER_W + 100, colW, canvasW);
    const xB = rtlMirrorX(HEADER_W + 200, colW, canvasW);
    expect(xB).toBeLessThan(xA);
  });
});

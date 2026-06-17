import { describe, it, expect } from 'vitest';
import { highlightBox } from './highlight-box.js';

// highlightBox is the single source of truth for the vertical extents of an
// OOXML text-highlight (marker) background rectangle, shared by the docx
// (§17.3.2.15) and pptx (§21.1.2.3.4) renderers. The spec fixes no box
// geometry; the band is top = baseline − 0.85·em, height = 1.1·em. These tests
// pin those coefficients so the two renderers can't silently diverge.
describe('highlightBox', () => {
  it('places the top 0.85·em above the baseline', () => {
    const { top } = highlightBox(100, 20);
    expect(top).toBeCloseTo(100 - 20 * 0.85, 6); // 83
  });

  it('makes the band 1.1·em tall', () => {
    const { height } = highlightBox(100, 20);
    expect(height).toBeCloseTo(20 * 1.1, 6); // 22
  });

  it('straddles the baseline so glyph descenders are covered', () => {
    // top is above the baseline and bottom (top + height) is below it.
    const { top, height } = highlightBox(100, 20);
    expect(top).toBeLessThan(100);
    expect(top + height).toBeGreaterThan(100);
  });

  it('scales linearly with the font size', () => {
    const a = highlightBox(0, 10);
    const b = highlightBox(0, 30);
    expect(b.height / a.height).toBeCloseTo(3, 6);
    expect(b.top / a.top).toBeCloseTo(3, 6);
  });
});

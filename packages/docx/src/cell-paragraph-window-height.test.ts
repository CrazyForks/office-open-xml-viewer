import { describe, expect, it } from 'vitest';
import { paintedParagraphHeight } from './renderer.js';

// The live table-cell measurement bridge consumes only the acquired paragraph
// window. These tests pin its cursor replay until that bridge moves under
// layout ownership.
describe('table-cell paragraph-window height', () => {
  type Line = { topY?: number; height: number };
  const heightOf = (line: Line) => line.height;

  it('sums the selected line window from its own origin', () => {
    const lines: Line[] = [{ height: 10 }, { height: 12 }, { height: 8 }, { height: 6 }];

    expect(paintedParagraphHeight(lines, 1, 3, 50, heightOf)).toBe(20);
    expect(paintedParagraphHeight(lines, 0, 2, 100, heightOf)).toBe(22);
  });

  it('includes forward float-clearance jumps without moving backward', () => {
    const lines: Line[] = [
      { height: 10 },
      { topY: 100, height: 12 },
      { topY: 5, height: 8 },
    ];

    expect(paintedParagraphHeight(lines, 0, lines.length, 0, heightOf)).toBe(120);
  });

  it('returns zero for an empty window', () => {
    expect(paintedParagraphHeight([{ height: 10 }], 1, 1, 0, heightOf)).toBe(0);
  });
});

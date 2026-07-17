import { describe, expect, it } from 'vitest';
import { selectParagraphFragment } from './paragraph-pagination.js';
import type { LineLayout, ParagraphLayout } from './types.js';

const paragraph = (): ParagraphLayout => ({
  kind: 'paragraph', id: 'p', source: { story: 'body', storyInstance: 'body', path: [0] },
  flowDomainId: 'body', ordinaryFlow: true,
  flowBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 30 },
  inkBounds: { xPt: 0, yPt: 0, widthPt: 100, heightPt: 30 }, advancePt: 30,
  spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
  lines: Array.from({ length: 3 }, (_, index) => ({
    range: { start: index, end: index + 1 },
    bounds: { xPt: 0, yPt: index * 10, widthPt: 100, heightPt: 10 },
    baselinePt: index * 10 + 8, advancePt: 10, placements: [],
  })) as LineLayout[],
  borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
});

describe('paragraph page-local reserve selection', () => {
  it('admits only the fresh-region extent when an indivisible first line must make progress', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, [
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ], 5, 5, false, { keepLines: false, widowControl: false },
    );

    expect(selected.fragment?.advancePt).toBe(10);
    expect(selected.admittedBlockExtentPt).toBe(5);
  });

  it('selects against the reserve owned by each candidate slice', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, [
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ], 25, 40, false, { keepLines: false, widowControl: false },
      (fragment) => fragment.lines.length >= 2 ? 10 : 0,
    );

    expect(selected.fragment?.lines).toHaveLength(1);
    expect(selected.additionalReservePt).toBe(0);
  });

  it('carries the uniform ruby advance into the exact next source cursor', () => {
    const selected = selectParagraphFragment(
      paragraph(), { boundary: null }, [
        { segIndex: 0, charOffset: 1 },
        { segIndex: 0, charOffset: 2 },
        { segIndex: 0, charOffset: 3 },
      ], 15, 40, false, { keepLines: false, widowControl: false },
      undefined,
      30,
    );

    expect(selected.nextCursor).toEqual({
      boundary: { segIndex: 0, charOffset: 1 },
      sourceRangeStart: 1,
      uniformRubyAdvancePt: 30,
    });
  });
});

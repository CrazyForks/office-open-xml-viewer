import { describe, expect, it } from 'vitest';
import {
  footnoteIdsInRetainedLines,
  footnoteIdsInRetainedSlice,
} from './note-reference-ownership.js';
import type { LineLayout, TableLayout } from './types.js';

describe('retained note reference ownership', () => {
  it('deduplicates only footnote markers painted by the retained line slice', () => {
    const lines = [{
      range: { start: 0, end: 1 }, baselinePt: 8, advancePt: 10,
      placements: [
        { kind: 'text', range: { start: 0, end: 1 }, xPt: 0, advancePt: 1,
          noteReference: { kind: 'footnote', id: '4' } },
        { kind: 'text', range: { start: 0, end: 1 }, xPt: 1, advancePt: 1,
          noteReference: { kind: 'footnote', id: '4' } },
      ],
    }] as unknown as LineLayout[];
    expect(footnoteIdsInRetainedLines(lines)).toEqual(['4']);
  });

  it('collects only references retained by the accepted table slice, including nested tables', () => {
    const paragraph = (id: string, noteId: string) => ({
      kind: 'paragraph', id,
      lines: [{
        placements: [{ kind: 'text', noteReference: { kind: 'footnote', id: noteId } }],
      }],
    });
    const nested = {
      kind: 'table', rows: [{ cells: [{ blocks: [{ layout: paragraph('nested-note', 'nested') }] }] }],
    };
    const retainedSlice = {
      kind: 'table',
      rows: [{ cells: [{ blocks: [
        { layout: paragraph('direct-note', 'direct') },
        { layout: nested },
      ] }] }],
    } as unknown as TableLayout;

    expect(footnoteIdsInRetainedSlice(retainedSlice)).toEqual(['direct', 'nested']);
  });
});

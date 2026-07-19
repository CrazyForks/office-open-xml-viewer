import { describe, expect, it } from 'vitest';
import {
  endnoteIdsInRetainedSlice,
  footnoteIdsInRetainedLines,
  footnoteIdsInRetainedSlice,
  noteReferenceIdsInDocumentOrder,
} from './note-reference-ownership.js';
import type { BodyElement } from '../types.js';
import type { LineLayout, ParagraphLayout, TableLayout } from './types.js';

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

  it('collects endnotes independently from footnotes', () => {
    const retained = {
      kind: 'paragraph',
      lines: [{
        placements: [
          { kind: 'text', noteReference: { kind: 'footnote', id: 'foot' } },
          { kind: 'text', noteReference: { kind: 'endnote', id: 'end' } },
        ],
      }],
    } as unknown as ParagraphLayout;

    expect(endnoteIdsInRetainedSlice(retained)).toEqual(['end']);
  });

  it('numbers only first references in main-story document order', () => {
    const paragraph = (kind: 'footnote' | 'endnote', id: string) => ({
      type: 'paragraph',
      runs: [{ type: 'text', noteRef: { kind, id } }],
    });
    const body = [
      paragraph('endnote', 'second-part-entry'),
      {
        type: 'table',
        rows: [{ cells: [{ content: [
          paragraph('endnote', 'first-part-entry'),
          paragraph('endnote', 'second-part-entry'),
        ] }] }],
      },
    ] as unknown as BodyElement[];

    expect(noteReferenceIdsInDocumentOrder(body, 'endnote'))
      .toEqual(['second-part-entry', 'first-part-entry']);
  });
});

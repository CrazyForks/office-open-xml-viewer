import { describe, expect, it } from 'vitest';
import { paragraphAcquisitionInput } from './parser-model.js';
import type { DocParagraph } from './types.js';

const paragraph = (path: number[]) => ({
  model: {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [],
    runs: [{
      type: 'text', text: 'cached result', fontSize: 10,
      bold: false, italic: false, underline: false, strikethrough: false,
      color: null, fontFamily: null, isLink: true, background: null,
      vertAlign: null, hyperlink: null, hyperlinkAnchor: 'DestinationBookmark',
    }],
    __complexFieldBoundaries: [{
      occurrenceId: 3,
      boundary: 'start',
      runIndex: 0,
      fieldType: 'ref',
      instruction: 'REF DestinationBookmark \\h',
      hyperlinkAnchor: 'DestinationBookmark',
    }],
  } as unknown as DocParagraph,
  source: { story: 'body' as const, storyInstance: 'body', path },
});

describe('complex-field parser-model projection', () => {
  it('replaces parser-private boundary wire with immutable layout input', () => {
    const fixture = paragraph([4]);
    const input = paragraphAcquisitionInput(fixture.model, fixture.source) as unknown as {
      __complexFieldBoundaries?: unknown;
      complexFieldBoundaries?: readonly {
        occurrenceKey: string;
        boundary: string;
        runIndex: number;
        fieldType: string;
        instruction: string;
        hyperlinkAnchor?: string;
      }[];
    };

    expect(input.__complexFieldBoundaries).toBeUndefined();
    expect(input.complexFieldBoundaries).toEqual([{
      occurrenceKey: 'complex-field:body:body::3',
      boundary: 'start',
      runIndex: 0,
      fieldType: 'ref',
      instruction: 'REF DestinationBookmark \\h',
      hyperlinkAnchor: 'DestinationBookmark',
    }]);
    expect(structuredClone(input.complexFieldBoundaries)).toEqual(input.complexFieldBoundaries);
    expect(Object.isFrozen(input.complexFieldBoundaries)).toBe(true);
    expect(Object.isFrozen(input.complexFieldBoundaries?.[0])).toBe(true);
  });

  it('keeps one occurrence identity across paragraphs but scopes nested containers', () => {
    const bodyA = paragraph([4]);
    const bodyB = paragraph([5]);
    const cell = paragraph([2, 0, 1, 4]);
    const cellNextParagraph = paragraph([2, 0, 1, 5]);

    const key = (fixture: ReturnType<typeof paragraph>) => (
      paragraphAcquisitionInput(fixture.model, fixture.source) as unknown as {
        complexFieldBoundaries: readonly { occurrenceKey: string }[];
      }
    ).complexFieldBoundaries[0].occurrenceKey;

    expect(key(bodyA)).toBe(key(bodyB));
    expect(key(cell)).toBe(key(cellNextParagraph));
    expect(key(cell)).not.toBe(key(bodyA));
  });
});

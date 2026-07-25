import { describe, expect, it } from 'vitest';
import { textBoxContentAcquisitionInput } from './parser-model.js';
import type { ShapeRun } from './types.js';

describe('private complete text-box parser wire', () => {
  it('projects the ordered block wire into an immutable plain-data snapshot', () => {
    const markerPath = [2];
    const parserWire = [
      { type: 'paragraph', runs: [{ type: 'text', text: 'Before' }] },
      {
        type: 'table',
        rows: [{
          cells: [{
            content: [{ type: 'paragraph', runs: [{ type: 'text', text: 'Nested' }] }],
          }],
        }],
      },
      {
        type: 'unsupportedTextBoxBlock',
        qName: 'w:altChunk',
        sourcePath: markerPath,
      },
    ];
    const parserShape = { textBoxContent: parserWire } as unknown as ShapeRun;

    const input = textBoxContentAcquisitionInput(parserShape);

    expect(input).toEqual(parserWire);
    expect(input).not.toBe(parserWire);
    expect(Object.isFrozen(input)).toBe(true);
    expect(Object.isFrozen(input?.[1])).toBe(true);
    const marker = input?.[2];
    expect(Object.isFrozen(marker)).toBe(true);
    if (marker?.type !== 'unsupportedTextBoxBlock') {
      throw new Error('expected the third block to retain its unsupported marker');
    }
    expect(Object.isFrozen(marker.sourcePath)).toBe(true);

    markerPath.push(9);
    expect(marker.sourcePath).toEqual([2]);
    expect(JSON.parse(JSON.stringify(input))).toEqual(input);
  });

  it('keeps ordinary public ShapeRun values free of parser-only content', () => {
    const publicShape = { type: 'shape', textBlocks: [] } as unknown as ShapeRun;

    expect(textBoxContentAcquisitionInput(publicShape)).toBeUndefined();
  });
});

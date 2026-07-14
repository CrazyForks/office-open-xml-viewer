import { describe, expect, it } from 'vitest';
import { docxRenderedTextUsages } from './document-content.js';
import type { InternalFieldRun } from './parser-model.js';
import type { DocxDocumentModel } from './types.js';

describe('docx rendered text inventory', () => {
  it('inventories field results on both non-CS/EA and CS formatting tuples', () => {
    const field: InternalFieldRun & { type: 'field' } = {
      type: 'field', fieldType: 'other', instruction: 'REF x', fallbackText: 'result',
      bold: true, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Latin Face', background: null,
      vertAlign: null, fontFamilyEastAsia: 'EA Face', fontFamilyCs: 'CS Face',
      boldCs: false, italicCs: true,
    };
    const doc = {
      body: [{ type: 'paragraph', runs: [field] }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;

    expect([...docxRenderedTextUsages(doc)].filter((usage) => usage.text === 'result')).toEqual([
      { text: 'result', fontFamilies: ['Latin Face', 'EA Face'], bold: true, italic: false },
      { text: 'result', fontFamilies: ['CS Face'], bold: false, italic: true },
    ]);
  });
});

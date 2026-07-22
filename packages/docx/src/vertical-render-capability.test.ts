import { describe, expect, it } from 'vitest';
import type { DocxDocumentModel } from './types.js';
import { documentRequiresDomVerticalGlyphLayout } from './vertical-render-capability.js';

function model(): DocxDocumentModel {
  return {
    section: { textDirection: null },
    body: [],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

describe('vertical render capability', () => {
  it.each(['tbRl', 'tbRlV', 'tbLrV'])('routes a %s section through the DOM glyph path', (textDirection) => {
    const document = model();
    document.section.textDirection = textDirection;
    expect(documentRequiresDomVerticalGlyphLayout(document)).toBe(true);
  });

  it('finds vertical section transitions', () => {
    const sectioned = model();
    sectioned.body = [{ type: 'sectionBreak', kind: 'nextPage', textDirection: 'tbRl' }];
    expect(documentRequiresDomVerticalGlyphLayout(sectioned)).toBe(true);
  });

  it('keeps horizontal, all-rotated, and text-box-only vertical text on the OffscreenCanvas path', () => {
    const document = model() as unknown as Record<string, unknown>;
    document.section = { textDirection: 'btLr' };
    document.body = [
      { textVert: 'vert' },
      { textVert: 'vert270' },
      { textVert: 'eaVert' },
      { textVert: 'mongolianVert' },
    ];
    expect(documentRequiresDomVerticalGlyphLayout(document as unknown as DocxDocumentModel)).toBe(false);
  });
});

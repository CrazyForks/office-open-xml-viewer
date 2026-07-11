import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ECMA-376 §17.18.44 (ST_Jc `both`) — the Knuth-Plass space-shrink fit tolerance
// (`SPACE_SHRINK_RATIO`) must NOT admit an extra word onto a FULLY-JUSTIFIED line.
//
// Word's line breaker is greedy and identical for justified and non-justified
// paragraphs; justification redistributes the residual slack by EXPANDING the
// inter-word spaces, and (in its default mode) never COMPRESSES a line below its
// natural width to admit one more word. So on a justified line a candidate word
// whose natural advance overflows the column must wrap — even if the overflow is
// smaller than the line's total inter-word shrink budget. The shrink tolerance
// exists only to absorb the Chromium-`measureText` vs Word advance-width bias on
// a line that will be drawn at (or compressed toward) its natural spacing, i.e. a
// NON-justified line (e.g. a substituted-font centred title that Word keeps on one
// row). See issue #698: in sample-15 p1's narrow (8.4 cm) justified copyright
// column the tolerance pulled "citation" up onto the "…and the full" line, where
// Word (PDF ground truth) breaks after "full".

const FONT_PX = 12; // linear stub: each code point advances FONT_PX at scale 1

/** Linear recording canvas: measureText advances FONT_PX per code point (space
 *  included), so a token's trailing-space width is exactly one FONT_PX. This makes
 *  the fit arithmetic explicit: a word overflowing the column by < the old
 *  `SPACE_SHRINK_RATIO · Σspace` budget was silently admitted before the fix. */
function makeLinearCanvas(): HTMLCanvasElement {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => Number(/([\d.]+)px/.exec(font)?.[1] ?? FONT_PX);
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    fontKerning: 'auto',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return canvas as unknown as HTMLCanvasElement;
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'serif', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(text: string, alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocRun],
    defaultFontSize: FONT_PX, defaultFontFamily: 'serif',
    widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(pageWidth: number): SectionProps {
  return {
    pageWidth, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    docGridCharSpace: undefined,
  } as SectionProps;
}

function doc(el: BodyElement, sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body: [el],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

/** Render one paragraph and return the number of visual text lines (distinct
 *  baseline y values reported by onTextRun). */
async function lineCount(text: string, alignment: DocParagraph['alignment'], pageWidth: number): Promise<number> {
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(para(text, alignment), section(pageWidth)), makeLinearCanvas(), 0, {
    dpr: 1,
    width: pageWidth, // scale = 1 px/pt
    onTextRun: (r) => { if (r.text && r.text.trim()) runs.push(r); },
  });
  return new Set(runs.map((r) => Math.round(r.y))).size;
}

// Four 4-glyph words: natural width = 4·(4·12) + 3·(12) = 192 + 36 = 228 px.
// Column = 225 px ⇒ the 4th word overflows by 3 px. The line carries 3 trailing
// spaces (Σspace = 36 px), so the OLD budget = 0.25·36 = 9 px ≥ 3 px admitted it.
const TEXT = 'AAAA AAAA AAAA AAAA';
const COLUMN = 225;

describe('§17.18.44 — space-shrink fit tolerance is suppressed on justified lines (issue #698)', () => {
  it('wraps a justified paragraph whose last word overflows the column at natural width', async () => {
    // Word breaks here (the overflow is a genuine word, not measurement bias);
    // the shrink budget must not pull it up. RED before the fix: 1 line.
    expect(await lineCount(TEXT, 'both', COLUMN)).toBe(2);
  });

  it('keeps the SAME overflow on ONE line for a non-justified paragraph (bias tolerance retained)', async () => {
    // A left-aligned paragraph will be drawn at (or compressed toward) natural
    // spacing, so the small overflow is absorbed as measurement bias — this is the
    // behaviour that keeps sample-10 p1's centred title on a single row.
    expect(await lineCount(TEXT, 'left', COLUMN)).toBe(1);
    expect(await lineCount(TEXT, 'center', COLUMN)).toBe(1);
  });

  it('does not force a wrap when the justified content genuinely fits at natural width', async () => {
    // Widen the column past the natural width: no wrap, no spurious break.
    expect(await lineCount(TEXT, 'both', 240)).toBe(1);
  });
});

import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ECMA-376 §17.3.2.33 `w:smallCaps`: small caps are sized PER CHARACTER. An
// originally-lowercase letter is drawn as a reduced-size capital; an
// originally-uppercase letter (and any non-cased character) is drawn at the FULL
// run size. So "Introduction" → "I" at full size + "NTRODUCTION" reduced —
// matching the leading "1." of a heading number (which is full size). A small
// caps word must never split between its full-cap initial and reduced remainder
// when it wraps.

interface Call { text: string; x: number; y: number; px: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y, px: px() }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y, px: px() }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function para(text: string, opts: { smallCaps?: boolean; allCaps?: boolean; size?: number } = {}): DocParagraph {
  const size = opts.size ?? 10;
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: size, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
      smallCaps: opts.smallCaps ?? false, allCaps: opts.allCaps ?? false,
    } as DocParagraph['runs'][number]],
    defaultFontSize: size, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageWidth = 400): DocxDocumentModel {
  const section = {
    pageWidth, pageHeight: 600,
    marginTop: 5, marginRight: 5, marginBottom: 5, marginLeft: 5,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(body: BodyElement[], pageWidth = 400): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(body, pageWidth), canvas, 0, { dpr: 1, width: pageWidth });
  return calls;
}

describe('small caps per-character sizing (§17.3.2.33)', () => {
  it('keeps the originally-UPPERCASE initial at full size and reduces the lowercase remainder', async () => {
    const calls = await render([para('Introduction', { smallCaps: true }) as unknown as BodyElement]);
    const init = calls.find((c) => c.text === 'I');
    const rest = calls.find((c) => c.text === 'NTRODUCTION');
    expect(init, 'full-size initial "I"').toBeDefined();
    expect(rest, 'reduced "NTRODUCTION"').toBeDefined();
    // §17.3.2.33: reduced size is TWO POINTS SMALLER (10 − 2 = 8), not a ratio.
    expect(init!.px).toBeCloseTo(10, 3);
    expect(rest!.px).toBeCloseTo(8, 3);
    // Both on the same line, in reading order.
    expect(init!.y).toBeCloseTo(rest!.y, 3);
    expect(rest!.x).toBeGreaterThan(init!.x);
  });

  it('reduces by exactly 2 points (subtractive), not a ratio, at larger sizes', async () => {
    // A 20pt heading: caps stay 20pt, small letters are 18pt (20 − 2), NOT 16 (0.8×).
    const calls = await render([para('Introduction', { smallCaps: true, size: 20 }) as unknown as BodyElement]);
    expect(calls.find((c) => c.text === 'I')!.px).toBeCloseTo(20, 3);
    expect(calls.find((c) => c.text === 'NTRODUCTION')!.px).toBeCloseTo(18, 3);
  });

  it('does NOT reduce non-alphabetic characters (digits stay full size)', async () => {
    // §17.3.2.33 affects "small letter characters" only — the "2" of "co2" is full.
    const calls = await render([para('co2', { smallCaps: true }) as unknown as BodyElement]);
    expect(calls.find((c) => c.text === 'CO')!.px, '"co" reduced').toBeCloseTo(8, 3);
    expect(calls.find((c) => c.text === '2')!.px, '"2" full size').toBeCloseTo(10, 3);
  });

  it('allCaps uppercases but does NOT reduce any character', async () => {
    const calls = await render([para('Introduction', { allCaps: true }) as unknown as BodyElement]);
    // allCaps emits a single uppercased piece, all at the full size.
    const piece = calls.find((c) => c.text.includes('INTRODUCTION'));
    expect(piece).toBeDefined();
    expect(piece!.px).toBeCloseTo(10, 3);
    // No reduced (8pt) glyphs at all.
    expect(calls.some((c) => Math.abs(c.px - 8) < 0.01)).toBe(false);
  });

  it('sizes the line box from the FULL run size even when a line is all reduced', async () => {
    // An all-lowercase small-caps word → every piece reduced. Its line box must
    // still be the full-size line height (§17.3.2.33 reduces glyphs, not leading),
    // so a following paragraph sits exactly where it does for full-size text.
    const small = await render([
      para('abcdef', { smallCaps: true }) as unknown as BodyElement,
      para('X') as unknown as BodyElement,
    ]);
    const full = await render([
      para('abcdef') as unknown as BodyElement,
      para('X') as unknown as BodyElement,
    ]);
    const xSmall = small.find((c) => c.text === 'X')!;
    const xFull = full.find((c) => c.text === 'X')!;
    expect(xSmall.y).toBeCloseTo(xFull.y, 1);
  });

  it('does not split a small-caps word between its case pieces when it wraps', async () => {
    // Narrow page → the small-caps line wraps mid-paragraph.
    const calls = await render(
      [para('Introduction Subject Methods Conclusion', { smallCaps: true }) as unknown as BodyElement],
      120,
    );
    // Group each painted piece by word: a full-cap single-letter initial must
    // sit on the SAME line (y) as the reduced remainder that immediately follows.
    for (let i = 0; i < calls.length - 1; i++) {
      const a = calls[i];
      const b = calls[i + 1];
      const isInitial = a.text.length === 1 && /[A-Z]/.test(a.text) && b.x > a.x;
      if (isInitial) {
        expect(a.y, `"${a.text}" + "${b.text}" must stay on one line`).toBeCloseTo(b.y, 3);
      }
    }
  });
});
